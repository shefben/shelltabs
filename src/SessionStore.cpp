#include "SessionStore.h"

#include "StringUtils.h"

#include "ColorSerialization.h"
#include "Logging.h"

#include <Shlwapi.h>

#include <algorithm>
#include <atomic>
#include <cerrno>
#include <cwctype>
#include <cwchar>
#include <string>
#include <string_view>

#include "Utilities.h"

namespace shelltabs {
namespace {
constexpr wchar_t kStorageFile[] = L"session.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kGroupToken[] = L"group";
constexpr wchar_t kTabToken[] = L"tab";
constexpr wchar_t kSelectedToken[] = L"selected";
constexpr wchar_t kSequenceToken[] = L"sequence";
constexpr wchar_t kDockToken[] = L"dock";
constexpr wchar_t kCommentChar = L'#';
constexpr wchar_t kCrashMarkerFile[] = L"session.lock";
constexpr wchar_t kIntegrityToken[] = L"integrity";

enum class IntegrityStatus {
    kNotPresent,
    kValid,
    kMalformedHeader,
    kEncodingError,
    kLengthMismatch,
    kChecksumMismatch,
};

struct IntegrityCheckResult {
    IntegrityStatus status = IntegrityStatus::kNotPresent;
    std::wstring_view payload;
    uint64_t expectedChecksum = 0;
    uint64_t actualChecksum = 0;
    uint64_t expectedLength = 0;
    size_t actualLength = 0;
};

constexpr uint64_t kFnvOffsetBasis = 1469598103934665603ull;
constexpr uint64_t kFnvPrime = 1099511628211ull;

uint64_t ComputeChecksum(std::string_view data) {
    uint64_t hash = kFnvOffsetBasis;
    for (unsigned char value : data) {
        hash ^= value;
        hash *= kFnvPrime;
    }
    return hash;
}

std::wstring BuildSiblingPath(const std::wstring& path, const wchar_t* extension) {
    std::wstring sibling = path;
    sibling += extension;
    return sibling;
}

void EnsureParentDirectory(const std::wstring& path) {
    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring::npos) {
        return;
    }
    std::wstring directory = path.substr(0, separator);
    if (directory.empty()) {
        return;
    }
    CreateDirectoryW(directory.c_str(), nullptr);
}

IntegrityCheckResult ValidateIntegrity(const std::wstring& content) {
    IntegrityCheckResult result;
    result.payload = std::wstring_view(content);
    if (content.empty()) {
        return result;
    }

    size_t lineEnd = content.find(L'\n');
    std::wstring firstLine = content.substr(0, lineEnd == std::wstring::npos ? std::wstring::npos : lineEnd);
    firstLine = Trim(firstLine);
    if (firstLine.rfind(kIntegrityToken, 0) != 0) {
        return result;
    }

    auto tokens = Split(firstLine, L'|');
    for (auto& token : tokens) {
        token = Trim(token);
    }
    if (tokens.size() < 3) {
        result.status = IntegrityStatus::kMalformedHeader;
        return result;
    }

    wchar_t* checksumEnd = nullptr;
    wchar_t* lengthEnd = nullptr;
    errno = 0;
    unsigned long long checksum = wcstoull(tokens[1].c_str(), &checksumEnd, 16);
    if (errno != 0 || checksumEnd == nullptr || *checksumEnd != L'\0') {
        result.status = IntegrityStatus::kMalformedHeader;
        return result;
    }

    errno = 0;
    unsigned long long length = wcstoull(tokens[2].c_str(), &lengthEnd, 10);
    if (errno != 0 || lengthEnd == nullptr || *lengthEnd != L'\0') {
        result.status = IntegrityStatus::kMalformedHeader;
        return result;
    }

    const size_t payloadOffset = lineEnd == std::wstring::npos ? content.size() : lineEnd + 1;
    std::wstring_view payloadView(content.data() + payloadOffset, content.size() - payloadOffset);
    const std::string payloadUtf8 = WideToUtf8(payloadView);
    if (!payloadView.empty() && payloadUtf8.empty()) {
        result.status = IntegrityStatus::kEncodingError;
        result.expectedChecksum = checksum;
        result.expectedLength = length;
        return result;
    }

    result.payload = payloadView;
    result.expectedChecksum = checksum;
    result.expectedLength = length;
    result.actualLength = payloadUtf8.size();
    result.actualChecksum = ComputeChecksum(payloadUtf8);

    if (result.actualLength != result.expectedLength) {
        result.status = IntegrityStatus::kLengthMismatch;
        return result;
    }
    if (result.actualChecksum != result.expectedChecksum) {
        result.status = IntegrityStatus::kChecksumMismatch;
        return result;
    }

    result.status = IntegrityStatus::kValid;
    return result;
}

enum class LoadAttemptStatus {
    kSuccess,
    kMissing,
    kCorrupt,
    kError,
};

bool ParseSessionContent(const std::wstring& content, SessionData* output) {
    if (!output) {
        return false;
    }

    SessionData parsed;
    bool versionSeen = false;
    int version = 1;
    SessionGroup* currentGroup = nullptr;

    const bool parsedOk =
        ParseConfigLines(content, kCommentChar, L'|', [&](const std::vector<std::wstring>& tokens) {
            if (tokens.empty()) {
                return true;
            }

            const std::wstring& header = tokens.front();
            if (header == kVersionToken) {
                if (tokens.size() < 2) {
                    return false;
                }
                version = std::max(1, _wtoi(tokens[1].c_str()));
                if (version > 6) {
                    return false;
                }
                versionSeen = true;
                return true;
            }

            if (header == kSelectedToken) {
                if (tokens.size() >= 3) {
                    parsed.selectedGroup = _wtoi(tokens[1].c_str());
                    parsed.selectedTab = _wtoi(tokens[2].c_str());
                }
                return true;
            }

            if (header == kSequenceToken) {
                if (tokens.size() >= 2) {
                    parsed.groupSequence = std::max(1, _wtoi(tokens[1].c_str()));
                }
                return true;
            }

            if (header == kDockToken) {
                if (tokens.size() >= 2) {
                    parsed.dockMode = ParseDockMode(tokens[1]);
                }
                return true;
            }

            if (header == kGroupToken) {
                if (tokens.size() < 3) {
                    return true;
                }
                SessionGroup group;
                group.name = tokens[1];
                group.collapsed = ParseBool(tokens[2]);
                size_t index = 3;
                if (version <= 2) {
                    if (tokens.size() > index) {
                        ++index;
                    }
                    if (tokens.size() > index) {
                        ++index;
                    }
                    if (tokens.size() > index) {
                        ++index;
                    }
                }
                if (version >= 2) {
                    if (tokens.size() > index) {
                        group.headerVisible = ParseBool(tokens[index]);
                        ++index;
                    }
                    if (tokens.size() > index) {
                        group.hasOutline = ParseBool(tokens[index]);
                        ++index;
                    }
                    if (tokens.size() > index) {
                        group.outlineColor = ParseColor(tokens[index], group.outlineColor);
                        ++index;
                    }
                    if (version >= 4 && tokens.size() > index) {
                        group.outlineStyle = ParseOutlineStyle(tokens[index], group.outlineStyle);
                        ++index;
                    }
                    if (tokens.size() > index) {
                        group.savedGroupId = tokens[index];
                        ++index;
                    }
                }
                parsed.groups.emplace_back(std::move(group));
                currentGroup = &parsed.groups.back();
                return true;
            }

            if (header == kTabToken) {
                if (!currentGroup || tokens.size() < 5) {
                    return true;
                }
                SessionTab tab;
                tab.name = tokens[1];
                tab.tooltip = tokens[2];
                tab.hidden = ParseBool(tokens[3]);
                tab.path = tokens[4];
                size_t index = 5;
                if (version >= 5) {
                    if (tokens.size() > index) {
                        tab.lastActivatedTick = _wcstoui64(tokens[index].c_str(), nullptr, 10);
                        ++index;
                    }
                    if (tokens.size() > index) {
                        tab.activationOrdinal = _wcstoui64(tokens[index].c_str(), nullptr, 10);
                        ++index;
                    }
                }
                currentGroup->tabs.emplace_back(std::move(tab));
                return true;
            }

            return true;
        });

    if (!parsedOk) {
        return false;
    }

    if (!versionSeen || parsed.groups.empty()) {
        return false;
    }

    *output = std::move(parsed);
    return true;
}

void CleanupCheckpointArtifacts(bool deleteCheckpoints) {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }

    const std::wstring tmpPattern = directory + L"session*.tmp";
    WIN32_FIND_DATAW findData{};
    HANDLE findHandle = FindFirstFileW(tmpPattern.c_str(), &findData);
    if (findHandle != INVALID_HANDLE_VALUE) {
        do {
            if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
                std::wstring filePath = directory + findData.cFileName;
                if (!DeleteFileW(filePath.c_str())) {
                    const DWORD error = GetLastError();
                    if (error != ERROR_FILE_NOT_FOUND) {
                        LogMessage(LogLevel::Warning,
                                   L"SessionStore::CleanupCheckpointArtifacts failed to delete %ls (error=%lu)",
                                   filePath.c_str(), error);
                    }
                }
            }
        } while (FindNextFileW(findHandle, &findData));
        FindClose(findHandle);
    }

    if (!deleteCheckpoints) {
        return;
    }

    const std::wstring previousPattern = directory + L"session*.previous";
    findHandle = FindFirstFileW(previousPattern.c_str(), &findData);
    if (findHandle == INVALID_HANDLE_VALUE) {
        return;
    }

    do {
        if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            continue;
        }
        std::wstring filePath = directory + findData.cFileName;
        if (!DeleteFileW(filePath.c_str())) {
            const DWORD error = GetLastError();
            if (error != ERROR_FILE_NOT_FOUND) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore::CleanupCheckpointArtifacts failed to delete checkpoint %ls (error=%lu)",
                           filePath.c_str(), error);
            }
        }
    } while (FindNextFileW(findHandle, &findData));
    FindClose(findHandle);
}

LoadAttemptStatus AttemptLoadPath(const std::wstring& path, bool createIfMissing, SessionData* output) {
    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(path, &content, &fileExists)) {
        const DWORD error = GetLastError();
        LogMessage(LogLevel::Error, L"SessionStore::Load unable to read %ls (error=%lu)", path.c_str(), error);
        return LoadAttemptStatus::kError;
    }

    if (!fileExists) {
        if (createIfMissing) {
            EnsureParentDirectory(path);
            HANDLE created = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW,
                                         FILE_ATTRIBUTE_NORMAL, nullptr);
            if (created != INVALID_HANDLE_VALUE) {
                CloseHandle(created);
            } else {
                const DWORD error = GetLastError();
                LogMessage(LogLevel::Warning, L"SessionStore::Load failed to create %ls (error=%lu)", path.c_str(),
                           error);
            }
            if (output) {
                *output = SessionData{};
            }
            return LoadAttemptStatus::kSuccess;
        }
        return LoadAttemptStatus::kMissing;
    }

    if (content.empty()) {
        if (output) {
            *output = SessionData{};
        }
        return LoadAttemptStatus::kSuccess;
    }

    IntegrityCheckResult integrity = ValidateIntegrity(content);
    if (integrity.status != IntegrityStatus::kNotPresent && integrity.status != IntegrityStatus::kValid) {
        switch (integrity.status) {
            case IntegrityStatus::kMalformedHeader:
                LogMessage(LogLevel::Warning, L"SessionStore::Load integrity header malformed in %ls", path.c_str());
                break;
            case IntegrityStatus::kEncodingError:
                LogMessage(LogLevel::Warning,
                           L"SessionStore::Load integrity payload encoding error in %ls (expected bytes=%llu)",
                           path.c_str(), static_cast<unsigned long long>(integrity.expectedLength));
                break;
            case IntegrityStatus::kLengthMismatch:
                LogMessage(LogLevel::Warning,
                           L"SessionStore::Load integrity length mismatch in %ls (expected=%llu actual=%zu)",
                           path.c_str(), static_cast<unsigned long long>(integrity.expectedLength),
                           integrity.actualLength);
                break;
            case IntegrityStatus::kChecksumMismatch:
                LogMessage(LogLevel::Warning,
                           L"SessionStore::Load integrity checksum mismatch in %ls (expected=%016llX actual=%016llX)",
                           path.c_str(), static_cast<unsigned long long>(integrity.expectedChecksum),
                           static_cast<unsigned long long>(integrity.actualChecksum));
                break;
            case IntegrityStatus::kNotPresent:
            case IntegrityStatus::kValid:
                break;
        }
        return LoadAttemptStatus::kCorrupt;
    }

    std::wstring buffer;
    const std::wstring* parseTarget = &content;
    if (integrity.status == IntegrityStatus::kValid) {
        buffer.assign(integrity.payload.begin(), integrity.payload.end());
        parseTarget = &buffer;
    }

    SessionData parsed;
    if (!ParseSessionContent(*parseTarget, &parsed)) {
        LogMessage(LogLevel::Warning, L"SessionStore::Load failed to parse %ls", path.c_str());
        return LoadAttemptStatus::kCorrupt;
    }

    if (output) {
        *output = std::move(parsed);
    }
    return LoadAttemptStatus::kSuccess;
}

bool WriteSessionFileAtomically(const std::wstring& path, std::wstring_view contents) {
    const std::string utf8 = WideToUtf8(contents);
    if (!contents.empty() && utf8.empty()) {
        LogMessage(LogLevel::Error, L"SessionStore::Save failed to encode contents for %ls", path.c_str());
        return false;
    }

    const std::wstring tempPath = BuildSiblingPath(path, L".tmp");
    const std::wstring previousPath = BuildSiblingPath(path, L".previous");

    DeleteFileW(tempPath.c_str());

    if (!DeleteFileW(previousPath.c_str())) {
        const DWORD deleteError = GetLastError();
        if (deleteError != ERROR_FILE_NOT_FOUND) {
            LogMessage(LogLevel::Warning, L"SessionStore::Save failed to remove previous checkpoint %ls (error=%lu)",
                       previousPath.c_str(), deleteError);
        }
    }

    if (GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES) {
        if (!MoveFileExW(path.c_str(), previousPath.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
            const DWORD rotateError = GetLastError();
            LogMessage(LogLevel::Error, L"SessionStore::Save failed to rotate %ls to %ls (error=%lu)", path.c_str(),
                       previousPath.c_str(), rotateError);
            return false;
        }
    }

    HANDLE file = CreateFileW(tempPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        const DWORD createError = GetLastError();
        LogMessage(LogLevel::Error, L"SessionStore::Save failed to create temporary file %ls (error=%lu)",
                   tempPath.c_str(), createError);
        MoveFileExW(previousPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
        return false;
    }

    DWORD bytesWritten = 0;
    BOOL writeOk = TRUE;
    if (!utf8.empty()) {
        writeOk = WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &bytesWritten, nullptr);
        if (writeOk && bytesWritten != utf8.size()) {
            writeOk = FALSE;
            SetLastError(ERROR_WRITE_FAULT);
        }
    }

    BOOL flushOk = writeOk ? FlushFileBuffers(file) : FALSE;
    const DWORD writeError = writeOk ? ERROR_SUCCESS : GetLastError();
    const DWORD flushError = flushOk ? ERROR_SUCCESS : GetLastError();
    CloseHandle(file);

    if (!writeOk || !flushOk) {
        if (!writeOk) {
            LogMessage(LogLevel::Error, L"SessionStore::Save failed to write %ls (error=%lu)", tempPath.c_str(), writeError);
        } else {
            LogMessage(LogLevel::Error, L"SessionStore::Save failed to flush %ls (error=%lu)", tempPath.c_str(), flushError);
        }
        DeleteFileW(tempPath.c_str());
        MoveFileExW(previousPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH);
        return false;
    }

    if (!MoveFileExW(tempPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        const DWORD moveError = GetLastError();
        LogMessage(LogLevel::Error, L"SessionStore::Save failed to replace %ls from %ls (error=%lu)", path.c_str(),
                   tempPath.c_str(), moveError);
        DeleteFileW(tempPath.c_str());
        if (!MoveFileExW(previousPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
            const DWORD restoreError = GetLastError();
            LogMessage(LogLevel::Error, L"SessionStore::Save failed to restore %ls from %ls (error=%lu)", path.c_str(),
                       previousPath.c_str(), restoreError);
        }
        return false;
    }

    if (!DeleteFileW(previousPath.c_str())) {
        const DWORD deleteError = GetLastError();
        if (deleteError != ERROR_FILE_NOT_FOUND) {
            LogMessage(LogLevel::Warning, L"SessionStore::Save failed to remove backup %ls (error=%lu)",
                       previousPath.c_str(), deleteError);
        }
    }

    return true;
}

std::wstring ResolveStoragePath() {
    std::wstring base = GetShellTabsDataDirectory();
    if (base.empty()) {
        return {};
    }
    if (!base.empty() && base.back() != L'\\') {
        base.push_back(L'\\');
    }
    base += kStorageFile;
    return base;
}
 
}  // namespace

std::atomic<long> g_activeSessionCount{0};

SessionStore::SessionStore() : SessionStore(ResolveStoragePath()) {}

SessionStore::SessionStore(std::wstring storagePath) : m_storagePath(std::move(storagePath)) {
    if (m_storagePath.empty()) {
        m_storagePath = ResolveStoragePath();
    }
}

std::wstring SessionStore::BuildPathForToken(const std::wstring& token) {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return {};
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }

    std::wstring sanitized;
    sanitized.reserve(token.size());
    for (wchar_t ch : token) {
        if (iswalnum(ch) || ch == L'-' || ch == L'_') {
            sanitized.push_back(ch);
        } else if (!iswspace(ch)) {
            sanitized.push_back(L'_');
        }
    }
    if (sanitized.empty()) {
        sanitized = L"window";
    }

    directory += L"session-";
    directory += sanitized;
    directory += L".db";
    return directory;
}

bool SessionStore::WasPreviousSessionUnclean() {
    if (g_activeSessionCount.load(std::memory_order_acquire) > 0) {
        return false;
    }
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return false;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += kCrashMarkerFile;
    return GetFileAttributesW(directory.c_str()) != INVALID_FILE_ATTRIBUTES;
}

void SessionStore::MarkSessionActive() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    const std::wstring path = directory + kCrashMarkerFile;
    const long previous = g_activeSessionCount.fetch_add(1, std::memory_order_acq_rel);
    if (previous > 0) {
        return;
    }
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                              FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (file != INVALID_HANDLE_VALUE) {
        CloseHandle(file);
    }

    CleanupCheckpointArtifacts(false);
}

void SessionStore::ClearSessionMarker() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        g_activeSessionCount.store(0, std::memory_order_release);
        return;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    const std::wstring path = directory + kCrashMarkerFile;
    long previous = g_activeSessionCount.fetch_sub(1, std::memory_order_acq_rel);
    if (previous <= 1) {
        g_activeSessionCount.store(0, std::memory_order_release);
        DeleteFileW(path.c_str());
        CleanupCheckpointArtifacts(true);
    }
}

bool SessionStore::Load(SessionData& data) const {
    data = {};
    if (m_storagePath.empty()) {
        return false;
    }

    SessionData parsed;
    const LoadAttemptStatus status = AttemptLoadPath(m_storagePath, true, &parsed);
    if (status == LoadAttemptStatus::kSuccess) {
        data = std::move(parsed);
        return true;
    }
    if (status == LoadAttemptStatus::kError) {
        return false;
    }
    if (status == LoadAttemptStatus::kMissing) {
        return true;
    }

    const std::wstring checkpointPath = BuildSiblingPath(m_storagePath, L".previous");
    SessionData checkpointData;
    const LoadAttemptStatus checkpointStatus = AttemptLoadPath(checkpointPath, false, &checkpointData);
    if (checkpointStatus == LoadAttemptStatus::kSuccess) {
        data = std::move(checkpointData);
        if (CopyFileW(checkpointPath.c_str(), m_storagePath.c_str(), FALSE)) {
            LogMessage(LogLevel::Warning, L"SessionStore::Load recovered session from checkpoint %ls",
                       checkpointPath.c_str());
        } else {
            const DWORD error = GetLastError();
            LogMessage(LogLevel::Warning,
                       L"SessionStore::Load recovered session from %ls but failed to restore %ls (error=%lu)",
                       checkpointPath.c_str(), m_storagePath.c_str(), error);
        }
        return true;
    }

    if (checkpointStatus == LoadAttemptStatus::kMissing) {
        LogMessage(LogLevel::Error, L"SessionStore::Load corrupt session at %ls with no checkpoint", m_storagePath.c_str());
    } else if (checkpointStatus == LoadAttemptStatus::kCorrupt) {
        LogMessage(LogLevel::Error, L"SessionStore::Load corrupt checkpoint at %ls", checkpointPath.c_str());
    } else if (checkpointStatus == LoadAttemptStatus::kError) {
        LogMessage(LogLevel::Error, L"SessionStore::Load failed to read checkpoint %ls", checkpointPath.c_str());
    }

    return false;
}

bool SessionStore::Save(const SessionData& data) const {
    if (m_storagePath.empty()) {
        return false;
    }

    EnsureParentDirectory(m_storagePath);

    std::wstring payload;
    payload += kVersionToken;
    payload += L"|6\n";
    payload += kSelectedToken;
    payload += L"|" + std::to_wstring(data.selectedGroup) + L"|" + std::to_wstring(data.selectedTab) + L"\n";
    payload += kSequenceToken;
    payload += L"|" + std::to_wstring(std::max(data.groupSequence, 1)) + L"\n";
    payload += kDockToken;
    payload += L"|" + DockModeToString(data.dockMode) + L"\n";

    for (const auto& group : data.groups) {
        payload += kGroupToken;
        payload += L"|" + group.name + L"|" + (group.collapsed ? L"1" : L"0") + L"|" +
                   (group.headerVisible ? L"1" : L"0") + L"|" + (group.hasOutline ? L"1" : L"0") + L"|" +
                   ColorToString(group.outlineColor) + L"|" + OutlineStyleToString(group.outlineStyle) + L"|" +
                   group.savedGroupId + L"\n";
        for (const auto& tab : group.tabs) {
            payload += kTabToken;
            payload += L"|" + tab.name + L"|" + tab.tooltip + L"|" + (tab.hidden ? L"1" : L"0") + L"|" + tab.path +
                       L"|" + std::to_wstring(static_cast<unsigned long long>(tab.lastActivatedTick)) + L"|" +
                       std::to_wstring(static_cast<unsigned long long>(tab.activationOrdinal)) + L"\n";
        }
    }

    const std::string payloadUtf8 = WideToUtf8(payload);
    if (!payload.empty() && payloadUtf8.empty()) {
        LogMessage(LogLevel::Error, L"SessionStore::Save failed to encode payload for %ls", m_storagePath.c_str());
        return false;
    }

    const uint64_t checksum = ComputeChecksum(payloadUtf8);
    const uint64_t length = static_cast<uint64_t>(payloadUtf8.size());

    wchar_t checksumBuffer[17] = {};
    _ui64tow(checksum, checksumBuffer, 16);
    for (wchar_t* ch = checksumBuffer; *ch; ++ch) {
        *ch = towupper(*ch);
    }

    wchar_t lengthBuffer[21] = {};
    _ui64tow(length, lengthBuffer, 10);

    std::wstring content;
    content.reserve(payload.size() + 32);
    content += kIntegrityToken;
    content += L"|";
    content += checksumBuffer;
    content += L"|";
    content += lengthBuffer;
    content += L"\n";
    content += payload;

    if (!WriteSessionFileAtomically(m_storagePath, content)) {
        return false;
    }

    return true;
}

}  // namespace shelltabs
