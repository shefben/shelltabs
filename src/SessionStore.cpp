#include "SessionStore.h"

#include "StringUtils.h"

#include "ColorSerialization.h"

#include "Logging.h"

#include <Shlwapi.h>

#include <algorithm>
#include <cstdint>
#include <mutex>
#include <unordered_map>
#include <cwctype>
#include <cwchar>
#include <string>

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
constexpr wchar_t kUndoToken[] = L"undo";
constexpr wchar_t kUndoTabToken[] = L"undotab";
constexpr wchar_t kCommentChar = L'#';
constexpr wchar_t kCrashMarkerFile[] = L"session.lock";
constexpr wchar_t kMarkerSuffix[] = L".lock";
constexpr wchar_t kTempSuffix[] = L".tmp";
constexpr wchar_t kCheckpointSuffix[] = L".previous";
constexpr wchar_t kChecksumToken[] = L"checksum";

struct SessionMarkerState {
    std::mutex mutex;
    std::unordered_map<std::wstring, long> counts;
};

SessionMarkerState& GetSessionMarkerState() {
    static SessionMarkerState state;
    return state;
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

std::wstring BuildLegacyMarkerPath() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return {};
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += kCrashMarkerFile;
    return directory;
}

std::wstring BuildMarkerPath(const std::wstring& storagePath) {
    if (storagePath.empty()) {
        return BuildLegacyMarkerPath();
    }
    return storagePath + kMarkerSuffix;
}

std::wstring BuildTempPath(const std::wstring& storagePath) {
    if (storagePath.empty()) {
        return {};
    }
    return storagePath + kTempSuffix;
}

std::wstring BuildCheckpointPath(const std::wstring& storagePath) {
    if (storagePath.empty()) {
        return {};
    }

    return storagePath + kCheckpointSuffix;
}

uint64_t ComputeChecksum(const std::wstring& payload) {
    static_assert(sizeof(wchar_t) == 2, "SessionStore assumes UTF-16 wchar_t");
    uint64_t hash = 1469598103934665603ull;  // FNV-1a offset basis
    constexpr uint64_t kPrime = 1099511628211ull;
    for (wchar_t ch : payload) {
        uint16_t value = static_cast<uint16_t>(ch);
        hash ^= static_cast<uint8_t>(value & 0xFF);
        hash *= kPrime;
        hash ^= static_cast<uint8_t>((value >> 8) & 0xFF);
        hash *= kPrime;
    }
    return hash;
}

bool CleanupStaleTemp(const std::wstring& storagePath) {
    const std::wstring tempPath = BuildTempPath(storagePath);
    if (tempPath.empty()) {
        return false;
    }

    const DWORD attributes = GetFileAttributesW(tempPath.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES) {
        return false;
    }

    LogMessage(LogLevel::Warning,
               L"SessionStore detected stale temp file %ls; removing stale snapshot",
               tempPath.c_str());
    if (!DeleteFileW(tempPath.c_str())) {
        LogMessage(LogLevel::Warning,
                   L"SessionStore failed to delete stale temp file %ls (error=%lu)",
                   tempPath.c_str(), GetLastError());
    }
    return true;
}

enum class SessionFileStatus {
    kSuccess,
    kEmpty,
    kChecksumMismatch,
    kParseError,
};

SessionFileStatus ParseSessionDocument(const std::wstring& content, SessionData& outData,
                                       std::wstring& snapshotOut) {
    if (content.empty()) {
        outData = SessionData{};
        snapshotOut.clear();
        return SessionFileStatus::kEmpty;
    }

    std::wstring payload = content;
    bool checksumPresent = false;
    bool checksumValid = true;

    const size_t newline = content.find(L'\n');
    if (newline != std::wstring::npos) {
        std::wstring headerLine = Trim(content.substr(0, newline));
        if (!headerLine.empty()) {
            auto headerTokens = Split(headerLine, L'|');
            for (auto& token : headerTokens) {
                token = Trim(token);
            }
            if (!headerTokens.empty() && headerTokens.front() == kChecksumToken) {
                checksumPresent = true;
                payload = content.substr(newline + 1);
                if (headerTokens.size() >= 2) {
                    wchar_t* end = nullptr;
                    const uint64_t expected = _wcstoui64(headerTokens[1].c_str(), &end, 10);
                    if (end == headerTokens[1].c_str()) {
                        checksumValid = false;
                    } else {
                        const uint64_t actual = ComputeChecksum(payload);
                        checksumValid = actual == expected;
                    }
                } else {
                    checksumValid = false;
                }
            }
        }
    } else {
        std::wstring headerLine = Trim(content);
        if (!headerLine.empty() && headerLine.rfind(kChecksumToken, 0) == 0) {
            checksumPresent = true;
            checksumValid = false;
            payload.clear();
        }
    }

    if (checksumPresent && !checksumValid) {
        return SessionFileStatus::kChecksumMismatch;
    }

    if (payload.empty()) {
        outData = SessionData{};
        snapshotOut = content;
        return SessionFileStatus::kEmpty;
    }

    SessionData parsedData;
    bool versionSeen = false;
    int version = 1;
    SessionGroup* currentGroup = nullptr;

    const bool parsed = ParseConfigLines(payload, kCommentChar, L'|',
                                         [&](const std::vector<std::wstring>& tokens) {
                                             if (tokens.empty()) {
                                                 return true;
                                             }

                                             const std::wstring& header = tokens.front();
                                             if (header == kVersionToken) {
                                                 if (tokens.size() < 2) {
                                                     return false;
                                                 }
                                                 version = std::max(1, _wtoi(tokens[1].c_str()));
                                                 if (version > 5) {
                                                     return false;
                                                 }
                                                 versionSeen = true;
                                                 return true;
                                             }

                                             if (header == kSelectedToken) {
                                                 if (tokens.size() >= 3) {
                                                     parsedData.selectedGroup = _wtoi(tokens[1].c_str());
                                                     parsedData.selectedTab = _wtoi(tokens[2].c_str());
                                                 }
                                                 return true;
                                             }

                                             if (header == kSequenceToken) {
                                                 if (tokens.size() >= 2) {
                                                     parsedData.groupSequence = std::max(1, _wtoi(tokens[1].c_str()));
                                                 }
                                                 return true;
                                             }

                                             if (header == kDockToken) {
                                                 if (tokens.size() >= 2) {
                                                     parsedData.dockMode = ParseDockMode(tokens[1]);
                                                 }
                                                 return true;
                                             }

                                             if (header == kUndoToken) {
                                                 SessionClosedSet undo;
                                                 if (tokens.size() >= 5) {
                                                     undo.groupIndex = _wtoi(tokens[1].c_str());
                                                     undo.groupRemoved = ParseBool(tokens[2]);
                                                     undo.selectionIndex = _wtoi(tokens[3].c_str());
                                                     undo.hasGroupInfo = ParseBool(tokens[4]);
                                                     size_t index = 5;
                                                     if (undo.hasGroupInfo && tokens.size() > index) {
                                                         undo.groupInfo.name = tokens[index++];
                                                         if (tokens.size() > index) {
                                                             undo.groupInfo.collapsed = ParseBool(tokens[index++]);
                                                         }
                                                         if (tokens.size() > index) {
                                                             undo.groupInfo.headerVisible = ParseBool(tokens[index++]);
                                                         }
                                                         if (tokens.size() > index) {
                                                             undo.groupInfo.hasOutline = ParseBool(tokens[index++]);
                                                         }
                                                         if (tokens.size() > index) {
                                                             undo.groupInfo.outlineColor =
                                                                 ParseColor(tokens[index++], undo.groupInfo.outlineColor);
                                                         }
                                                         if (tokens.size() > index) {
                                                             undo.groupInfo.outlineStyle =
                                                                 ParseOutlineStyle(tokens[index++], undo.groupInfo.outlineStyle);
                                                         }
                                                         if (tokens.size() > index) {
                                                             undo.groupInfo.savedGroupId = tokens[index++];
                                                         }
                                                     }
                                                 }
                                                 parsedData.lastClosed = std::move(undo);
                                                 return true;
                                             }

                                             if (header == kUndoTabToken) {
                                                 if (!parsedData.lastClosed || tokens.size() < 6) {
                                                     return true;
                                                 }
                                                 SessionClosedTab entry;
                                                 entry.index = _wtoi(tokens[1].c_str());
                                                 entry.tab.name = tokens[2];
                                                 entry.tab.tooltip = tokens[3];
                                                 entry.tab.hidden = ParseBool(tokens[4]);
                                                 entry.tab.path = tokens[5];
                                                 parsedData.lastClosed->tabs.emplace_back(std::move(entry));
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
                                                 parsedData.groups.emplace_back(std::move(group));
                                                 currentGroup = &parsedData.groups.back();
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
                                                         tab.lastActivatedTick =
                                                             _wcstoui64(tokens[index].c_str(), nullptr, 10);
                                                         ++index;
                                                     }
                                                     if (tokens.size() > index) {
                                                         tab.activationOrdinal =
                                                             _wcstoui64(tokens[index].c_str(), nullptr, 10);
                                                         ++index;
                                                     }
                                                 }
                                                 currentGroup->tabs.emplace_back(std::move(tab));
                                                 return true;
                                             }

                                             return true;
                                         });

    if (!parsed) {
        return SessionFileStatus::kParseError;
    }

    if (!versionSeen) {
        return SessionFileStatus::kParseError;
    }

    if (parsedData.groups.empty()) {
        return SessionFileStatus::kParseError;
    }

    outData = std::move(parsedData);
    snapshotOut = content;
    return SessionFileStatus::kSuccess;
}

}  // namespace

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

bool SessionStore::WasPreviousSessionUnclean() const {
    const std::wstring markerPath = BuildMarkerPath(m_storagePath);
    if (!markerPath.empty()) {
        auto& state = GetSessionMarkerState();
        {
            std::scoped_lock lock(state.mutex);
            const auto it = state.counts.find(markerPath);
            if (it != state.counts.end() && it->second > 0) {
                return false;
            }
        }
    }

    const bool staleTempDetected = CleanupStaleTemp(m_storagePath);

    bool checkpointDetected = false;
    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);
    if (!checkpointPath.empty() &&
        GetFileAttributesW(checkpointPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        checkpointDetected = true;
        m_pendingCheckpointCleanup = true;
    }

    if (!markerPath.empty() && GetFileAttributesW(markerPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    const std::wstring legacyMarker = BuildLegacyMarkerPath();
    if (!legacyMarker.empty() && legacyMarker != markerPath &&
        GetFileAttributesW(legacyMarker.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    return staleTempDetected || checkpointDetected;
}

void SessionStore::MarkSessionActive() const {
    const std::wstring markerPath = BuildMarkerPath(m_storagePath);
    if (markerPath.empty()) {
        return;
    }

    auto& state = GetSessionMarkerState();
    bool shouldCreateMarker = false;
    {
        std::scoped_lock lock(state.mutex);
        long& count = state.counts[markerPath];
        if (count == 0) {
            shouldCreateMarker = true;
        }
        ++count;
    }

    if (!shouldCreateMarker) {
        return;
    }

    HANDLE file = CreateFileW(markerPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS,
                              FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (file != INVALID_HANDLE_VALUE) {
        CloseHandle(file);
    } else {
        LogMessage(LogLevel::Warning, L"SessionStore failed to create crash marker %ls (error=%lu)",
                   markerPath.c_str(), GetLastError());
    }

    const std::wstring legacyMarker = BuildLegacyMarkerPath();
    if (!legacyMarker.empty() && legacyMarker != markerPath) {
        DeleteFileW(legacyMarker.c_str());
    }

    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);
    if (!checkpointPath.empty() &&
        GetFileAttributesW(checkpointPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        m_pendingCheckpointCleanup = true;
    }
}

void SessionStore::ClearSessionMarker() const {
    const std::wstring markerPath = BuildMarkerPath(m_storagePath);
    if (markerPath.empty()) {
        return;
    }

    auto& state = GetSessionMarkerState();
    bool shouldDeleteMarker = false;
    {
        std::scoped_lock lock(state.mutex);
        auto it = state.counts.find(markerPath);
        if (it == state.counts.end()) {
            return;
        }
        long& count = it->second;
        if (count > 0) {
            --count;
        }
        if (count <= 0) {
            state.counts.erase(it);
            shouldDeleteMarker = true;
        }
    }

    if (!shouldDeleteMarker) {
        return;
    }

    if (!DeleteFileW(markerPath.c_str())) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to delete crash marker %ls (error=%lu)",
                   markerPath.c_str(), GetLastError());
    }

    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);
    if (!checkpointPath.empty() &&
        (m_pendingCheckpointCleanup ||
         GetFileAttributesW(checkpointPath.c_str()) != INVALID_FILE_ATTRIBUTES)) {
        if (!DeleteFileW(checkpointPath.c_str())) {
            const DWORD checkpointError = GetLastError();
            if (checkpointError != ERROR_FILE_NOT_FOUND) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to delete checkpoint %ls (error=%lu)",
                           checkpointPath.c_str(), checkpointError);
            }
        } else {
            m_pendingCheckpointCleanup = false;
        }
    }
}

bool SessionStore::Load(SessionData& data) const {
    data = {};
    if (m_storagePath.empty()) {
        return false;
    }

    CleanupStaleTemp(m_storagePath);

    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(m_storagePath, &content, &fileExists)) {
        const DWORD readError = GetLastError();
        LogMessage(LogLevel::Warning, L"SessionStore failed to read %ls (error=%lu)", m_storagePath.c_str(),
                   readError);
        return false;
    }

    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);

    auto restoreFromCheckpoint = [&](const wchar_t* reason) -> bool {
        if (checkpointPath.empty()) {
            return false;
        }

        std::wstring checkpointContent;
        bool checkpointExists = false;
        if (!ReadUtf8File(checkpointPath, &checkpointContent, &checkpointExists)) {
            const DWORD checkpointError = GetLastError();
            LogMessage(LogLevel::Warning,
                       L"SessionStore failed to read checkpoint %ls (error=%lu) while handling %ls",
                       checkpointPath.c_str(), checkpointError, reason);
            return false;
        }
        if (!checkpointExists) {
            LogMessage(LogLevel::Warning,
                       L"SessionStore checkpoint %ls missing while handling %ls",
                       checkpointPath.c_str(), reason);
            return false;
        }

        SessionData fallbackData;
        std::wstring fallbackSnapshot;
        const SessionFileStatus status = ParseSessionDocument(checkpointContent, fallbackData, fallbackSnapshot);
        if (status == SessionFileStatus::kChecksumMismatch) {
            LogMessage(LogLevel::Warning,
                       L"SessionStore checkpoint %ls failed checksum while handling %ls",
                       checkpointPath.c_str(), reason);
            return false;
        }
        if (status == SessionFileStatus::kParseError) {
            LogMessage(LogLevel::Warning,
                       L"SessionStore checkpoint %ls was malformed while handling %ls",
                       checkpointPath.c_str(), reason);
            return false;
        }

        LogMessage(LogLevel::Warning,
                   L"SessionStore restoring checkpoint %ls after %ls",
                   checkpointPath.c_str(), reason);

        data = std::move(fallbackData);
        m_lastSerializedSnapshot = std::move(fallbackSnapshot);

        if (!MoveFileExW(checkpointPath.c_str(), m_storagePath.c_str(),
                         MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
            LogMessage(LogLevel::Warning,
                       L"SessionStore failed to promote checkpoint %ls -> %ls (error=%lu)",
                       checkpointPath.c_str(), m_storagePath.c_str(), GetLastError());
            m_pendingCheckpointCleanup = true;
        } else {
            m_pendingCheckpointCleanup = false;
        }

        return true;
    };

    if (!fileExists) {
        if (restoreFromCheckpoint(L"missing session file")) {
            return true;
        }

        const size_t separator = m_storagePath.find_last_of(L"\\/");
        if (separator != std::wstring::npos) {
            std::wstring directory = m_storagePath.substr(0, separator);
            if (!directory.empty()) {
                CreateDirectoryW(directory.c_str(), nullptr);
            }
        }

        HANDLE created = CreateFileW(m_storagePath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW,
                                     FILE_ATTRIBUTE_NORMAL, nullptr);
        if (created != INVALID_HANDLE_VALUE) {
            CloseHandle(created);
        }
        m_lastSerializedSnapshot = std::wstring();
        m_pendingCheckpointCleanup = false;
        return true;
    }

    SessionData parsedData;
    std::wstring snapshot;
    const SessionFileStatus status = ParseSessionDocument(content, parsedData, snapshot);

    switch (status) {
        case SessionFileStatus::kSuccess:
            data = std::move(parsedData);
            m_lastSerializedSnapshot = std::move(snapshot);
            break;
        case SessionFileStatus::kEmpty:
            data = SessionData{};
            m_lastSerializedSnapshot = std::move(snapshot);
            break;
        case SessionFileStatus::kChecksumMismatch:
            LogMessage(LogLevel::Warning,
                       L"SessionStore checksum mismatch detected for %ls", m_storagePath.c_str());
            return restoreFromCheckpoint(L"checksum mismatch");
        case SessionFileStatus::kParseError:
            LogMessage(LogLevel::Warning,
                       L"SessionStore failed to parse %ls", m_storagePath.c_str());
            return restoreFromCheckpoint(L"parse failure");
    }

    if (!checkpointPath.empty() &&
        GetFileAttributesW(checkpointPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        if (!DeleteFileW(checkpointPath.c_str())) {
            const DWORD checkpointError = GetLastError();
            if (checkpointError != ERROR_FILE_NOT_FOUND) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to delete checkpoint %ls (error=%lu)",
                           checkpointPath.c_str(), checkpointError);
                m_pendingCheckpointCleanup = true;
            }
        } else {
            m_pendingCheckpointCleanup = false;
        }
    }

    return true;
}

bool SessionStore::Save(const SessionData& data) const {
    if (m_storagePath.empty()) {
        return false;
    }

    const size_t separator = m_storagePath.find_last_of(L"\\/");
    if (separator != std::wstring::npos) {
        std::wstring directory = m_storagePath.substr(0, separator);
        if (!directory.empty()) {
            CreateDirectoryW(directory.c_str(), nullptr);
        }
    }

    std::wstring payload;
    payload += kVersionToken;
    payload += L"|5\n";
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

    if (data.lastClosed && !data.lastClosed->tabs.empty()) {
        const auto& undo = *data.lastClosed;
        payload += kUndoToken;
        payload += L"|" + std::to_wstring(undo.groupIndex) + L"|" + (undo.groupRemoved ? L"1" : L"0") + L"|" +
                    std::to_wstring(undo.selectionIndex) + L"|" + (undo.hasGroupInfo ? L"1" : L"0");
        if (undo.hasGroupInfo) {
            payload += L"|" + undo.groupInfo.name + L"|" + (undo.groupInfo.collapsed ? L"1" : L"0") + L"|" +
                        (undo.groupInfo.headerVisible ? L"1" : L"0") + L"|" +
                        (undo.groupInfo.hasOutline ? L"1" : L"0") + L"|" +
                        ColorToString(undo.groupInfo.outlineColor) + L"|" +
                        OutlineStyleToString(undo.groupInfo.outlineStyle) + L"|" + undo.groupInfo.savedGroupId;
        }
        payload += L"\n";
        for (const auto& entry : undo.tabs) {
            payload += kUndoTabToken;
            payload += L"|" + std::to_wstring(entry.index) + L"|" + entry.tab.name + L"|" + entry.tab.tooltip +
                        L"|" + (entry.tab.hidden ? L"1" : L"0") + L"|" + entry.tab.path + L"\n";
        }
    }

    const uint64_t checksum = ComputeChecksum(payload);
    std::wstring serialized;
    serialized.reserve(payload.size() + 32);
    serialized += kChecksumToken;
    serialized += L"|";
    serialized += std::to_wstring(checksum);
    serialized += L"\n";
    serialized += payload;

    if (m_lastSerializedSnapshot && *m_lastSerializedSnapshot == serialized) {
        return true;
    }

    const std::string utf8 = WideToUtf8(serialized);
    if (!serialized.empty() && utf8.empty()) {
        return false;
    }

    const std::wstring tempPath = BuildTempPath(m_storagePath);
    if (tempPath.empty()) {
        return false;
    }

    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);
    bool checkpointCreated = false;
    if (!checkpointPath.empty()) {
        if (MoveFileExW(m_storagePath.c_str(), checkpointPath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
            checkpointCreated = true;
        } else {
            const DWORD rotateError = GetLastError();
            if (rotateError != ERROR_FILE_NOT_FOUND) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to rotate %ls -> %ls (error=%lu)",
                           m_storagePath.c_str(), checkpointPath.c_str(), rotateError);
                return false;
            }
        }
    }

    DeleteFileW(tempPath.c_str());
    HANDLE tempFile = CreateFileW(tempPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                  FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (tempFile == INVALID_HANDLE_VALUE) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to create temp file %ls (error=%lu)", tempPath.c_str(),
                   GetLastError());
        if (checkpointCreated) {
            if (!MoveFileExW(checkpointPath.c_str(), m_storagePath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to restore checkpoint %ls after temp creation failure (error=%lu)",
                           checkpointPath.c_str(), GetLastError());
                m_pendingCheckpointCleanup = true;
            }
        }
        return false;
    }

    bool writeSucceeded = true;
    DWORD writeError = ERROR_SUCCESS;
    if (!utf8.empty()) {
        DWORD bytesWritten = 0;
        if (!WriteFile(tempFile, utf8.data(), static_cast<DWORD>(utf8.size()), &bytesWritten, nullptr) ||
            bytesWritten != utf8.size()) {
            writeSucceeded = false;
            writeError = GetLastError();
        }
    }
    if (writeSucceeded && !FlushFileBuffers(tempFile)) {
        writeSucceeded = false;
        writeError = GetLastError();
    }
    CloseHandle(tempFile);

    if (!writeSucceeded) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to serialize temp file %ls (error=%lu)", tempPath.c_str(),
                   writeError);
        DeleteFileW(tempPath.c_str());
        if (checkpointCreated) {
            if (!MoveFileExW(checkpointPath.c_str(), m_storagePath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to restore checkpoint %ls after write failure (error=%lu)",
                           checkpointPath.c_str(), GetLastError());
                m_pendingCheckpointCleanup = true;
            }
        }
        return false;
    }

    if (!MoveFileExW(tempPath.c_str(), m_storagePath.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        const DWORD promoteError = GetLastError();
        LogMessage(LogLevel::Warning, L"SessionStore failed to promote temp file %ls -> %ls (error=%lu)",
                   tempPath.c_str(), m_storagePath.c_str(), promoteError);
        DeleteFileW(tempPath.c_str());
        if (checkpointCreated) {
            if (!MoveFileExW(checkpointPath.c_str(), m_storagePath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to restore checkpoint %ls after promotion failure (error=%lu)",
                           checkpointPath.c_str(), GetLastError());
                m_pendingCheckpointCleanup = true;
            }
        }
        return false;
    }

    if (!checkpointPath.empty() &&
        (checkpointCreated || GetFileAttributesW(checkpointPath.c_str()) != INVALID_FILE_ATTRIBUTES)) {
        if (!DeleteFileW(checkpointPath.c_str())) {
            const DWORD checkpointError = GetLastError();
            if (checkpointError != ERROR_FILE_NOT_FOUND) {
                LogMessage(LogLevel::Warning,
                           L"SessionStore failed to delete checkpoint %ls (error=%lu)",
                           checkpointPath.c_str(), checkpointError);
                m_pendingCheckpointCleanup = true;
            }
        } else {
            m_pendingCheckpointCleanup = false;
        }
    }

    m_lastSerializedSnapshot = std::move(serialized);
    return true;
}

}  // namespace shelltabs
