#include "SessionStore.h"

#include "StringUtils.h"

#include "ColorSerialization.h"

#include "Logging.h"

#include <Shlwapi.h>

#include <algorithm>
#include <cstdint>
#include <mutex>
#include <string_view>
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
constexpr wchar_t kBackupSuffix[] = L".bak";
constexpr wchar_t kChecksumToken[] = L"checksum";

void NotifySessionChecksumMismatch(const std::wstring& corruptedPath) {
    static std::once_flag s_corruptionNoticeOnce;
    std::call_once(s_corruptionNoticeOnce, [&]() {
        std::wstring message =
            L"ShellTabs could not restore saved tabs because the session data failed an integrity check.";
        if (!corruptedPath.empty()) {
            message.append(L"\n\nFile: ");
            message.append(corruptedPath);
        }
        message.append(L"\n\nA new session has been started.");

        MessageBoxW(nullptr, message.c_str(), L"ShellTabs",
                    MB_OK | MB_ICONWARNING | MB_SETFOREGROUND | MB_TOPMOST);
    });
}

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

std::wstring BuildBackupPath(const std::wstring& storagePath) {
    if (storagePath.empty()) {
        return {};
    }

    return storagePath + kBackupSuffix;
}

}  // namespace

void SessionStore::SetMarkerReady(bool ready) const {
    m_markerReady.store(ready, std::memory_order_release);
}

bool SessionStore::MarkerReady() const noexcept {
    return m_markerReady.load(std::memory_order_acquire);
}

namespace {

uint64_t ComputeChecksum(std::wstring_view payload) {
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

    std::wstring_view contentView{content};
    std::wstring_view payload = contentView;
    bool checksumPresent = false;
    bool checksumValid = true;

    const size_t newline = contentView.find(L'\n');
    if (newline != std::wstring::npos) {
        std::wstring_view headerLine = TrimView(contentView.substr(0, newline));
        if (!headerLine.empty()) {
            auto headerTokens = Split(headerLine, L'|');
            for (auto& token : headerTokens) {
                token = TrimView(token);
            }
            if (!headerTokens.empty() && headerTokens.front() == kChecksumToken) {
                checksumPresent = true;
                payload = contentView.substr(newline + 1);
                if (headerTokens.size() >= 2) {
                    uint64_t expected = 0;
                    if (!TryParseUint64(headerTokens[1], &expected)) {
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
        std::wstring_view headerLine = TrimView(contentView);
        if (!headerLine.empty() && headerLine.rfind(kChecksumToken, 0) == 0) {
            checksumPresent = true;
            checksumValid = false;
            payload = {};
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
                                         [&](const std::vector<std::wstring_view>& tokens) {
                                             if (tokens.empty()) {
                                                 return true;
                                             }

                                             const std::wstring_view header = tokens.front();
                                             if (header == kVersionToken) {
                                                 if (tokens.size() < 2) {
                                                     return false;
                                                 }
                                                 version = std::max(1, ParseInt(tokens[1]));
                                                 if (version > 6) {
                                                     return false;
                                                 }
                                                 versionSeen = true;
                                                 return true;
                                             }

                                             if (header == kSelectedToken) {
                                                 if (tokens.size() >= 3) {
                                                     parsedData.selectedGroup = ParseInt(tokens[1]);
                                                     parsedData.selectedTab = ParseInt(tokens[2]);
                                                 }
                                                 return true;
                                             }

                                             if (header == kSequenceToken) {
                                                 if (tokens.size() >= 2) {
                                                     parsedData.groupSequence = std::max(1, ParseInt(tokens[1]));
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
                                                     undo.groupIndex = ParseInt(tokens[1]);
                                                     undo.groupRemoved = ParseBool(tokens[2]);
                                                     undo.selectionIndex = ParseInt(tokens[3]);
                                                     undo.hasGroupInfo = ParseBool(tokens[4]);
                                                     size_t index = 5;
                                                     if (undo.hasGroupInfo && tokens.size() > index) {
                                                         const std::wstring_view nameToken = tokens[index++];
                                                         undo.groupInfo.name.assign(nameToken.begin(), nameToken.end());
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
                                                             const std::wstring colorToken(tokens[index++]);
                                                             undo.groupInfo.outlineColor =
                                                                 ParseColor(colorToken, undo.groupInfo.outlineColor);
                                                         }
                                                         if (tokens.size() > index) {
                                                             const std::wstring outlineToken(tokens[index++]);
                                                             undo.groupInfo.outlineStyle =
                                                                 ParseOutlineStyle(outlineToken, undo.groupInfo.outlineStyle);
                                                         }
                                                         if (tokens.size() > index) {
                                                             const std::wstring_view groupIdToken = tokens[index++];
                                                             undo.groupInfo.savedGroupId.assign(groupIdToken.begin(),
                                                                                                groupIdToken.end());
                                                         }
                                                     }
                                                 }
                                                 parsedData.lastClosed = std::move(undo);
                                                 return true;
                                             }

                                             if (header == kUndoTabToken) {
                                                 if (!parsedData.lastClosed) {
                                                     return true;
                                                 }
                                                 SessionClosedTab entry;
                                                 size_t index = 1;
                                                 if (tokens.size() > index) {
                                                     entry.index = ParseInt(tokens[index]);
                                                     ++index;
                                                 }
                                                 if (tokens.size() > index) {
                                                     const std::wstring_view nameToken = tokens[index++];
                                                     entry.tab.name.assign(nameToken.begin(), nameToken.end());
                                                 }
                                                 if (tokens.size() > index) {
                                                     const std::wstring_view tooltipToken = tokens[index++];
                                                     entry.tab.tooltip.assign(tooltipToken.begin(), tooltipToken.end());
                                                 }
                                                 if (tokens.size() > index) {
                                                     entry.tab.hidden = ParseBool(tokens[index]);
                                                     ++index;
                                                 }
                                                 if (version >= 6 && tokens.size() > index) {
                                                     entry.tab.pinned = ParseBool(tokens[index]);
                                                     ++index;
                                                 }
                                                 if (tokens.size() > index) {
                                                     const std::wstring_view pathToken = tokens[index];
                                                     entry.tab.path.assign(pathToken.begin(), pathToken.end());
                                                 }
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
                                                         const std::wstring outlineColorToken(tokens[index]);
                                                         group.outlineColor =
                                                             ParseColor(outlineColorToken, group.outlineColor);
                                                         ++index;
                                                     }
                                                     if (version >= 4 && tokens.size() > index) {
                                                         const std::wstring outlineStyleToken(tokens[index]);
                                                         group.outlineStyle =
                                                             ParseOutlineStyle(outlineStyleToken, group.outlineStyle);
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
                                                         uint64_t tick = 0;
                                                         TryParseUint64(tokens[index], &tick);
                                                         tab.lastActivatedTick = static_cast<ULONGLONG>(tick);
                                                         ++index;
                                                     }
                                                     if (tokens.size() > index) {
                                                         uint64_t ordinal = 0;
                                                         TryParseUint64(tokens[index], &ordinal);
                                                         tab.activationOrdinal = ordinal;
                                                         ++index;
                                                     }
                                                     if (version >= 6 && tokens.size() > index) {
                                                         tab.pinned = ParseBool(tokens[index]);
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
    if (!MarkerReady()) {
        return false;
    }

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

    if (!markerPath.empty() && GetFileAttributesW(markerPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    const std::wstring legacyMarker = BuildLegacyMarkerPath();
    if (!legacyMarker.empty() && legacyMarker != markerPath &&
        GetFileAttributesW(legacyMarker.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    const bool staleTempDetected = CleanupStaleTemp(m_storagePath);
    if (staleTempDetected) {
        return true;
    }

    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);
    if (!checkpointPath.empty() &&
        GetFileAttributesW(checkpointPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    const std::wstring backupPath = BuildBackupPath(m_storagePath);
    if (!backupPath.empty() &&
        GetFileAttributesW(backupPath.c_str()) != INVALID_FILE_ATTRIBUTES &&
        GetFileAttributesW(m_storagePath.c_str()) == INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    return false;
}

void SessionStore::MarkSessionActive() const {
    if (!MarkerReady()) {
        return;
    }

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
}

bool SessionStore::Load(SessionData& data) const {
    data = {};
    if (m_storagePath.empty()) {
        return false;
    }

    const std::wstring tempPath = BuildTempPath(m_storagePath);
    const std::wstring checkpointPath = BuildCheckpointPath(m_storagePath);
    const std::wstring backupPath = BuildBackupPath(m_storagePath);

    struct SnapshotCandidate {
        std::wstring path;
        const wchar_t* label = L"";
    };

    std::vector<SnapshotCandidate> candidates;
    if (!tempPath.empty() && GetFileAttributesW(tempPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        candidates.push_back({tempPath, L"temp"});
    }
    candidates.push_back({m_storagePath, L"primary"});
    if (!checkpointPath.empty()) {
        candidates.push_back({checkpointPath, L"checkpoint"});
    }
    if (!backupPath.empty()) {
        candidates.push_back({backupPath, L"backup"});
    }

    bool checksumMismatch = false;
    std::wstring corruptionPath;

    for (const auto& candidate : candidates) {
        std::wstring content;
        bool exists = false;
        if (!ReadUtf8File(candidate.path, &content, &exists)) {
            LogMessage(LogLevel::Warning, L"SessionStore failed to read %ls (%ls) (error=%lu)",
                       candidate.path.c_str(), candidate.label, GetLastError());
            continue;
        }
        if (!exists) {
            continue;
        }

        SessionData candidateData;
        std::wstring snapshot;
        const SessionFileStatus status = ParseSessionDocument(content, candidateData, snapshot);
        if (status == SessionFileStatus::kChecksumMismatch) {
            checksumMismatch = true;
            corruptionPath = candidate.path;
            LogMessage(LogLevel::Warning, L"SessionStore checksum mismatch for %ls (%ls)", candidate.path.c_str(),
                       candidate.label);
            continue;
        }
        if (status == SessionFileStatus::kParseError || status == SessionFileStatus::kEmpty) {
            LogMessage(LogLevel::Warning, L"SessionStore invalid snapshot %ls (%ls) status=%d", candidate.path.c_str(),
                       candidate.label, static_cast<int>(status));
            continue;
        }

        if (candidate.path == tempPath) {
            DeleteFileW(tempPath.c_str());
        }

        LogMessage(LogLevel::Info, L"SessionStore restored snapshot from %ls (%ls)", candidate.path.c_str(),
                   candidate.label);
        data = std::move(candidateData);
        m_lastSerializedSnapshot = std::move(snapshot);
        return true;
    }

    if (checksumMismatch) {
        NotifySessionChecksumMismatch(corruptionPath);
    }

    return false;
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
                       std::to_wstring(static_cast<unsigned long long>(tab.activationOrdinal)) + L"|" +
                       (tab.pinned ? L"1" : L"0") + L"\n";
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
                        L"|" + (entry.tab.hidden ? L"1" : L"0") + L"|" + (entry.tab.pinned ? L"1" : L"0") + L"|" +
                        entry.tab.path + L"\n";
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
    const std::wstring backupPath = BuildBackupPath(m_storagePath);

    if (tempPath.empty()) {
        return false;
    }

    DeleteFileW(tempPath.c_str());
    HANDLE tempFile = CreateFileW(tempPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                  FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (tempFile == INVALID_HANDLE_VALUE) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to create temp file %ls (error=%lu)", tempPath.c_str(),
                   GetLastError());
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
        return false;
    }

    if (!backupPath.empty()) {
        MoveFileExW(m_storagePath.c_str(), backupPath.c_str(), MOVEFILE_REPLACE_EXISTING);
    }

    if (!MoveFileExW(tempPath.c_str(), m_storagePath.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        const DWORD promoteError = GetLastError();
        LogMessage(LogLevel::Warning, L"SessionStore failed to promote temp file %ls -> %ls (error=%lu)",
                   tempPath.c_str(), m_storagePath.c_str(), promoteError);
        DeleteFileW(tempPath.c_str());
        if (!backupPath.empty()) {
            MoveFileExW(backupPath.c_str(), m_storagePath.c_str(), MOVEFILE_REPLACE_EXISTING);
        }
        return false;
    }

    m_lastSerializedSnapshot = std::move(serialized);
    return true;
}

}  // namespace shelltabs
