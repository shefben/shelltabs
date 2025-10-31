#include "SessionStore.h"

#include "StringUtils.h"

#include "ColorSerialization.h"

#include "Logging.h"

#include <Shlwapi.h>

#include <algorithm>
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
constexpr wchar_t kJournalSuffix[] = L".journal";

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

std::wstring BuildJournalPath(const std::wstring& storagePath) {
    if (storagePath.empty()) {
        return {};
    }
    return storagePath + kJournalSuffix;
}

bool CleanupStaleJournal(const std::wstring& storagePath) {
    const std::wstring journalPath = BuildJournalPath(storagePath);
    if (journalPath.empty()) {
        return false;
    }

    const DWORD attributes = GetFileAttributesW(journalPath.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES) {
        return false;
    }

    LogMessage(LogLevel::Warning,
               L"SessionStore detected stale journal %ls; rolling back to previous snapshot",
               journalPath.c_str());
    if (!DeleteFileW(journalPath.c_str())) {
        LogMessage(LogLevel::Warning,
                   L"SessionStore failed to delete stale journal %ls (error=%lu)",
                   journalPath.c_str(), GetLastError());
    }
    return true;
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

    const bool staleJournalDetected = CleanupStaleJournal(m_storagePath);

    if (!markerPath.empty() && GetFileAttributesW(markerPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    const std::wstring legacyMarker = BuildLegacyMarkerPath();
    if (!legacyMarker.empty() && legacyMarker != markerPath &&
        GetFileAttributesW(legacyMarker.c_str()) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    return staleJournalDetected;
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

    CleanupStaleJournal(m_storagePath);

    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(m_storagePath, &content, &fileExists)) {
        return false;
    }

    if (!fileExists) {
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
        return true;
    }

    if (content.empty()) {
        m_lastSerializedSnapshot = std::wstring();
        return true;
    }

    bool versionSeen = false;
    int version = 1;
    SessionGroup* currentGroup = nullptr;

    const bool parsed = ParseConfigLines(content, kCommentChar, L'|', [&](const std::vector<std::wstring>& tokens) {
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
                data.selectedGroup = _wtoi(tokens[1].c_str());
                data.selectedTab = _wtoi(tokens[2].c_str());
            }
            return true;
        }

        if (header == kSequenceToken) {
            if (tokens.size() >= 2) {
                data.groupSequence = std::max(1, _wtoi(tokens[1].c_str()));
            }
            return true;
        }

        if (header == kDockToken) {
            if (tokens.size() >= 2) {
                data.dockMode = ParseDockMode(tokens[1]);
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
                        undo.groupInfo.outlineColor = ParseColor(tokens[index++], undo.groupInfo.outlineColor);
                    }
                    if (tokens.size() > index) {
                        undo.groupInfo.outlineStyle = ParseOutlineStyle(tokens[index++], undo.groupInfo.outlineStyle);
                    }
                    if (tokens.size() > index) {
                        undo.groupInfo.savedGroupId = tokens[index++];
                    }
                }
            }
            data.lastClosed = std::move(undo);
            return true;
        }

        if (header == kUndoTabToken) {
            if (!data.lastClosed || tokens.size() < 6) {
                return true;
            }
            SessionClosedTab entry;
            entry.index = _wtoi(tokens[1].c_str());
            entry.tab.name = tokens[2];
            entry.tab.tooltip = tokens[3];
            entry.tab.hidden = ParseBool(tokens[4]);
            entry.tab.path = tokens[5];
            data.lastClosed->tabs.emplace_back(std::move(entry));
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
                    ++index;  // legacy placeholder flag
                }
                if (tokens.size() > index) {
                    ++index;  // legacy placeholder primary tab
                }
                if (tokens.size() > index) {
                    ++index;  // legacy placeholder secondary tab
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
            data.groups.emplace_back(std::move(group));
            currentGroup = &data.groups.back();
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

    if (!parsed) {
        return false;
    }

    if (versionSeen) {
        m_lastSerializedSnapshot = content;
    }

    return versionSeen && !data.groups.empty();
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

    std::wstring serialized;
    serialized += kVersionToken;
    serialized += L"|5\n";
    serialized += kSelectedToken;
    serialized += L"|" + std::to_wstring(data.selectedGroup) + L"|" + std::to_wstring(data.selectedTab) + L"\n";
    serialized += kSequenceToken;
    serialized += L"|" + std::to_wstring(std::max(data.groupSequence, 1)) + L"\n";
    serialized += kDockToken;
    serialized += L"|" + DockModeToString(data.dockMode) + L"\n";

    for (const auto& group : data.groups) {
        serialized += kGroupToken;
        serialized += L"|" + group.name + L"|" + (group.collapsed ? L"1" : L"0") + L"|" +
                      (group.headerVisible ? L"1" : L"0") + L"|" + (group.hasOutline ? L"1" : L"0") + L"|" +
                      ColorToString(group.outlineColor) + L"|" + OutlineStyleToString(group.outlineStyle) + L"|" +
                      group.savedGroupId + L"\n";
        for (const auto& tab : group.tabs) {
            serialized += kTabToken;
            serialized += L"|" + tab.name + L"|" + tab.tooltip + L"|" + (tab.hidden ? L"1" : L"0") + L"|" + tab.path +
                          L"|" + std::to_wstring(static_cast<unsigned long long>(tab.lastActivatedTick)) + L"|" +
                          std::to_wstring(static_cast<unsigned long long>(tab.activationOrdinal)) + L"\n";
        }
    }

    if (data.lastClosed && !data.lastClosed->tabs.empty()) {
        const auto& undo = *data.lastClosed;
        serialized += kUndoToken;
        serialized += L"|" + std::to_wstring(undo.groupIndex) + L"|" + (undo.groupRemoved ? L"1" : L"0") + L"|" +
                       std::to_wstring(undo.selectionIndex) + L"|" + (undo.hasGroupInfo ? L"1" : L"0");
        if (undo.hasGroupInfo) {
            serialized += L"|" + undo.groupInfo.name + L"|" + (undo.groupInfo.collapsed ? L"1" : L"0") + L"|" +
                           (undo.groupInfo.headerVisible ? L"1" : L"0") + L"|" +
                           (undo.groupInfo.hasOutline ? L"1" : L"0") + L"|" +
                           ColorToString(undo.groupInfo.outlineColor) + L"|" +
                           OutlineStyleToString(undo.groupInfo.outlineStyle) + L"|" + undo.groupInfo.savedGroupId;
        }
        serialized += L"\n";
        for (const auto& entry : undo.tabs) {
            serialized += kUndoTabToken;
            serialized += L"|" + std::to_wstring(entry.index) + L"|" + entry.tab.name + L"|" + entry.tab.tooltip +
                           L"|" + (entry.tab.hidden ? L"1" : L"0") + L"|" + entry.tab.path + L"\n";
        }
    }

    if (m_lastSerializedSnapshot && *m_lastSerializedSnapshot == serialized) {
        return true;
    }

    const std::string utf8 = WideToUtf8(serialized);
    if (!serialized.empty() && utf8.empty()) {
        return false;
    }

    const std::wstring journalPath = BuildJournalPath(m_storagePath);
    if (journalPath.empty()) {
        return false;
    }

    DeleteFileW(journalPath.c_str());
    HANDLE journal = CreateFileW(journalPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                 FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (journal == INVALID_HANDLE_VALUE) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to create journal %ls (error=%lu)",
                   journalPath.c_str(), GetLastError());
        return false;
    }

    bool writeSucceeded = true;
    DWORD writeError = ERROR_SUCCESS;
    if (!utf8.empty()) {
        DWORD bytesWritten = 0;
        if (!WriteFile(journal, utf8.data(), static_cast<DWORD>(utf8.size()), &bytesWritten, nullptr) ||
            bytesWritten != utf8.size()) {
            writeSucceeded = false;
            if (writeError == ERROR_SUCCESS) {
                writeError = GetLastError();
            }
        }
    }
    if (writeSucceeded && !FlushFileBuffers(journal)) {
        writeSucceeded = false;
        if (writeError == ERROR_SUCCESS) {
            writeError = GetLastError();
        }
    }
    CloseHandle(journal);

    if (!writeSucceeded) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to serialize journal %ls (error=%lu)",
                   journalPath.c_str(), writeError);
        DeleteFileW(journalPath.c_str());
        return false;
    }

    if (!MoveFileExW(journalPath.c_str(), m_storagePath.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        LogMessage(LogLevel::Warning, L"SessionStore failed to promote journal %ls -> %ls (error=%lu)",
                   journalPath.c_str(), m_storagePath.c_str(), GetLastError());
        DeleteFileW(journalPath.c_str());
        return false;
    }

    m_lastSerializedSnapshot = std::move(serialized);
    return true;
}

}  // namespace shelltabs
