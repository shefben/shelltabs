#include "SessionStore.h"

#include "StringUtils.h"
#include "ColorSerialization.h"
#include "Logging.h"

#include <Shlwapi.h>

#include <algorithm>
#include <cstdint>
#include <string_view>
#include <cwctype>
#include <cwchar>
#include <string>

#include "Utilities.h"

namespace shelltabs {
namespace {

constexpr wchar_t kSessionFile[] = L"session.db";
constexpr wchar_t kSessionTmpFile[] = L"session.db.tmp";
constexpr wchar_t kSessionBakFile[] = L"session.db.bak";
constexpr wchar_t kMarkerFile[] = L"session.active";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kGroupToken[] = L"group";
constexpr wchar_t kTabToken[] = L"tab";
constexpr wchar_t kSelectedToken[] = L"selected";
constexpr wchar_t kSequenceToken[] = L"sequence";
constexpr wchar_t kDockToken[] = L"dock";
constexpr wchar_t kUndoToken[] = L"undo";
constexpr wchar_t kUndoTabToken[] = L"undotab";
constexpr wchar_t kWindowToken[] = L"window";
constexpr wchar_t kWindowEndToken[] = L"window_end";
constexpr wchar_t kCommentChar = L'#';
constexpr wchar_t kChecksumToken[] = L"checksum";
constexpr int kCurrentVersion = 8;

// Legacy file patterns for migration
constexpr wchar_t kLegacySessionPrefix[] = L"session-";
constexpr wchar_t kLegacySessionPattern[] = L"session-*.db";
constexpr wchar_t kLegacyLockPattern[] = L"session-*.db.lock";
constexpr wchar_t kLegacyTmpPattern[] = L"session-*.db.tmp";
constexpr wchar_t kLegacyPreviousPattern[] = L"session-*.db.previous";
constexpr wchar_t kGlobalSessionFile[] = L"global-session.db";
constexpr wchar_t kGlobalSessionBak[] = L"global-session.db.bak";

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

enum class SessionFileStatus {
    kSuccess,
    kEmpty,
    kChecksumMismatch,
    kParseError,
};

// Parse a single-window session block (version 6 format).
// This parses the content between window/window_end delimiters, or a standalone v6 file.
SessionFileStatus ParseSingleWindowBlock(std::wstring_view payload, SessionData& outData) {
    if (payload.empty()) {
        outData = SessionData{};
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

                                             // Skip window/window_end tokens — handled by caller.
                                             if (header == kWindowToken || header == kWindowEndToken) {
                                                 return true;
                                             }

                                             if (header == kVersionToken) {
                                                 if (tokens.size() < 2) {
                                                     return false;
                                                 }
                                                 version = std::max(1, ParseInt(tokens[1]));
                                                 if (version > kCurrentVersion) {
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

    // For legacy single-window files (v1-v6), version is required.
    // For v8 multi-window blocks, version line is in the header, not per-window.
    // If we got groups, consider it successful even without a version line.
    if (!versionSeen && parsedData.groups.empty()) {
        return SessionFileStatus::kParseError;
    }

    if (parsedData.groups.empty()) {
        return SessionFileStatus::kEmpty;
    }

    outData = std::move(parsedData);
    return SessionFileStatus::kSuccess;
}

// Parse the v8 multi-window format: checksum header, version header, then window/window_end blocks.
bool ParseMultiWindowFile(const std::wstring& content, std::vector<SessionData>& outWindows) {
    outWindows.clear();

    if (content.empty()) {
        return true;
    }

    std::wstring_view contentView{content};
    std::wstring_view payload = contentView;

    // Check for checksum header.
    const size_t firstNewline = contentView.find(L'\n');
    if (firstNewline != std::wstring::npos) {
        std::wstring_view headerLine = TrimView(contentView.substr(0, firstNewline));
        if (!headerLine.empty()) {
            auto headerTokens = Split(headerLine, L'|');
            for (auto& token : headerTokens) {
                token = TrimView(token);
            }
            if (!headerTokens.empty() && headerTokens.front() == kChecksumToken) {
                payload = contentView.substr(firstNewline + 1);
                if (headerTokens.size() >= 2) {
                    uint64_t expected = 0;
                    if (!TryParseUint64(headerTokens[1], &expected)) {
                        return false;
                    }
                    const uint64_t actual = ComputeChecksum(payload);
                    if (actual != expected) {
                        LogMessage(LogLevel::Warning, L"SessionCoordinator checksum mismatch in session file");
                        return false;
                    }
                } else {
                    return false;
                }
            }
        }
    }

    // Check version header.
    int version = 0;
    bool hasVersion = false;

    // Split payload into lines and look for version and window blocks.
    // We need to find the version line first, then split by window/window_end.
    std::vector<std::wstring> windowBlocks;
    std::wstring currentBlock;
    bool inWindow = false;

    size_t pos = 0;
    while (pos < payload.size()) {
        size_t lineEnd = payload.find(L'\n', pos);
        if (lineEnd == std::wstring::npos) {
            lineEnd = payload.size();
        }

        std::wstring_view line = TrimView(payload.substr(pos, lineEnd - pos));
        pos = lineEnd + 1;

        if (line.empty() || line[0] == kCommentChar) {
            continue;
        }

        // Check for version line (before any window block)
        if (!hasVersion && line.rfind(kVersionToken, 0) == 0) {
            auto tokens = Split(line, L'|');
            if (tokens.size() >= 2) {
                version = ParseInt(TrimView(tokens[1]));
                hasVersion = true;
            }
            continue;
        }

        if (line == kWindowToken) {
            inWindow = true;
            currentBlock.clear();
            continue;
        }

        if (line == kWindowEndToken) {
            if (inWindow && !currentBlock.empty()) {
                windowBlocks.emplace_back(std::move(currentBlock));
                currentBlock.clear();
            }
            inWindow = false;
            continue;
        }

        if (inWindow) {
            currentBlock += line;
            currentBlock += L'\n';
        }
    }

    if (!hasVersion) {
        return false;
    }

    if (version < kCurrentVersion) {
        // Versions below 8 are legacy single-window format, shouldn't reach here.
        return false;
    }

    for (const auto& block : windowBlocks) {
        SessionData windowData;
        // Inject a synthetic version line so the parser recognizes version-gated fields.
        std::wstring blockWithVersion = L"version|";
        blockWithVersion += std::to_wstring(version);
        blockWithVersion += L'\n';
        blockWithVersion += block;
        SessionFileStatus status = ParseSingleWindowBlock(blockWithVersion, windowData);
        if (status == SessionFileStatus::kSuccess) {
            outWindows.emplace_back(std::move(windowData));
        }
        // Skip empty/failed blocks silently — partial recovery is better than total failure.
    }

    return true;
}

// Parse a legacy single-window file (v1-v6).
SessionFileStatus ParseLegacyFile(const std::wstring& content, SessionData& outData) {
    if (content.empty()) {
        outData = SessionData{};
        return SessionFileStatus::kEmpty;
    }

    std::wstring_view contentView{content};
    std::wstring_view payload = contentView;

    // Check for checksum header.
    const size_t newline = contentView.find(L'\n');
    if (newline != std::wstring::npos) {
        std::wstring_view headerLine = TrimView(contentView.substr(0, newline));
        if (!headerLine.empty()) {
            auto headerTokens = Split(headerLine, L'|');
            for (auto& token : headerTokens) {
                token = TrimView(token);
            }
            if (!headerTokens.empty() && headerTokens.front() == kChecksumToken) {
                payload = contentView.substr(newline + 1);
                if (headerTokens.size() >= 2) {
                    uint64_t expected = 0;
                    if (!TryParseUint64(headerTokens[1], &expected)) {
                        return SessionFileStatus::kChecksumMismatch;
                    }
                    const uint64_t actual = ComputeChecksum(payload);
                    if (actual != expected) {
                        return SessionFileStatus::kChecksumMismatch;
                    }
                } else {
                    return SessionFileStatus::kChecksumMismatch;
                }
            }
        }
    }

    return ParseSingleWindowBlock(payload, outData);
}

// Serialize a single window's SessionData into the body format (no checksum/version header).
std::wstring SerializeWindowBlock(const SessionData& data) {
    std::wstring payload;
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

    return payload;
}

// Delete files matching a wildcard pattern in a directory.
void DeleteMatchingFiles(const std::wstring& directory, const wchar_t* pattern) {
    std::wstring searchPath = directory;
    if (!searchPath.empty() && searchPath.back() != L'\\') {
        searchPath.push_back(L'\\');
    }
    searchPath += pattern;

    WIN32_FIND_DATAW findData{};
    HANDLE findHandle = FindFirstFileW(searchPath.c_str(), &findData);
    if (findHandle == INVALID_HANDLE_VALUE) {
        return;
    }

    std::wstring dirPrefix = directory;
    if (!dirPrefix.empty() && dirPrefix.back() != L'\\') {
        dirPrefix.push_back(L'\\');
    }

    do {
        if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            continue;
        }
        std::wstring filePath = dirPrefix + findData.cFileName;
        if (DeleteFileW(filePath.c_str())) {
            LogMessage(LogLevel::Info, L"SessionCoordinator cleaned up legacy file: %ls", filePath.c_str());
        }
    } while (FindNextFileW(findHandle, &findData));

    FindClose(findHandle);
}

// Atomic write: write to tmp, backup current, rename tmp to target.
bool AtomicWriteFile(const std::wstring& targetPath, const std::wstring& tmpPath,
                     const std::wstring& bakPath, const std::string& utf8Content) {
    // Write to temp file with write-through.
    DeleteFileW(tmpPath.c_str());
    HANDLE tempFile = CreateFileW(tmpPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                                  FILE_ATTRIBUTE_NORMAL | FILE_FLAG_WRITE_THROUGH, nullptr);
    if (tempFile == INVALID_HANDLE_VALUE) {
        LogMessage(LogLevel::Warning, L"SessionCoordinator failed to create temp file %ls (error=%lu)",
                   tmpPath.c_str(), GetLastError());
        return false;
    }

    bool writeOk = true;
    if (!utf8Content.empty()) {
        DWORD bytesWritten = 0;
        if (!WriteFile(tempFile, utf8Content.data(), static_cast<DWORD>(utf8Content.size()), &bytesWritten, nullptr) ||
            bytesWritten != utf8Content.size()) {
            writeOk = false;
        }
    }
    if (writeOk && !FlushFileBuffers(tempFile)) {
        writeOk = false;
    }
    CloseHandle(tempFile);

    if (!writeOk) {
        LogMessage(LogLevel::Warning, L"SessionCoordinator failed to write temp file %ls", tmpPath.c_str());
        DeleteFileW(tmpPath.c_str());
        return false;
    }

    // Backup current session.db -> session.db.bak
    if (GetFileAttributesW(targetPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        CopyFileW(targetPath.c_str(), bakPath.c_str(), FALSE);
    }

    // Atomic rename tmp -> target.
    if (!MoveFileExW(tmpPath.c_str(), targetPath.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        LogMessage(LogLevel::Warning, L"SessionCoordinator failed to rename temp -> session file (error=%lu)",
                   GetLastError());
        DeleteFileW(tmpPath.c_str());
        return false;
    }

    return true;
}

}  // namespace

// ---------------------------------------------------------------------------
// SessionCoordinator
// ---------------------------------------------------------------------------

SessionCoordinator& SessionCoordinator::Instance() {
    static SessionCoordinator instance;
    return instance;
}

SessionCoordinator::SessionCoordinator() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    m_sessionPath = directory + kSessionFile;
    m_markerPath = directory + kMarkerFile;
}

int SessionCoordinator::Register(TabBand* band) {
    std::scoped_lock lock(m_mutex);

    // Load session file on first registration.
    if (!m_loaded) {
        LoadSessionFile();
        m_loaded = true;
    }

    int slot = m_nextSlot++;
    SlotEntry entry;
    entry.band = band;
    m_slots.emplace_back(slot, std::move(entry));

    LogMessage(LogLevel::Info, L"SessionCoordinator::Register slot=%d (total=%d)",
               slot, static_cast<int>(m_slots.size()));

    // Mark active on first registration.
    if (m_slots.size() == 1) {
        MarkActive();
    }

    return slot;
}

void SessionCoordinator::Unregister(int slot) {
    std::scoped_lock lock(m_mutex);

    auto it = std::find_if(m_slots.begin(), m_slots.end(),
                           [slot](const auto& p) { return p.first == slot; });
    if (it == m_slots.end()) {
        return;
    }

    // If the departing window has session data, push it back into the pending
    // pool so that a future window can claim it.  This handles the case where
    // Explorer creates a transient first window that claims the crash-recovery
    // data but closes almost immediately — without this the data would be lost.
    // Push to FRONT so the next window gets the most-recent data, not a stale entry.
    if (it->second.hasData && !it->second.data.groups.empty()) {
        LogMessage(LogLevel::Info,
                   L"SessionCoordinator::Unregister slot=%d returned %d groups to pending pool (front)",
                   slot, static_cast<int>(it->second.data.groups.size()));
        m_pendingWindows.insert(m_pendingWindows.begin(), std::move(it->second.data));
        m_lastReturnTick = GetTickCount64();
    }

    m_slots.erase(it);
    LogMessage(LogLevel::Info, L"SessionCoordinator::Unregister slot=%d (remaining=%d)",
               slot, static_cast<int>(m_slots.size()));

    // Save so the pending pool is persisted to session.db.
    SaveAll();

    // Only mark clean when last window unregisters AND there is no pending
    // session data waiting to be claimed by a future window.
    if (m_slots.empty() && m_pendingWindows.empty()) {
        MarkClean();
    }
}

bool SessionCoordinator::ClaimWindowData(int slot, SessionData& outData) {
    std::scoped_lock lock(m_mutex);

    if (m_pendingWindows.empty()) {
        return false;
    }

    // Cooldown: if a previous window just returned data to the pending pool
    // (likely a transient probe window that opened and closed within seconds),
    // suppress this claim for 3 seconds.  This breaks the infinite cycle of
    // claim → restore → navigate → probe-window-closes → return → claim.
    if (m_lastReturnTick != 0) {
        const ULONGLONG now = GetTickCount64();
        const ULONGLONG elapsed = now - m_lastReturnTick;
        if (elapsed < 3000) {
            LogMessage(LogLevel::Info,
                       L"SessionCoordinator::ClaimWindowData slot=%d suppressed (cooldown %llums remaining)",
                       slot, 3000 - elapsed);
            return false;
        }
    }

    outData = std::move(m_pendingWindows.front());
    m_pendingWindows.erase(m_pendingWindows.begin());

    // Store it in the slot.
    for (auto& [id, entry] : m_slots) {
        if (id == slot) {
            entry.data = outData;
            entry.hasData = true;
            break;
        }
    }

    LogMessage(LogLevel::Info, L"SessionCoordinator::ClaimWindowData slot=%d claimed window data (%d groups, %d pending remaining)",
               slot, static_cast<int>(outData.groups.size()), static_cast<int>(m_pendingWindows.size()));
    return true;
}

void SessionCoordinator::UpdateWindowData(int slot, const SessionData& data) {
    std::scoped_lock lock(m_mutex);

    for (auto& [id, entry] : m_slots) {
        if (id == slot) {
            entry.data = data;
            entry.hasData = true;
            return;
        }
    }
}

std::wstring SessionCoordinator::SerializeAllWindows() const {
    // Build the full multi-window file content.
    std::wstring payload;
    payload += kVersionToken;
    payload += L"|";
    payload += std::to_wstring(kCurrentVersion);
    payload += L"\n";

    for (const auto& [id, entry] : m_slots) {
        if (!entry.hasData || entry.data.groups.empty()) {
            continue;
        }
        payload += kWindowToken;
        payload += L"\n";
        payload += SerializeWindowBlock(entry.data);
        payload += kWindowEndToken;
        payload += L"\n";
    }

    // Also write pending (unclaimed) windows so they survive restarts.
    for (const auto& pending : m_pendingWindows) {
        if (pending.groups.empty()) {
            continue;
        }
        payload += kWindowToken;
        payload += L"\n";
        payload += SerializeWindowBlock(pending);
        payload += kWindowEndToken;
        payload += L"\n";
    }

    // Add checksum header.
    const uint64_t checksum = ComputeChecksum(payload);
    std::wstring serialized;
    serialized.reserve(payload.size() + 40);
    serialized += kChecksumToken;
    serialized += L"|";
    serialized += std::to_wstring(checksum);
    serialized += L"\n";
    serialized += payload;

    return serialized;
}

bool SessionCoordinator::SaveAll() {
    std::scoped_lock lock(m_mutex);

    if (m_sessionPath.empty()) {
        return false;
    }

    // Ensure directory exists.
    const size_t separator = m_sessionPath.find_last_of(L"\\/");
    if (separator != std::wstring::npos) {
        std::wstring directory = m_sessionPath.substr(0, separator);
        if (!directory.empty()) {
            CreateDirectoryW(directory.c_str(), nullptr);
        }
    }

    std::wstring serialized = SerializeAllWindows();

    // Dedup — skip write if nothing changed.
    if (m_lastSnapshot && *m_lastSnapshot == serialized) {
        return true;
    }

    const std::string utf8 = WideToUtf8(serialized);
    if (!serialized.empty() && utf8.empty()) {
        return false;
    }

    std::wstring directory = m_sessionPath.substr(0, m_sessionPath.find_last_of(L"\\/"));
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    const std::wstring tmpPath = directory + kSessionTmpFile;
    const std::wstring bakPath = directory + kSessionBakFile;

    if (!AtomicWriteFile(m_sessionPath, tmpPath, bakPath, utf8)) {
        return false;
    }

    m_lastSnapshot = std::move(serialized);
    return true;
}

bool SessionCoordinator::WasCrash() const {
    std::scoped_lock lock(m_mutex);
    return m_wasCrash;
}

void SessionCoordinator::MarkActive() {
    // No lock needed — only called from Register which already holds the lock.
    if (m_markerPath.empty()) {
        return;
    }

    HANDLE file = CreateFileW(m_markerPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr,
                              CREATE_ALWAYS, FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY, nullptr);
    if (file != INVALID_HANDLE_VALUE) {
        CloseHandle(file);
        LogMessage(LogLevel::Info, L"SessionCoordinator::MarkActive created marker %ls", m_markerPath.c_str());
    } else {
        LogMessage(LogLevel::Warning, L"SessionCoordinator::MarkActive failed (error=%lu)", GetLastError());
    }
}

void SessionCoordinator::MarkClean() {
    // No lock needed — only called from Unregister which already holds the lock.
    if (m_markerPath.empty()) {
        return;
    }

    if (DeleteFileW(m_markerPath.c_str())) {
        LogMessage(LogLevel::Info, L"SessionCoordinator::MarkClean deleted marker %ls", m_markerPath.c_str());
    } else {
        const DWORD error = GetLastError();
        if (error != ERROR_FILE_NOT_FOUND) {
            LogMessage(LogLevel::Warning, L"SessionCoordinator::MarkClean failed (error=%lu)", error);
        }
    }
}

void SessionCoordinator::ClearCrashState() {
    std::scoped_lock lock(m_mutex);
    if (!m_wasCrash) {
        return;
    }
    m_wasCrash = false;

    // Purge trivial single-tab entries from the pending pool.  These accumulate
    // when transient Explorer windows open, create a placeholder tab, then close
    // immediately — pushing the placeholder session back into the pool.
    const size_t before = m_pendingWindows.size();
    m_pendingWindows.erase(
        std::remove_if(m_pendingWindows.begin(), m_pendingWindows.end(),
                        [](const SessionData& d) {
                            if (d.groups.size() != 1) return false;
                            const auto& group = d.groups[0];
                            return group.tabs.size() <= 1 &&
                                   group.savedGroupId.empty();
                        }),
        m_pendingWindows.end());

    LogMessage(LogLevel::Info,
               L"SessionCoordinator::ClearCrashState cleared crash flag, purged %d stale entries, %d pending remain",
               static_cast<int>(before - m_pendingWindows.size()),
               static_cast<int>(m_pendingWindows.size()));

    // Persist the cleaned state.
    SaveAll();
}

int SessionCoordinator::RegisteredCount() const {
    std::scoped_lock lock(m_mutex);
    return static_cast<int>(m_slots.size());
}

bool SessionCoordinator::HasPendingData() const {
    std::scoped_lock lock(m_mutex);
    return !m_pendingWindows.empty();
}

void SessionCoordinator::LoadSessionFile() {
    // Check crash marker.
    if (!m_markerPath.empty() &&
        GetFileAttributesW(m_markerPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        m_wasCrash = true;
        LogMessage(LogLevel::Info, L"SessionCoordinator detected crash marker (previous session was unclean)");
    }

    // Also check for legacy crash markers (session.lock and session-*.db.lock files)
    if (!m_wasCrash) {
        std::wstring directory = GetShellTabsDataDirectory();
        if (!directory.empty()) {
            if (directory.back() != L'\\') {
                directory.push_back(L'\\');
            }
            std::wstring legacyMarker = directory + L"session.lock";
            if (GetFileAttributesW(legacyMarker.c_str()) != INVALID_FILE_ATTRIBUTES) {
                m_wasCrash = true;
                LogMessage(LogLevel::Info, L"SessionCoordinator detected legacy crash marker");
            }
        }
    }

    // Try to load the new multi-window session.db
    if (m_sessionPath.empty()) {
        return;
    }

    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(m_sessionPath, &content, &fileExists) || !fileExists) {
        // No session.db — try migrating from old per-window files.
        MigrateFromOldFormat();
        return;
    }

    if (content.empty()) {
        MigrateFromOldFormat();
        return;
    }

    // Try parsing as v8 multi-window format first.
    std::vector<SessionData> windows;
    if (ParseMultiWindowFile(content, windows)) {
        m_pendingWindows = std::move(windows);
        LogMessage(LogLevel::Info, L"SessionCoordinator loaded %d windows from session.db",
                   static_cast<int>(m_pendingWindows.size()));
        // Clean up any leftover old files.
        CleanupOldFiles();
        return;
    }

    // Fall back: try parsing as a legacy single-window file.
    SessionData legacyData;
    SessionFileStatus status = ParseLegacyFile(content, legacyData);
    if (status == SessionFileStatus::kSuccess) {
        m_pendingWindows.push_back(std::move(legacyData));
        LogMessage(LogLevel::Info, L"SessionCoordinator loaded legacy single-window session.db");
        CleanupOldFiles();
        return;
    }

    // Corrupt file — try migration from per-window files.
    LogMessage(LogLevel::Warning, L"SessionCoordinator failed to parse session.db, attempting migration");
    MigrateFromOldFormat();
}

void SessionCoordinator::MigrateFromOldFormat() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }

    // Scan for session-*.db files, sort by modification time (newest first).
    const std::wstring pattern = directory + kLegacySessionPattern;
    WIN32_FIND_DATAW findData{};
    HANDLE findHandle = FindFirstFileW(pattern.c_str(), &findData);
    if (findHandle == INVALID_HANDLE_VALUE) {
        LogMessage(LogLevel::Info, L"SessionCoordinator no legacy session files found for migration");
        return;
    }

    struct LegacyFile {
        std::wstring path;
        ULONGLONG timestamp = 0;
    };
    std::vector<LegacyFile> legacyFiles;

    do {
        if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            continue;
        }
        // Skip files that are .lock, .tmp, or .previous — we only want .db files.
        const std::wstring name = findData.cFileName;
        if (name.size() < 4) {
            continue;
        }
        // Must end with ".db" and not ".db.lock" etc.
        if (name.size() > 8 && name.substr(name.size() - 8) == L".db.lock") {
            continue;
        }
        if (name.size() > 7 && name.substr(name.size() - 7) == L".db.tmp") {
            continue;
        }
        if (name.size() > 12 && name.substr(name.size() - 12) == L".db.previous") {
            continue;
        }
        if (name.substr(name.size() - 3) != L".db") {
            continue;
        }

        LegacyFile entry;
        entry.path = directory + name;
        entry.timestamp = (static_cast<ULONGLONG>(findData.ftLastWriteTime.dwHighDateTime) << 32) |
                          findData.ftLastWriteTime.dwLowDateTime;
        legacyFiles.emplace_back(std::move(entry));
    } while (FindNextFileW(findHandle, &findData));

    FindClose(findHandle);

    if (legacyFiles.empty()) {
        LogMessage(LogLevel::Info, L"SessionCoordinator no valid legacy .db files for migration");
        return;
    }

    // Sort by modification time, newest first.
    std::sort(legacyFiles.begin(), legacyFiles.end(),
              [](const LegacyFile& a, const LegacyFile& b) { return a.timestamp > b.timestamp; });

    // Parse each as a single-window session.
    for (const auto& file : legacyFiles) {
        std::wstring content;
        bool exists = false;
        if (!ReadUtf8File(file.path, &content, &exists) || !exists || content.empty()) {
            continue;
        }

        SessionData data;
        SessionFileStatus status = ParseLegacyFile(content, data);
        if (status == SessionFileStatus::kSuccess && !data.groups.empty()) {
            m_pendingWindows.emplace_back(std::move(data));
            LogMessage(LogLevel::Info, L"SessionCoordinator migrated legacy session: %ls", file.path.c_str());
        }
    }

    LogMessage(LogLevel::Info, L"SessionCoordinator migration: %d windows recovered from %d legacy files",
               static_cast<int>(m_pendingWindows.size()), static_cast<int>(legacyFiles.size()));

    // Write the combined session file.
    if (!m_pendingWindows.empty()) {
        // Temporarily put pending windows into fake slots so SerializeAllWindows picks them up.
        // (They're already in m_pendingWindows, and SerializeAllWindows serializes those too.)
        SaveAll();
    }

    // Clean up old files.
    CleanupOldFiles();
}

void SessionCoordinator::CleanupOldFiles() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return;
    }
    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }

    // Delete all session-*.db, session-*.db.lock, session-*.db.tmp, session-*.db.previous files.
    DeleteMatchingFiles(directory, kLegacySessionPattern);
    DeleteMatchingFiles(directory, kLegacyLockPattern);
    DeleteMatchingFiles(directory, kLegacyTmpPattern);
    DeleteMatchingFiles(directory, kLegacyPreviousPattern);

    // Delete global-session.db and global-session.db.bak.
    std::wstring globalSession = directory + kGlobalSessionFile;
    if (GetFileAttributesW(globalSession.c_str()) != INVALID_FILE_ATTRIBUTES) {
        if (DeleteFileW(globalSession.c_str())) {
            LogMessage(LogLevel::Info, L"SessionCoordinator cleaned up %ls", globalSession.c_str());
        }
    }
    std::wstring globalBak = directory + kGlobalSessionBak;
    if (GetFileAttributesW(globalBak.c_str()) != INVALID_FILE_ATTRIBUTES) {
        if (DeleteFileW(globalBak.c_str())) {
            LogMessage(LogLevel::Info, L"SessionCoordinator cleaned up %ls", globalBak.c_str());
        }
    }

    // Delete legacy session.lock marker.
    std::wstring legacyMarker = directory + L"session.lock";
    if (GetFileAttributesW(legacyMarker.c_str()) != INVALID_FILE_ATTRIBUTES) {
        DeleteFileW(legacyMarker.c_str());
    }
}

// ---------------------------------------------------------------------------
// SessionStore — thin wrapper
// ---------------------------------------------------------------------------

SessionStore::SessionStore() {
    m_slot = SessionCoordinator::Instance().Register(nullptr);
    LogMessage(LogLevel::Info, L"SessionStore created (slot=%d)", m_slot);
}

SessionStore::~SessionStore() {
    if (m_slot >= 0) {
        SessionCoordinator::Instance().Unregister(m_slot);
        LogMessage(LogLevel::Info, L"SessionStore destroyed (slot=%d)", m_slot);
    }
}

bool SessionStore::Load(SessionData& data) const {
    data = {};
    return SessionCoordinator::Instance().ClaimWindowData(m_slot, data);
}

bool SessionStore::Save(const SessionData& data) {
    SessionCoordinator::Instance().UpdateWindowData(m_slot, data);
    return SessionCoordinator::Instance().SaveAll();
}

}  // namespace shelltabs
