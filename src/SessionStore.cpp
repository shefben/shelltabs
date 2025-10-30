#include "SessionStore.h"

#include "StringUtils.h"

#include "ColorSerialization.h"

#include <Shlwapi.h>

#include <algorithm>
#include <atomic>
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
constexpr wchar_t kCommentChar = L'#';
constexpr wchar_t kCrashMarkerFile[] = L"session.lock";

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
    }
}

bool SessionStore::Load(SessionData& data) const {
    data = {};
    if (m_storagePath.empty()) {
        return false;
    }

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
        return true;
    }

    if (content.empty()) {
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
            if (version > 4) {
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
            currentGroup->tabs.emplace_back(std::move(tab));
            return true;
        }

        return true;
    });

    if (!parsed) {
        return false;
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

    std::wstring content;
    content += kVersionToken;
    content += L"|4\n";
    content += kSelectedToken;
    content += L"|" + std::to_wstring(data.selectedGroup) + L"|" + std::to_wstring(data.selectedTab) + L"\n";
    content += kSequenceToken;
    content += L"|" + std::to_wstring(std::max(data.groupSequence, 1)) + L"\n";
    content += kDockToken;
    content += L"|" + DockModeToString(data.dockMode) + L"\n";

    for (const auto& group : data.groups) {
        content += kGroupToken;
        content += L"|" + group.name + L"|" + (group.collapsed ? L"1" : L"0") + L"|" +
                   (group.headerVisible ? L"1" : L"0") + L"|" + (group.hasOutline ? L"1" : L"0") + L"|" +
                   ColorToString(group.outlineColor) + L"|" + OutlineStyleToString(group.outlineStyle) + L"|" +
                   group.savedGroupId + L"\n";
        for (const auto& tab : group.tabs) {
            content += kTabToken;
            content += L"|" + tab.name + L"|" + tab.tooltip + L"|" + (tab.hidden ? L"1" : L"0") + L"|" + tab.path + L"\n";
        }
    }

    return WriteUtf8File(m_storagePath, content);
}

}  // namespace shelltabs
