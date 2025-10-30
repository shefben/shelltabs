#include "SessionStore.h"

#include <ShlObj.h>
#include <Shlwapi.h>

#include <algorithm>
#include <atomic>
#include <cwctype>
#include <cwchar>
#include <sstream>
#include <string>

#include "Utilities.h"

namespace shelltabs {
namespace {
constexpr wchar_t kStorageDirectory[] = L"ShellTabs";
constexpr wchar_t kStorageFile[] = L"session.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kGroupToken[] = L"group";
constexpr wchar_t kTabToken[] = L"tab";
constexpr wchar_t kSelectedToken[] = L"selected";
constexpr wchar_t kSequenceToken[] = L"sequence";
constexpr wchar_t kDockToken[] = L"dock";
constexpr wchar_t kCommentChar = L'#';
constexpr wchar_t kCrashMarkerFile[] = L"session.lock";

std::wstring Trim(const std::wstring& value) {
    const auto begin = value.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos) {
        return {};
    }
    const auto end = value.find_last_not_of(L" \t\r\n");
    return value.substr(begin, end - begin + 1);
}

std::vector<std::wstring> Split(const std::wstring& value, wchar_t delimiter) {
    std::vector<std::wstring> parts;
    size_t start = 0;
    while (start <= value.size()) {
        const size_t pos = value.find(delimiter, start);
        if (pos == std::wstring::npos) {
            parts.emplace_back(value.substr(start));
            break;
        }
        parts.emplace_back(value.substr(start, pos - start));
        start = pos + 1;
    }
    return parts;
}

bool ParseBool(const std::wstring& token) {
    return token == L"1" || token == L"true" || token == L"TRUE";
}

TabBandDockMode ParseDockMode(const std::wstring& token) {
    if (token.empty()) {
        return TabBandDockMode::kAutomatic;
    }
    if (_wcsicmp(token.c_str(), L"top") == 0) {
        return TabBandDockMode::kTop;
    }
    if (_wcsicmp(token.c_str(), L"bottom") == 0) {
        return TabBandDockMode::kBottom;
    }
    if (_wcsicmp(token.c_str(), L"left") == 0) {
        return TabBandDockMode::kLeft;
    }
    if (_wcsicmp(token.c_str(), L"right") == 0) {
        return TabBandDockMode::kRight;
    }
    return TabBandDockMode::kAutomatic;
}

std::wstring DockModeToString(TabBandDockMode mode) {
    switch (mode) {
        case TabBandDockMode::kTop:
            return L"top";
        case TabBandDockMode::kBottom:
            return L"bottom";
        case TabBandDockMode::kLeft:
            return L"left";
        case TabBandDockMode::kRight:
            return L"right";
        default:
            return L"auto";
    }
}

COLORREF ParseColor(const std::wstring& token, COLORREF fallback) {
    if (token.empty()) {
        return fallback;
    }
    unsigned int value = 0;
    std::wistringstream stream(token);
    if (token.size() > 2 && token[0] == L'0' && (token[1] == L'x' || token[1] == L'X')) {
        stream.ignore(2);
        stream >> std::hex >> value;
    } else if (!token.empty() && token[0] == L'#') {
        stream.ignore(1);
        stream >> std::hex >> value;
    } else {
        stream >> std::hex >> value;
    }
    return RGB((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
}

std::wstring ColorToString(COLORREF color) {
    wchar_t buffer[16];
    swprintf_s(buffer, L"%06X", (GetRValue(color) << 16) | (GetGValue(color) << 8) | GetBValue(color));
    return buffer;
}

TabGroupOutlineStyle ParseOutlineStyle(const std::wstring& token, TabGroupOutlineStyle fallback) {
    if (token.empty()) {
        return fallback;
    }

    if (_wcsicmp(token.c_str(), L"solid") == 0) {
        return TabGroupOutlineStyle::kSolid;
    }
    if (_wcsicmp(token.c_str(), L"dashed") == 0) {
        return TabGroupOutlineStyle::kDashed;
    }
    if (_wcsicmp(token.c_str(), L"dotted") == 0) {
        return TabGroupOutlineStyle::kDotted;
    }

    const int value = _wtoi(token.c_str());
    switch (value) {
        case 0:
            return TabGroupOutlineStyle::kSolid;
        case 1:
            return TabGroupOutlineStyle::kDashed;
        case 2:
            return TabGroupOutlineStyle::kDotted;
        default:
            break;
    }

    return fallback;
}

std::wstring OutlineStyleToString(TabGroupOutlineStyle style) {
    switch (style) {
        case TabGroupOutlineStyle::kDashed:
            return L"dashed";
        case TabGroupOutlineStyle::kDotted:
            return L"dotted";
        case TabGroupOutlineStyle::kSolid:
        default:
            return L"solid";
    }
}

std::wstring ResolveStorageDirectory() {
    PWSTR knownFolder = nullptr;
    if (FAILED(SHGetKnownFolderPath(FOLDERID_RoamingAppData, KF_FLAG_CREATE, nullptr, &knownFolder)) || !knownFolder) {
        return {};
    }

    std::wstring base(knownFolder);
    CoTaskMemFree(knownFolder);

    if (!base.empty() && base.back() != L'\\') {
        base.push_back(L'\\');
    }
    base += kStorageDirectory;
    CreateDirectoryW(base.c_str(), nullptr);

    return base;
}

std::wstring ResolveStoragePath() {
    std::wstring base = ResolveStorageDirectory();
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
    std::wstring directory = ResolveStorageDirectory();
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
    std::wstring directory = ResolveStorageDirectory();
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
    std::wstring directory = ResolveStorageDirectory();
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
    std::wstring directory = ResolveStorageDirectory();
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

    HANDLE file = CreateFileW(m_storagePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                              FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND) {
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
        return false;
    }

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file, &size) || size.HighPart != 0) {
        CloseHandle(file);
        return false;
    }

    if (size.QuadPart == 0) {
        CloseHandle(file);
        return true;
    }

    std::string buffer;
    buffer.resize(static_cast<size_t>(size.QuadPart));
    DWORD bytesRead = 0;
    if (!buffer.empty() && !ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr)) {
        CloseHandle(file);
        return false;
    }
    CloseHandle(file);
    buffer.resize(bytesRead);

    const std::wstring content = Utf8ToWide(buffer);
    if (content.empty()) {
        return false;
    }

    bool versionSeen = false;
    int version = 1;
    SessionGroup* currentGroup = nullptr;

    size_t lineStart = 0;
    while (lineStart < content.size()) {
        size_t lineEnd = content.find(L'\n', lineStart);
        std::wstring line = content.substr(lineStart, lineEnd == std::wstring::npos ? std::wstring::npos : lineEnd - lineStart);
        if (lineEnd == std::wstring::npos) {
            lineStart = content.size();
        } else {
            lineStart = lineEnd + 1;
        }

        line = Trim(line);
        if (line.empty() || line.front() == kCommentChar) {
            continue;
        }

        auto tokens = Split(line, L'|');
        if (tokens.empty()) {
            continue;
        }

        for (auto& token : tokens) {
            token = Trim(token);
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
        } else if (header == kSelectedToken) {
            if (tokens.size() >= 3) {
                data.selectedGroup = _wtoi(tokens[1].c_str());
                data.selectedTab = _wtoi(tokens[2].c_str());
            }
        } else if (header == kSequenceToken) {
            if (tokens.size() >= 2) {
                data.groupSequence = std::max(1, _wtoi(tokens[1].c_str()));
            }
        } else if (header == kDockToken) {
            if (tokens.size() >= 2) {
                data.dockMode = ParseDockMode(tokens[1]);
            }
        } else if (header == kGroupToken) {
            if (tokens.size() < 3) {
                continue;
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
        } else if (header == kTabToken) {
            if (!currentGroup || tokens.size() < 5) {
                continue;
            }
            SessionTab tab;
            tab.name = tokens[1];
            tab.tooltip = tokens[2];
            tab.hidden = ParseBool(tokens[3]);
            tab.path = tokens[4];
            currentGroup->tabs.emplace_back(std::move(tab));
        }
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

    const std::string utf8 = WideToUtf8(content);

    HANDLE file = CreateFileW(m_storagePath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }

    DWORD bytesWritten = 0;
    const BOOL result = WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &bytesWritten, nullptr);
    CloseHandle(file);
    return result != FALSE;
}

}  // namespace shelltabs
