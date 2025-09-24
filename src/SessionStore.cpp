#include "SessionStore.h"

#include <ShlObj.h>
#include <Shlwapi.h>

#include <algorithm>
#include <string>

namespace shelltabs {
namespace {
constexpr wchar_t kStorageDirectory[] = L"ShellTabs";
constexpr wchar_t kStorageFile[] = L"session.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kGroupToken[] = L"group";
constexpr wchar_t kTabToken[] = L"tab";
constexpr wchar_t kSelectedToken[] = L"selected";
constexpr wchar_t kSequenceToken[] = L"sequence";
constexpr wchar_t kCommentChar = L'#';

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

std::wstring Utf8ToWide(const std::string& utf8) {
    if (utf8.empty()) {
        return {};
    }
    const int length = MultiByteToWideChar(CP_UTF8, 0, utf8.data(), static_cast<int>(utf8.size()), nullptr, 0);
    if (length <= 0) {
        return {};
    }
    std::wstring result(length, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.data(), static_cast<int>(utf8.size()), result.data(), length);
    return result;
}

std::string WideToUtf8(const std::wstring& wide) {
    if (wide.empty()) {
        return {};
    }
    const int length = WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, nullptr, nullptr);
    if (length <= 0) {
        return {};
    }
    std::string result(length, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), result.data(), length, nullptr, nullptr);
    return result;
}

bool ParseBool(const std::wstring& token) {
    return token == L"1" || token == L"true" || token == L"TRUE";
}

std::wstring ResolveStoragePath() {
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

    if (!base.empty() && base.back() != L'\\') {
        base.push_back(L'\\');
    }
    base += kStorageFile;
    return base;
}

}  // namespace

SessionStore::SessionStore() : m_storagePath(ResolveStoragePath()) {}

bool SessionStore::Load(SessionData& data) const {
    data = {};
    if (m_storagePath.empty()) {
        return false;
    }

    HANDLE file = CreateFileW(m_storagePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file, &size) || size.HighPart != 0) {
        CloseHandle(file);
        return false;
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
            if (tokens.size() < 2 || tokens[1] != L"1") {
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
        } else if (header == kGroupToken) {
            if (tokens.size() < 3) {
                continue;
            }
            SessionGroup group;
            group.name = tokens[1];
            group.collapsed = ParseBool(tokens[2]);
            if (tokens.size() >= 4) {
                group.splitView = ParseBool(tokens[3]);
            }
            if (tokens.size() >= 5) {
                group.splitPrimary = _wtoi(tokens[4].c_str());
            }
            if (tokens.size() >= 6) {
                group.splitSecondary = _wtoi(tokens[5].c_str());
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
    content += L"|1\n";
    content += kSelectedToken;
    content += L"|" + std::to_wstring(data.selectedGroup) + L"|" + std::to_wstring(data.selectedTab) + L"\n";
    content += kSequenceToken;
    content += L"|" + std::to_wstring(std::max(data.groupSequence, 1)) + L"\n";

    for (const auto& group : data.groups) {
        content += kGroupToken;
        content += L"|" + group.name + L"|" + (group.collapsed ? L"1" : L"0") + L"|" +
                   (group.splitView ? L"1" : L"0") + L"|" + std::to_wstring(group.splitPrimary) + L"|" +
                   std::to_wstring(group.splitSecondary) + L"\n";
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
