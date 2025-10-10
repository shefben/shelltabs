#include "GroupStore.h"

#include <ShlObj.h>
#include <Shlwapi.h>

#include <algorithm>
#include <sstream>

namespace shelltabs {
namespace {
constexpr wchar_t kStorageDirectory[] = L"ShellTabs";
constexpr wchar_t kStorageFile[] = L"groups.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kGroupToken[] = L"group";
constexpr wchar_t kTabToken[] = L"tab";
constexpr wchar_t kCommentChar = L'#';

std::wstring Trim(const std::wstring& value) {
    const size_t begin = value.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos) {
        return {};
    }
    const size_t end = value.find_last_not_of(L" \t\r\n");
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
    const int length = WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, nullptr,
                                           nullptr);
    if (length <= 0) {
        return {};
    }
    std::string result(length, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), result.data(), length, nullptr, nullptr);
    return result;
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
    } else if (token.size() > 1 && token[0] == L'#') {
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

}  // namespace

GroupStore& GroupStore::Instance() {
    static GroupStore instance;
    return instance;
}

std::vector<std::wstring> GroupStore::GroupNames() const {
    std::vector<std::wstring> names;
    names.reserve(m_groups.size());
    for (const auto& group : m_groups) {
        names.push_back(group.name);
    }
    std::sort(names.begin(), names.end(), [](const std::wstring& a, const std::wstring& b) {
        return _wcsicmp(a.c_str(), b.c_str()) < 0;
    });
    return names;
}

const SavedGroup* GroupStore::Find(const std::wstring& name) const {
    if (!EnsureLoaded()) {
        return nullptr;
    }
    for (const auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            return &group;
        }
    }
    return nullptr;
}

bool GroupStore::Load() {
    if (m_loaded) {
        return true;
    }

    const std::wstring path = ResolveStoragePath();
    if (path.empty()) {
        return false;
    }

    m_storagePath = path;
    m_groups.clear();

    HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        m_loaded = true;
        return true;
    }

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file, &size) || size.HighPart != 0) {
        CloseHandle(file);
        m_loaded = true;
        return false;
    }

    std::string buffer;
    buffer.resize(static_cast<size_t>(size.QuadPart));
    DWORD bytesRead = 0;
    if (!buffer.empty() && !ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr)) {
        CloseHandle(file);
        m_loaded = true;
        return false;
    }
    CloseHandle(file);
    buffer.resize(bytesRead);

    const std::wstring content = Utf8ToWide(buffer);
    if (content.empty()) {
        m_loaded = true;
        return true;
    }

    size_t lineStart = 0;
    SavedGroup* currentGroup = nullptr;
    int version = 1;

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

        if (tokens[0] == kVersionToken) {
            if (tokens.size() >= 2) {
                version = _wtoi(tokens[1].c_str());
                if (version < 1) {
                    version = 1;
                }
            }
            continue;
        }

        if (tokens[0] == kGroupToken) {
            if (tokens.size() < 3) {
                currentGroup = nullptr;
                continue;
            }
            SavedGroup group;
            group.name = tokens[1];
            group.color = ParseColor(tokens[2], RGB(0, 120, 215));
            m_groups.emplace_back(std::move(group));
            currentGroup = &m_groups.back();
            continue;
        }

        if (tokens[0] == kTabToken) {
            if (!currentGroup || tokens.size() < 2) {
                continue;
            }
            currentGroup->tabPaths.emplace_back(tokens[1]);
        }
    }

    m_loaded = true;
    return true;
}

bool GroupStore::Save() const {
    if (!EnsureLoaded()) {
        return false;
    }

    std::wstring path = m_storagePath;
    if (path.empty()) {
        path = ResolveStoragePath();
        if (path.empty()) {
            return false;
        }
        m_storagePath = path;
    }

    const size_t separator = path.find_last_of(L"\\/");
    if (separator != std::wstring::npos) {
        std::wstring directory = path.substr(0, separator);
        if (!directory.empty()) {
            CreateDirectoryW(directory.c_str(), nullptr);
        }
    }

    std::wstring content;
    content += kVersionToken;
    content += L"|1\n";
    for (const auto& group : m_groups) {
        content += kGroupToken;
        content += L"|" + group.name + L"|" + ColorToString(group.color) + L"\n";
        for (const auto& pathEntry : group.tabPaths) {
            content += kTabToken;
            content += L"|" + pathEntry + L"\n";
        }
    }

    const std::string utf8 = WideToUtf8(content);

    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }

    DWORD bytesWritten = 0;
    const BOOL result = WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &bytesWritten, nullptr);
    CloseHandle(file);
    return result != FALSE;
}

bool GroupStore::Upsert(SavedGroup group) {
    if (!EnsureLoaded()) {
        return false;
    }

    for (auto& existing : m_groups) {
        if (_wcsicmp(existing.name.c_str(), group.name.c_str()) == 0) {
            existing.name = group.name;
            existing.color = group.color;
            existing.tabPaths = std::move(group.tabPaths);
            return Save();
        }
    }

    m_groups.emplace_back(std::move(group));
    return Save();
}

bool GroupStore::UpdateTabs(const std::wstring& name, const std::vector<std::wstring>& tabPaths) {
    if (!EnsureLoaded()) {
        return false;
    }

    for (auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            group.tabPaths = tabPaths;
            return Save();
        }
    }
    return false;
}

bool GroupStore::UpdateColor(const std::wstring& name, COLORREF color) {
    if (!EnsureLoaded()) {
        return false;
    }

    for (auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            group.color = color;
            return Save();
        }
    }
    return false;
}

bool GroupStore::Remove(const std::wstring& name) {
    if (!EnsureLoaded()) {
        return false;
    }
    auto it = std::remove_if(m_groups.begin(), m_groups.end(), [&](const SavedGroup& group) {
        return _wcsicmp(group.name.c_str(), name.c_str()) == 0;
    });
    if (it == m_groups.end()) {
        return false;
    }
    m_groups.erase(it, m_groups.end());
    return Save();
}

std::wstring GroupStore::ResolveStoragePath() const {
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

bool GroupStore::EnsureLoaded() const {
    if (!m_loaded) {
        const_cast<GroupStore*>(this)->Load();
    }
    return m_loaded;
}

}  // namespace shelltabs
