#include "Tagging.h"

#include <ShlObj.h>
#include <Shlwapi.h>

#include <algorithm>
#include <functional>

#include "Utilities.h"

namespace shelltabs {
namespace {
constexpr wchar_t kStorageFileName[] = L"tags.db";
constexpr wchar_t kStorageDirectory[] = L"ShellTabs";

std::wstring ToLowerCopy(std::wstring value) {
    if (!value.empty()) {
        CharLowerBuffW(value.data(), static_cast<DWORD>(value.size()));
    }
    return value;
}

}  // namespace

TagStore& TagStore::Instance() {
    static TagStore store;
    return store;
}

TagStore::TagStore() = default;

std::vector<std::wstring> TagStore::GetTagsForPath(const std::wstring& path) const {
    std::lock_guard lock(m_mutex);
    EnsureLoadedLocked();
    const std::wstring key = NormalizeKey(path);
    auto it = m_tagMap.find(key);
    if (it == m_tagMap.end()) {
        return {};
    }
    return it->second;
}

std::wstring TagStore::GetTagListForPath(const std::wstring& path) const {
    const auto tags = GetTagsForPath(path);
    if (tags.empty()) {
        return {};
    }
    std::wstring result;
    for (size_t i = 0; i < tags.size(); ++i) {
        if (i > 0) {
            result += L", ";
        }
        result += tags[i];
    }
    return result;
}

bool TagStore::TryGetColorForPath(const std::wstring& path, COLORREF* color) const {
    return TryGetColorAndTags(path, color, nullptr);
}

bool TagStore::TryGetColorAndTags(const std::wstring& path, COLORREF* color, std::vector<std::wstring>* tags) const {
    std::lock_guard lock(m_mutex);
    EnsureLoadedLocked();
    const std::wstring key = NormalizeKey(path);
    auto it = m_tagMap.find(key);
    if (it == m_tagMap.end()) {
        return false;
    }
    if (tags) {
        *tags = it->second;
    }
    if (color) {
        if (it->second.empty()) {
            return false;
        }
        COLORREF blended = CalculateColor(it->second.front());
        for (size_t i = 1; i < it->second.size(); ++i) {
            blended = BlendColors(blended, CalculateColor(it->second[i]));
        }
        *color = blended;
    }
    return true;
}

std::wstring TagStore::NormalizeKey(const std::wstring& path) const {
    if (path.empty()) {
        return {};
    }
    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }
    return ToLowerCopy(std::move(normalized));
}

void TagStore::EnsureLoadedLocked() const {
    if (m_loaded) {
        return;
    }
    LoadLocked();
    m_loaded = true;
}

void TagStore::LoadLocked() const {
    m_tagMap.clear();
    const std::wstring storagePath = StoragePath();
    if (storagePath.empty()) {
        return;
    }

    HANDLE file = CreateFileW(storagePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return;
    }

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file, &size) || size.HighPart != 0) {
        CloseHandle(file);
        return;
    }

    std::string buffer;
    buffer.resize(static_cast<size_t>(size.QuadPart));
    DWORD bytesRead = 0;
    if (!buffer.empty() && !ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr)) {
        CloseHandle(file);
        return;
    }
    CloseHandle(file);
    buffer.resize(bytesRead);

    const std::wstring content = Utf8ToWide(buffer);
    size_t start = 0;
    while (start < content.size()) {
        size_t end = content.find(L'\n', start);
        std::wstring line = content.substr(start, end == std::wstring::npos ? std::wstring::npos : end - start);
        if (end == std::wstring::npos) {
            start = content.size();
        } else {
            start = end + 1;
        }

        if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
        }
        line = Trim(line);
        if (line.empty() || line.front() == L'#') {
            continue;
        }

        const size_t delimiter = line.find(L'|');
        if (delimiter == std::wstring::npos) {
            continue;
        }

        std::wstring pathPart = Trim(line.substr(0, delimiter));
        std::wstring tagsPart = Trim(line.substr(delimiter + 1));
        if (pathPart.empty() || tagsPart.empty()) {
            continue;
        }

        auto tagsVector = SplitTags(tagsPart);
        if (tagsVector.empty()) {
            continue;
        }

        const std::wstring key = NormalizeKey(pathPart);
        m_tagMap.emplace(key, std::move(tagsVector));
    }
}

std::wstring TagStore::StoragePath() const {
    PWSTR known = nullptr;
    if (FAILED(SHGetKnownFolderPath(FOLDERID_RoamingAppData, KF_FLAG_CREATE, nullptr, &known)) || !known) {
        return {};
    }

    std::wstring base(known);
    CoTaskMemFree(known);

    if (!base.empty() && base.back() != L'\\') {
        base += L'\\';
    }
    base += kStorageDirectory;
    CreateDirectoryW(base.c_str(), nullptr);

    if (base.empty() || base.back() != L'\\') {
        base += L'\\';
    }
    base += kStorageFileName;
    return base;
}

std::wstring TagStore::Utf8ToWide(const std::string& utf8) {
    if (utf8.empty()) {
        return {};
    }
    const int required = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), nullptr, 0);
    if (required <= 0) {
        return {};
    }
    std::wstring wide(static_cast<size_t>(required), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), static_cast<int>(utf8.size()), wide.data(), required);
    return wide;
}

std::wstring TagStore::Trim(const std::wstring& value) {
    size_t start = value.find_first_not_of(L" \t\r\n");
    if (start == std::wstring::npos) {
        return {};
    }
    size_t end = value.find_last_not_of(L" \t\r\n");
    return value.substr(start, end - start + 1);
}

std::vector<std::wstring> TagStore::SplitTags(const std::wstring& tags) {
    std::vector<std::wstring> result;
    size_t start = 0;
    while (start < tags.size()) {
        size_t end = tags.find_first_of(L";,", start);
        std::wstring token = Trim(tags.substr(start, end == std::wstring::npos ? std::wstring::npos : end - start));
        if (!token.empty()) {
            result.emplace_back(std::move(token));
        }
        if (end == std::wstring::npos) {
            break;
        }
        start = end + 1;
    }
    return result;
}

COLORREF TagStore::CalculateColor(const std::wstring& tag) {
    if (tag.empty()) {
        return RGB(0x33, 0x33, 0x33);
    }

    static constexpr COLORREF kPalette[] = {
        RGB(0xE5, 0x39, 0x35), RGB(0xD8, 0x64, 0x3D), RGB(0xFB, 0x8C, 0x00), RGB(0xF4, 0xB4, 0x1F),
        RGB(0x43, 0xA0, 0x47), RGB(0x1E, 0x88, 0xE5), RGB(0x5E, 0x35, 0xB1), RGB(0x8E, 0x24, 0xAA),
    };

    std::wstring lower = ToLowerCopy(tag);
    std::hash<std::wstring> hashFn;
    const size_t index = hashFn(lower) % std::size(kPalette);
    return kPalette[index];
}

COLORREF TagStore::BlendColors(COLORREF a, COLORREF b) {
    const int r = (GetRValue(a) + GetRValue(b)) / 2;
    const int g = (GetGValue(a) + GetGValue(b)) / 2;
    const int bl = (GetBValue(a) + GetBValue(b)) / 2;
    return RGB(r, g, bl);
}

}  // namespace shelltabs

