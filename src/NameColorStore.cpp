#include "NameColorStore.h"

#include <algorithm>
#include <cwctype>
#include <mutex>

#include "Logging.h"
#include "Utilities.h"

namespace shelltabs {

namespace {

constexpr wchar_t kStorageFile[] = L"namecolors.json";

std::wstring ToLowerCopy(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return value;
}

std::wstring NormalizePathKey(const std::wstring& path) {
    std::wstring key = ToLowerCopy(path);
    // Keep drive roots ("c:\") distinguishable but drop other trailing separators.
    while (key.size() > 3 && (key.back() == L'\\' || key.back() == L'/')) {
        key.pop_back();
    }
    if (key.size() == 3 && key[1] == L':' && key[2] == L'\\') {
        key.pop_back();
    }
    return key;
}

std::wstring FormatColor(COLORREF color) {
    wchar_t buffer[8];
    swprintf(buffer, ARRAYSIZE(buffer), L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
    return buffer;
}

std::optional<COLORREF> ParseColor(const std::wstring& text) {
    if (text.size() != 7 || text[0] != L'#') {
        return std::nullopt;
    }
    unsigned int value = 0;
    for (size_t i = 1; i < text.size(); ++i) {
        const wchar_t ch = text[i];
        value <<= 4;
        if (ch >= L'0' && ch <= L'9') {
            value |= static_cast<unsigned int>(ch - L'0');
        } else if (ch >= L'a' && ch <= L'f') {
            value |= static_cast<unsigned int>(ch - L'a' + 10);
        } else if (ch >= L'A' && ch <= L'F') {
            value |= static_cast<unsigned int>(ch - L'A' + 10);
        } else {
            return std::nullopt;
        }
    }
    return RGB((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
}

std::wstring EscapeJson(const std::wstring& value) {
    std::wstring out;
    out.reserve(value.size() + 8);
    for (wchar_t ch : value) {
        switch (ch) {
            case L'\\': out += L"\\\\"; break;
            case L'"': out += L"\\\""; break;
            case L'\n': out += L"\\n"; break;
            case L'\r': out += L"\\r"; break;
            case L'\t': out += L"\\t"; break;
            default:
                if (ch < 0x20) {
                    wchar_t buffer[8];
                    swprintf(buffer, ARRAYSIZE(buffer), L"\\u%04X", static_cast<unsigned int>(ch));
                    out += buffer;
                } else {
                    out.push_back(ch);
                }
                break;
        }
    }
    return out;
}

// Minimal parser for a flat JSON object of string -> string.
bool ParseFlatJsonObject(const std::wstring& text, std::vector<std::pair<std::wstring, std::wstring>>* out) {
    size_t pos = 0;
    auto skipWs = [&]() {
        while (pos < text.size() && std::iswspace(text[pos])) {
            ++pos;
        }
    };
    auto parseString = [&](std::wstring* value) -> bool {
        if (pos >= text.size() || text[pos] != L'"') {
            return false;
        }
        ++pos;
        value->clear();
        while (pos < text.size()) {
            wchar_t ch = text[pos++];
            if (ch == L'"') {
                return true;
            }
            if (ch != L'\\') {
                value->push_back(ch);
                continue;
            }
            if (pos >= text.size()) {
                return false;
            }
            const wchar_t esc = text[pos++];
            switch (esc) {
                case L'"': value->push_back(L'"'); break;
                case L'\\': value->push_back(L'\\'); break;
                case L'/': value->push_back(L'/'); break;
                case L'b': value->push_back(L'\b'); break;
                case L'f': value->push_back(L'\f'); break;
                case L'n': value->push_back(L'\n'); break;
                case L'r': value->push_back(L'\r'); break;
                case L't': value->push_back(L'\t'); break;
                case L'u': {
                    if (pos + 4 > text.size()) {
                        return false;
                    }
                    unsigned int code = 0;
                    for (int i = 0; i < 4; ++i) {
                        const wchar_t h = text[pos++];
                        code <<= 4;
                        if (h >= L'0' && h <= L'9') code |= static_cast<unsigned int>(h - L'0');
                        else if (h >= L'a' && h <= L'f') code |= static_cast<unsigned int>(h - L'a' + 10);
                        else if (h >= L'A' && h <= L'F') code |= static_cast<unsigned int>(h - L'A' + 10);
                        else return false;
                    }
                    value->push_back(static_cast<wchar_t>(code));
                    break;
                }
                default:
                    return false;
            }
        }
        return false;
    };

    // Skip UTF-8 BOM decoded as U+FEFF.
    if (!text.empty() && text[0] == 0xFEFF) {
        pos = 1;
    }
    skipWs();
    if (pos >= text.size() || text[pos] != L'{') {
        return false;
    }
    ++pos;
    skipWs();
    if (pos < text.size() && text[pos] == L'}') {
        return true;
    }
    while (pos < text.size()) {
        std::wstring key;
        std::wstring value;
        skipWs();
        if (!parseString(&key)) {
            return false;
        }
        skipWs();
        if (pos >= text.size() || text[pos] != L':') {
            return false;
        }
        ++pos;
        skipWs();
        if (!parseString(&value)) {
            return false;
        }
        out->emplace_back(std::move(key), std::move(value));
        skipWs();
        if (pos < text.size() && text[pos] == L',') {
            ++pos;
            continue;
        }
        if (pos < text.size() && text[pos] == L'}') {
            return true;
        }
        return false;
    }
    return false;
}

bool IsExtensionKey(const std::wstring& key) {
    return key.size() > 2 && key[0] == L'*' && key[1] == L'.';
}

}  // namespace

NameColorStore& NameColorStore::Instance() {
    static NameColorStore instance;
    return instance;
}

NameColorStore::NameColorStore() {
    std::unique_lock lock(m_mutex);
    LoadLocked();
}

std::wstring NameColorStore::ResolveStoragePath() const {
    std::wstring base = GetShellTabsDataDirectory();
    if (base.empty()) {
        return {};
    }
    if (base.back() != L'\\') {
        base.push_back(L'\\');
    }
    return base + kStorageFile;
}

void NameColorStore::LoadLocked() {
    m_colors.clear();
    const std::wstring path = ResolveStoragePath();
    if (path.empty()) {
        return;
    }
    std::wstring content;
    bool exists = false;
    if (!ReadUtf8File(path, &content, &exists) || content.empty()) {
        return;
    }
    std::vector<std::pair<std::wstring, std::wstring>> entries;
    if (!ParseFlatJsonObject(content, &entries)) {
        LogMessage(LogLevel::Warning, L"NameColorStore: failed to parse %ls", path.c_str());
        return;
    }
    for (const auto& [key, value] : entries) {
        auto color = ParseColor(value);
        if (!color || key.empty()) {
            continue;
        }
        if (IsExtensionKey(key)) {
            const std::wstring normalized = NormalizeExtensionKey(key);
            if (!normalized.empty()) {
                m_colors[normalized] = *color;
            }
        } else {
            m_colors[NormalizePathKey(key)] = *color;
        }
    }
}

bool NameColorStore::SaveLocked() const {
    const std::wstring path = ResolveStoragePath();
    if (path.empty()) {
        return false;
    }
    std::vector<std::pair<std::wstring, COLORREF>> sorted(m_colors.begin(), m_colors.end());
    std::sort(sorted.begin(), sorted.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
    std::wstring json = L"{";
    bool first = true;
    for (const auto& [key, color] : sorted) {
        if (!first) {
            json += L",";
        }
        first = false;
        json += L"\n  \"";
        json += EscapeJson(key);
        json += L"\": \"";
        json += FormatColor(color);
        json += L"\"";
    }
    json += sorted.empty() ? L"}\n" : L"\n}\n";
    if (!WriteUtf8File(path, json)) {
        LogMessage(LogLevel::Warning, L"NameColorStore: failed to write %ls", path.c_str());
        return false;
    }
    return true;
}

void NameColorStore::Reload() {
    std::unique_lock lock(m_mutex);
    LoadLocked();
}

bool NameColorStore::IsEmpty() const {
    std::shared_lock lock(m_mutex);
    return m_colors.empty();
}

std::wstring NameColorStore::NormalizeExtensionKey(const std::wstring& extension) {
    std::wstring value = extension;
    // Trim whitespace.
    while (!value.empty() && std::iswspace(value.front())) value.erase(value.begin());
    while (!value.empty() && std::iswspace(value.back())) value.pop_back();
    if (value.rfind(L"*.", 0) == 0) {
        value.erase(0, 2);
    } else if (!value.empty() && value.front() == L'.') {
        value.erase(0, 1);
    }
    if (value.empty() || value.find_first_of(L"\\/:*?\"<>|. ") != std::wstring::npos) {
        return {};
    }
    return L"*." + ToLowerCopy(value);
}

std::wstring NameColorStore::ExtensionFromPath(const std::wstring& path) {
    const size_t slash = path.find_last_of(L"\\/");
    const size_t dot = path.find_last_of(L'.');
    if (dot == std::wstring::npos || (slash != std::wstring::npos && dot < slash) || dot + 1 >= path.size()) {
        return {};
    }
    // Hidden-file style names (".gitignore") use the whole name as the extension, like Explorer.
    return path.substr(dot + 1);
}

std::optional<COLORREF> NameColorStore::Lookup(const std::wstring& parsingPath, bool isFolder) const {
    if (parsingPath.empty()) {
        return std::nullopt;
    }
    std::shared_lock lock(m_mutex);
    if (m_colors.empty()) {
        return std::nullopt;
    }
    auto it = m_colors.find(NormalizePathKey(parsingPath));
    if (it != m_colors.end()) {
        return it->second;
    }
    if (isFolder) {
        return std::nullopt;
    }
    const std::wstring extension = ExtensionFromPath(parsingPath);
    if (extension.empty()) {
        return std::nullopt;
    }
    it = m_colors.find(L"*." + ToLowerCopy(extension));
    if (it != m_colors.end()) {
        return it->second;
    }
    return std::nullopt;
}

std::optional<COLORREF> NameColorStore::GetPathColor(const std::wstring& parsingPath) const {
    std::shared_lock lock(m_mutex);
    auto it = m_colors.find(NormalizePathKey(parsingPath));
    if (it == m_colors.end()) {
        return std::nullopt;
    }
    return it->second;
}

std::optional<COLORREF> NameColorStore::GetExtensionColor(const std::wstring& extension) const {
    const std::wstring key = NormalizeExtensionKey(extension);
    if (key.empty()) {
        return std::nullopt;
    }
    std::shared_lock lock(m_mutex);
    auto it = m_colors.find(key);
    if (it == m_colors.end()) {
        return std::nullopt;
    }
    return it->second;
}

std::vector<std::pair<std::wstring, COLORREF>> NameColorStore::GetExtensionColors() const {
    std::vector<std::pair<std::wstring, COLORREF>> result;
    {
        std::shared_lock lock(m_mutex);
        for (const auto& [key, color] : m_colors) {
            if (IsExtensionKey(key)) {
                result.emplace_back(key.substr(2), color);
            }
        }
    }
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
    return result;
}


bool NameColorStore::SetPathColors(const std::vector<std::wstring>& parsingPaths, std::optional<COLORREF> color) {
    bool saved = false;
    {
        std::unique_lock lock(m_mutex);
        // Pick up edits made by other Explorer processes before writing.
        LoadLocked();
        for (const auto& path : parsingPaths) {
            if (path.empty()) {
                continue;
            }
            const std::wstring key = NormalizePathKey(path);
            if (color) {
                m_colors[key] = *color;
            } else {
                m_colors.erase(key);
            }
        }
        saved = SaveLocked();
    }
    if (saved) {
        BroadcastChanged();
    }
    return saved;
}

bool NameColorStore::SetExtensionColor(const std::wstring& extension, std::optional<COLORREF> color) {
    const std::wstring key = NormalizeExtensionKey(extension);
    if (key.empty()) {
        return false;
    }
    bool saved = false;
    {
        std::unique_lock lock(m_mutex);
        LoadLocked();
        if (color) {
            m_colors[key] = *color;
        } else {
            m_colors.erase(key);
        }
        saved = SaveLocked();
    }
    if (saved) {
        BroadcastChanged();
    }
    return saved;
}


UINT NameColorStore::ChangedMessage() {
    static const UINT message = RegisterWindowMessageW(L"ShellTabs.NameColorsChanged");
    return message;
}

void NameColorStore::BroadcastChanged() {
    const UINT message = ChangedMessage();
    if (!message) {
        return;
    }
    HWND frame = nullptr;
    while ((frame = FindWindowExW(nullptr, frame, L"CabinetWClass", nullptr)) != nullptr) {
        PostMessageW(frame, message, 0, 0);
    }
}

}  // namespace shelltabs
