#include "OptionsStore.h"

#include <ShlObj.h>
#include <KnownFolders.h>
#include <Shlwapi.h>

#include <algorithm>
#include <cstdlib>
#include <cwchar>
#include <sstream>
#include <string>
#include <vector>

namespace shelltabs {
namespace {
constexpr wchar_t kStorageDirectory[] = L"ShellTabs";
constexpr wchar_t kStorageFile[] = L"options.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kReopenToken[] = L"reopen_on_crash";
constexpr wchar_t kPersistToken[] = L"persist_group_paths";
constexpr wchar_t kBreadcrumbGradientToken[] = L"breadcrumb_gradient";
constexpr wchar_t kBreadcrumbFontGradientToken[] = L"breadcrumb_font_gradient";
constexpr wchar_t kBreadcrumbGradientTransparencyToken[] = L"breadcrumb_gradient_transparency";
constexpr wchar_t kBreadcrumbFontBrightnessToken[] = L"breadcrumb_font_brightness";
constexpr wchar_t kBreadcrumbFontTransparencyToken[] = L"breadcrumb_font_transparency";  // legacy
constexpr wchar_t kBreadcrumbGradientColorsToken[] = L"breadcrumb_gradient_colors";
constexpr wchar_t kBreadcrumbFontGradientColorsToken[] = L"breadcrumb_font_gradient_colors";
constexpr wchar_t kTabSelectedColorToken[] = L"tab_selected_color";
constexpr wchar_t kTabUnselectedColorToken[] = L"tab_unselected_color";
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
    std::wstring wide(length, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.data(), static_cast<int>(utf8.size()), wide.data(), length);
    return wide;
}

std::string WideToUtf8(const std::wstring& wide) {
    if (wide.empty()) {
        return {};
    }
    const int length = WideCharToMultiByte(
        CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, nullptr, nullptr);
    if (length <= 0) {
        return {};
    }
    std::string utf8(length, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), utf8.data(), length, nullptr, nullptr);
    return utf8;
}

std::wstring ResolveDirectory() {
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

bool ParseBool(const std::wstring& token) {
    return token == L"1" || token == L"true" || token == L"TRUE" || token == L"yes" || token == L"on";
}

int ParseIntInRange(const std::wstring& token, int minimum, int maximum, int fallback) {
    if (token.empty()) {
        return fallback;
    }
    int value = _wtoi(token.c_str());
    if (value < minimum) {
        return minimum;
    }
    if (value > maximum) {
        return maximum;
    }
    return value;
}

COLORREF ParseColorValue(const std::wstring& token, COLORREF fallback) {
    if (token.empty()) {
        return fallback;
    }

    unsigned int parsed = 0;
    if (swscanf_s(token.c_str(), L"%x", &parsed) != 1) {
        return fallback;
    }

    return RGB((parsed >> 16) & 0xFF, (parsed >> 8) & 0xFF, parsed & 0xFF);
}

std::wstring ColorToHexString(COLORREF color) {
    std::wostringstream stream;
    stream << std::uppercase << std::hex;
    stream.width(6);
    stream.fill(L'0');
    const unsigned int packed = (static_cast<unsigned int>(GetRValue(color)) << 16) |
                                (static_cast<unsigned int>(GetGValue(color)) << 8) |
                                static_cast<unsigned int>(GetBValue(color));
    stream << packed;
    return stream.str();
}

}  // namespace

OptionsStore& OptionsStore::Instance() {
    static OptionsStore store;
    return store;
}

bool OptionsStore::EnsureLoaded() const {
    if (m_loaded) {
        return true;
    }
    return const_cast<OptionsStore*>(this)->Load();
}

std::wstring OptionsStore::ResolveStoragePath() const {
    if (!m_storagePath.empty()) {
        return m_storagePath;
    }

    std::wstring directory = ResolveDirectory();
    if (directory.empty()) {
        return {};
    }

    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += kStorageFile;
    return directory;
}

bool OptionsStore::Load() {
    m_loaded = true;
    m_options = {};

    m_storagePath = ResolveStoragePath();
    if (m_storagePath.empty()) {
        return false;
    }

    HANDLE file = CreateFileW(m_storagePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return true;
    }

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file, &size) || size.HighPart != 0) {
        CloseHandle(file);
        return false;
    }

    std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
    DWORD bytesRead = 0;
    if (!buffer.empty() &&
        !ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr)) {
        CloseHandle(file);
        return false;
    }
    CloseHandle(file);
    buffer.resize(bytesRead);

    const std::wstring content = Utf8ToWide(buffer);
    if (content.empty()) {
        return true;
    }

    int version = 1;
    size_t lineStart = 0;
    while (lineStart < content.size()) {
        size_t lineEnd = content.find(L'\n', lineStart);
        std::wstring line = content.substr(lineStart,
                                           lineEnd == std::wstring::npos ? std::wstring::npos : lineEnd - lineStart);
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

        if (tokens[0] == kReopenToken) {
            if (tokens.size() >= 2) {
                m_options.reopenOnCrash = ParseBool(tokens[1]);
            }
            continue;
        }

        if (tokens[0] == kPersistToken) {
            if (tokens.size() >= 2) {
                m_options.persistGroupPaths = ParseBool(tokens[1]);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbGradientToken) {
            if (tokens.size() >= 2) {
                m_options.enableBreadcrumbGradient = ParseBool(tokens[1]);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbFontGradientToken) {
            if (tokens.size() >= 2) {
                m_options.enableBreadcrumbFontGradient = ParseBool(tokens[1]);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbGradientTransparencyToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbGradientTransparency =
                    ParseIntInRange(tokens[1], 0, 100, m_options.breadcrumbGradientTransparency);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbFontBrightnessToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbFontBrightness =
                    ParseIntInRange(tokens[1], 0, 100, m_options.breadcrumbFontBrightness);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbFontTransparencyToken) {
            if (tokens.size() >= 2) {
                const int defaultBrightness = m_options.breadcrumbFontBrightness;
                const int legacyTransparency =
                    ParseIntInRange(tokens[1], 0, 100, 100 - defaultBrightness);
                const int legacyOpacity = 100 - legacyTransparency;
                m_options.breadcrumbFontBrightness =
                    std::clamp(legacyOpacity * defaultBrightness / 100, 0, 100);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbGradientColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomBreadcrumbGradientColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.breadcrumbGradientStartColor =
                    ParseColorValue(tokens[2], m_options.breadcrumbGradientStartColor);
            }
            if (tokens.size() >= 4) {
                m_options.breadcrumbGradientEndColor =
                    ParseColorValue(tokens[3], m_options.breadcrumbGradientEndColor);
            }
            continue;
        }

        if (tokens[0] == kBreadcrumbFontGradientColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomBreadcrumbFontColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.breadcrumbFontGradientStartColor =
                    ParseColorValue(tokens[2], m_options.breadcrumbFontGradientStartColor);
            }
            if (tokens.size() >= 4) {
                m_options.breadcrumbFontGradientEndColor =
                    ParseColorValue(tokens[3], m_options.breadcrumbFontGradientEndColor);
            }
            continue;
        }

        if (tokens[0] == kTabSelectedColorToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomTabSelectedColor = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.customTabSelectedColor =
                    ParseColorValue(tokens[2], m_options.customTabSelectedColor);
            }
            continue;
        }

        if (tokens[0] == kTabUnselectedColorToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomTabUnselectedColor = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.customTabUnselectedColor =
                    ParseColorValue(tokens[2], m_options.customTabUnselectedColor);
            }
            continue;
        }
    }

    return true;
}

bool OptionsStore::Save() const {
    if (!EnsureLoaded()) {
        return false;
    }

    const_cast<OptionsStore*>(this)->m_storagePath = ResolveStoragePath();
    if (m_storagePath.empty()) {
        return false;
    }

    std::wstring content = L"version|1\n";
    content += kReopenToken;
    content += L"|";
    content += m_options.reopenOnCrash ? L"1" : L"0";
    content += L"\n";
    content += kPersistToken;
    content += L"|";
    content += m_options.persistGroupPaths ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbGradientToken;
    content += L"|";
    content += m_options.enableBreadcrumbGradient ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbFontGradientToken;
    content += L"|";
    content += m_options.enableBreadcrumbFontGradient ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbGradientTransparencyToken;
    content += L"|";
    content += std::to_wstring(std::clamp(m_options.breadcrumbGradientTransparency, 0, 100));
    content += L"\n";
    content += kBreadcrumbFontBrightnessToken;
    content += L"|";
    content += std::to_wstring(std::clamp(m_options.breadcrumbFontBrightness, 0, 100));
    content += L"\n";
    content += kBreadcrumbGradientColorsToken;
    content += L"|";
    content += m_options.useCustomBreadcrumbGradientColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(m_options.breadcrumbGradientStartColor);
    content += L"|";
    content += ColorToHexString(m_options.breadcrumbGradientEndColor);
    content += L"\n";
    content += kBreadcrumbFontGradientColorsToken;
    content += L"|";
    content += m_options.useCustomBreadcrumbFontColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(m_options.breadcrumbFontGradientStartColor);
    content += L"|";
    content += ColorToHexString(m_options.breadcrumbFontGradientEndColor);
    content += L"\n";
    content += kTabSelectedColorToken;
    content += L"|";
    content += m_options.useCustomTabSelectedColor ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(m_options.customTabSelectedColor);
    content += L"\n";
    content += kTabUnselectedColorToken;
    content += L"|";
    content += m_options.useCustomTabUnselectedColor ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(m_options.customTabUnselectedColor);
    content += L"\n";

    const std::string utf8 = WideToUtf8(content);

    HANDLE file = CreateFileW(m_storagePath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL,
                               nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }

    DWORD bytesWritten = 0;
    const BOOL result = WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &bytesWritten, nullptr);
    CloseHandle(file);
    return result != FALSE;
}

bool operator==(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept {
    return left.reopenOnCrash == right.reopenOnCrash && left.persistGroupPaths == right.persistGroupPaths &&
           left.enableBreadcrumbGradient == right.enableBreadcrumbGradient &&
           left.enableBreadcrumbFontGradient == right.enableBreadcrumbFontGradient &&
           left.breadcrumbGradientTransparency == right.breadcrumbGradientTransparency &&
           left.breadcrumbFontBrightness == right.breadcrumbFontBrightness &&
           left.useCustomBreadcrumbGradientColors == right.useCustomBreadcrumbGradientColors &&
           left.breadcrumbGradientStartColor == right.breadcrumbGradientStartColor &&
           left.breadcrumbGradientEndColor == right.breadcrumbGradientEndColor &&
           left.useCustomBreadcrumbFontColors == right.useCustomBreadcrumbFontColors &&
           left.breadcrumbFontGradientStartColor == right.breadcrumbFontGradientStartColor &&
           left.breadcrumbFontGradientEndColor == right.breadcrumbFontGradientEndColor &&
           left.useCustomTabSelectedColor == right.useCustomTabSelectedColor &&
           left.customTabSelectedColor == right.customTabSelectedColor &&
           left.useCustomTabUnselectedColor == right.useCustomTabUnselectedColor &&
           left.customTabUnselectedColor == right.customTabUnselectedColor;
}

}  // namespace shelltabs

