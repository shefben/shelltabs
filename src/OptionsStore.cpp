#include "OptionsStore.h"

#include "BackgroundCache.h"
#include "StringUtils.h"
#include "Utilities.h"

#include <Shlwapi.h>

#include <algorithm>
#include <cstdlib>
#include <cwchar>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

namespace shelltabs {
namespace {
constexpr wchar_t kStorageFile[] = L"options.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kReopenToken[] = L"reopen_on_crash";
constexpr wchar_t kPersistToken[] = L"persist_group_paths";
constexpr wchar_t kNewTabTemplateToken[] = L"new_tab_template";
constexpr wchar_t kNewTabCustomPathToken[] = L"new_tab_custom_path";
constexpr wchar_t kNewTabSavedGroupToken[] = L"new_tab_saved_group";
constexpr wchar_t kBreadcrumbGradientToken[] = L"breadcrumb_gradient";
constexpr wchar_t kBreadcrumbFontGradientToken[] = L"breadcrumb_font_gradient";
constexpr wchar_t kBreadcrumbGradientTransparencyToken[] = L"breadcrumb_gradient_transparency";
constexpr wchar_t kBreadcrumbFontBrightnessToken[] = L"breadcrumb_font_brightness";
constexpr wchar_t kBreadcrumbFontTransparencyToken[] = L"breadcrumb_font_transparency";  // legacy
constexpr wchar_t kBreadcrumbHighlightAlphaMultiplierToken[] =
    L"breadcrumb_highlight_alpha_multiplier";
constexpr wchar_t kBreadcrumbDropdownAlphaMultiplierToken[] =
    L"breadcrumb_dropdown_alpha_multiplier";
constexpr wchar_t kBreadcrumbGradientColorsToken[] = L"breadcrumb_gradient_colors";
constexpr wchar_t kBreadcrumbFontGradientColorsToken[] = L"breadcrumb_font_gradient_colors";
constexpr wchar_t kProgressGradientColorsToken[] = L"progress_gradient_colors";
constexpr wchar_t kTabSelectedColorToken[] = L"tab_selected_color";
constexpr wchar_t kTabUnselectedColorToken[] = L"tab_unselected_color";
constexpr wchar_t kFolderBackgroundsEnabledToken[] = L"folder_backgrounds_enabled";
constexpr wchar_t kFolderBackgroundUniversalToken[] = L"folder_background_universal";
constexpr wchar_t kFolderBackgroundEntryToken[] = L"folder_background_entry";
constexpr wchar_t kTabDockingToken[] = L"tab_docking";
constexpr wchar_t kCommentChar = L'#';

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

bool HasDirectoryPrefix(const std::wstring& path, const std::wstring& directory) {
    if (path.size() < directory.size()) {
        return false;
    }
    if (_wcsnicmp(path.c_str(), directory.c_str(), directory.size()) != 0) {
        return false;
    }
    if (path.size() == directory.size()) {
        return true;
    }
    const wchar_t separator = path[directory.size()];
    return separator == L'\\';
}

std::wstring NormalizeCachePath(const std::wstring& path, const std::wstring& directory) {
    if (path.empty() || directory.empty()) {
        return {};
    }

    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }

    if (!HasDirectoryPrefix(normalized, directory)) {
        return {};
    }

    return normalized;
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

    std::wstring directory = GetShellTabsDataDirectory();
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
    m_options = {};

    m_storagePath = ResolveStoragePath();
    if (m_storagePath.empty()) {
        m_loaded = false;
        return false;
    }

    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(m_storagePath, &content, &fileExists)) {
        m_loaded = false;
        return false;
    }

    const std::wstring storageDirectory = GetShellTabsDataDirectory();

    if (!fileExists || content.empty()) {
        m_loaded = true;
        return true;
    }

    int version = 1;
    ParseConfigLines(content, kCommentChar, L'|', [&](const std::vector<std::wstring>& tokens) {
        if (tokens.empty()) {
            return true;
        }

        const std::wstring& header = tokens[0];
        if (header == kVersionToken) {
            if (tokens.size() >= 2) {
                version = _wtoi(tokens[1].c_str());
                if (version < 1) {
                    version = 1;
                }
            }
            return true;
        }

        if (header == kReopenToken) {
            if (tokens.size() >= 2) {
                m_options.reopenOnCrash = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kPersistToken) {
            if (tokens.size() >= 2) {
                m_options.persistGroupPaths = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kNewTabTemplateToken) {
            if (tokens.size() >= 2) {
                m_options.newTabTemplate = ParseNewTabTemplate(tokens[1]);
            }
            return true;
        }

        if (header == kNewTabCustomPathToken) {
            if (tokens.size() >= 2) {
                m_options.newTabCustomPath = Trim(tokens[1]);
            }
            return true;
        }

        if (header == kNewTabSavedGroupToken) {
            if (tokens.size() >= 2) {
                m_options.newTabSavedGroup = Trim(tokens[1]);
            }
            return true;
        }

        if (header == kBreadcrumbGradientToken) {
            if (tokens.size() >= 2) {
                m_options.enableBreadcrumbGradient = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kBreadcrumbFontGradientToken) {
            if (tokens.size() >= 2) {
                m_options.enableBreadcrumbFontGradient = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kBreadcrumbGradientTransparencyToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbGradientTransparency =
                    ParseIntInRange(tokens[1], 0, 100, m_options.breadcrumbGradientTransparency);
            }
            return true;
        }

        if (header == kBreadcrumbFontBrightnessToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbFontBrightness =
                    ParseIntInRange(tokens[1], 0, 100, m_options.breadcrumbFontBrightness);
            }
            return true;
        }

        if (header == kBreadcrumbHighlightAlphaMultiplierToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbHighlightAlphaMultiplier =
                    ParseIntInRange(tokens[1], 0, 200, m_options.breadcrumbHighlightAlphaMultiplier);
            }
            return true;
        }

        if (header == kBreadcrumbDropdownAlphaMultiplierToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbDropdownAlphaMultiplier =
                    ParseIntInRange(tokens[1], 0, 200, m_options.breadcrumbDropdownAlphaMultiplier);
            }
            return true;
        }

        if (header == kBreadcrumbFontTransparencyToken) {
            if (tokens.size() >= 2) {
                const int defaultBrightness = m_options.breadcrumbFontBrightness;
                const int legacyTransparency =
                    ParseIntInRange(tokens[1], 0, 100, 100 - defaultBrightness);
                const int legacyOpacity = 100 - legacyTransparency;
                m_options.breadcrumbFontBrightness =
                    std::clamp(legacyOpacity * defaultBrightness / 100, 0, 100);
            }
            return true;
        }

        if (header == kBreadcrumbGradientColorsToken) {
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
            return true;
        }

        if (header == kBreadcrumbFontGradientColorsToken) {
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
            return true;
        }

        if (header == kProgressGradientColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomProgressBarGradientColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.progressBarGradientStartColor =
                    ParseColorValue(tokens[2], m_options.progressBarGradientStartColor);
            }
            if (tokens.size() >= 4) {
                m_options.progressBarGradientEndColor =
                    ParseColorValue(tokens[3], m_options.progressBarGradientEndColor);
            }
            return true;
        }

        if (header == kTabSelectedColorToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomTabSelectedColor = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.customTabSelectedColor =
                    ParseColorValue(tokens[2], m_options.customTabSelectedColor);
            }
            return true;
        }

        if (header == kTabUnselectedColorToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomTabUnselectedColor = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.customTabUnselectedColor =
                    ParseColorValue(tokens[2], m_options.customTabUnselectedColor);
            }
            return true;
        }

        if (header == kFolderBackgroundsEnabledToken) {
            if (tokens.size() >= 2) {
                m_options.enableFolderBackgrounds = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kFolderBackgroundUniversalToken) {
            if (tokens.size() >= 2) {
                const std::wstring cachePath = NormalizeCachePath(tokens[1], storageDirectory);
                if (!cachePath.empty()) {
                    m_options.universalFolderBackgroundImage.cachedImagePath = cachePath;
                    if (tokens.size() >= 3) {
                        m_options.universalFolderBackgroundImage.displayName = tokens[2];
                    }
                }
            }
            return true;
        }

        if (header == kFolderBackgroundEntryToken) {
            if (tokens.size() >= 3) {
                FolderBackgroundEntry entry;
                entry.folderPath = NormalizeFileSystemPath(tokens[1]);
                if (!entry.folderPath.empty()) {
                    entry.image.cachedImagePath = NormalizeCachePath(tokens[2], storageDirectory);
                    if (!entry.image.cachedImagePath.empty()) {
                        if (tokens.size() >= 4) {
                            entry.image.displayName = tokens[3];
                        }
                        m_options.folderBackgroundEntries.emplace_back(std::move(entry));
                    }
                }
            }
            return true;
        }

        if (header == kTabDockingToken) {
            if (tokens.size() >= 2) {
                m_options.tabDockMode = ParseDockMode(tokens[1]);
            }
            return true;
        }

        return true;
    });

    m_loaded = true;
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

    const std::wstring storageDirectory = GetShellTabsDataDirectory();
    if (storageDirectory.empty()) {
        return false;
    }

    const size_t separator = m_storagePath.find_last_of(L"\\/");
    if (separator != std::wstring::npos) {
        std::wstring directory = m_storagePath.substr(0, separator);
        if (!directory.empty()) {
            CreateDirectoryW(directory.c_str(), nullptr);
        }
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
    content += kNewTabTemplateToken;
    content += L"|";
    content += NewTabTemplateToString(m_options.newTabTemplate);
    content += L"\n";
    content += kNewTabCustomPathToken;
    content += L"|";
    content += Trim(m_options.newTabCustomPath);
    content += L"\n";
    content += kNewTabSavedGroupToken;
    content += L"|";
    content += Trim(m_options.newTabSavedGroup);
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
    content += kBreadcrumbHighlightAlphaMultiplierToken;
    content += L"|";
    content += std::to_wstring(std::clamp(m_options.breadcrumbHighlightAlphaMultiplier, 0, 200));
    content += L"\n";
    content += kBreadcrumbDropdownAlphaMultiplierToken;
    content += L"|";
    content += std::to_wstring(std::clamp(m_options.breadcrumbDropdownAlphaMultiplier, 0, 200));
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
    content += kProgressGradientColorsToken;
    content += L"|";
    content += m_options.useCustomProgressBarGradientColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(m_options.progressBarGradientStartColor);
    content += L"|";
    content += ColorToHexString(m_options.progressBarGradientEndColor);
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

    content += kFolderBackgroundsEnabledToken;
    content += L"|";
    content += m_options.enableFolderBackgrounds ? L"1" : L"0";
    content += L"\n";

    const auto appendCachedImageLine = [&](const wchar_t* token, const std::wstring& path,
                                           const std::wstring& displayName) {
        const std::wstring normalizedPath = NormalizeCachePath(path, storageDirectory);
        if (normalizedPath.empty()) {
            return;
        }

        content += token;
        content += L"|";
        content += normalizedPath;
        content += L"|";
        content += displayName;
        content += L"\n";
    };

    appendCachedImageLine(kFolderBackgroundUniversalToken, m_options.universalFolderBackgroundImage.cachedImagePath,
                          m_options.universalFolderBackgroundImage.displayName);

    for (const auto& entry : m_options.folderBackgroundEntries) {
        const std::wstring normalizedFolder = NormalizeFileSystemPath(entry.folderPath);
        const std::wstring normalizedCache =
            NormalizeCachePath(entry.image.cachedImagePath, storageDirectory);
        if (normalizedFolder.empty() || normalizedCache.empty()) {
            continue;
        }

        content += kFolderBackgroundEntryToken;
        content += L"|";
        content += normalizedFolder;
        content += L"|";
        content += normalizedCache;
        content += L"|";
        content += entry.image.displayName;
        content += L"\n";
    }

    content += kTabDockingToken;
    content += L"|";
    content += DockModeToString(m_options.tabDockMode);
    content += L"\n";

    if (!WriteUtf8File(m_storagePath, content)) {
        return false;
    }

    UpdateCachedImageUsage(m_options);
    return true;
}

bool operator==(const CachedImageMetadata& left, const CachedImageMetadata& right) noexcept {
    return left.cachedImagePath == right.cachedImagePath && left.displayName == right.displayName;
}

bool operator==(const FolderBackgroundEntry& left, const FolderBackgroundEntry& right) noexcept {
    return left.folderPath == right.folderPath && left.image == right.image;
}

bool operator==(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept {
    return left.reopenOnCrash == right.reopenOnCrash && left.persistGroupPaths == right.persistGroupPaths &&
           left.enableBreadcrumbGradient == right.enableBreadcrumbGradient &&
           left.enableBreadcrumbFontGradient == right.enableBreadcrumbFontGradient &&
           left.breadcrumbGradientTransparency == right.breadcrumbGradientTransparency &&
           left.breadcrumbFontBrightness == right.breadcrumbFontBrightness &&
           left.breadcrumbHighlightAlphaMultiplier == right.breadcrumbHighlightAlphaMultiplier &&
           left.breadcrumbDropdownAlphaMultiplier == right.breadcrumbDropdownAlphaMultiplier &&
           left.useCustomBreadcrumbGradientColors == right.useCustomBreadcrumbGradientColors &&
           left.breadcrumbGradientStartColor == right.breadcrumbGradientStartColor &&
           left.breadcrumbGradientEndColor == right.breadcrumbGradientEndColor &&
           left.useCustomBreadcrumbFontColors == right.useCustomBreadcrumbFontColors &&
           left.breadcrumbFontGradientStartColor == right.breadcrumbFontGradientStartColor &&
           left.breadcrumbFontGradientEndColor == right.breadcrumbFontGradientEndColor &&
           left.useCustomProgressBarGradientColors == right.useCustomProgressBarGradientColors &&
           left.progressBarGradientStartColor == right.progressBarGradientStartColor &&
           left.progressBarGradientEndColor == right.progressBarGradientEndColor &&
           left.useCustomTabSelectedColor == right.useCustomTabSelectedColor &&
           left.customTabSelectedColor == right.customTabSelectedColor &&
           left.useCustomTabUnselectedColor == right.useCustomTabUnselectedColor &&
           left.customTabUnselectedColor == right.customTabUnselectedColor &&
           left.enableFolderBackgrounds == right.enableFolderBackgrounds &&
           left.universalFolderBackgroundImage == right.universalFolderBackgroundImage &&
           left.folderBackgroundEntries == right.folderBackgroundEntries &&
           left.tabDockMode == right.tabDockMode &&
           left.newTabTemplate == right.newTabTemplate &&
           left.newTabCustomPath == right.newTabCustomPath &&
           left.newTabSavedGroup == right.newTabSavedGroup;
}

}  // namespace shelltabs

