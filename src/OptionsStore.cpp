#include "OptionsStore.h"

#include "BackgroundCache.h"
#include "StringUtils.h"
#include "Utilities.h"

#include <Shlwapi.h>

#include <algorithm>
#include <array>
#include <cstdlib>
#include <cwchar>
#include <sstream>
#include <string>
#include <string_view>
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
constexpr wchar_t kGlowEnabledToken[] = L"neon_glow_enabled";
constexpr wchar_t kGlowGradientToken[] = L"neon_glow_gradient";
constexpr wchar_t kGlowColorsToken[] = L"neon_glow_colors";
constexpr wchar_t kGlowSurfaceToken[] = L"glow_surface";
constexpr wchar_t kTabSelectedColorToken[] = L"tab_selected_color";
constexpr wchar_t kTabUnselectedColorToken[] = L"tab_unselected_color";
constexpr wchar_t kExplorerAccentColorsToken[] = L"explorer_listview_accents";
constexpr wchar_t kFolderBackgroundsEnabledToken[] = L"folder_backgrounds_enabled";
constexpr wchar_t kFolderBackgroundUniversalToken[] = L"folder_background_universal";
constexpr wchar_t kFolderBackgroundEntryToken[] = L"folder_background_entry";
constexpr wchar_t kTabDockingToken[] = L"tab_docking";
constexpr wchar_t kCommentChar = L'#';

std::wstring FormatLoadError(std::wstring message, DWORD error) {
    if (error != ERROR_SUCCESS) {
        message += L" (error=";
        message += std::to_wstring(static_cast<unsigned long long>(error));
        message += L")";
    }
    return message;
}

int ParseIntInRange(std::wstring_view token, int minimum, int maximum, int fallback) {
    if (token.empty()) {
        return fallback;
    }
    int value = ParseInt(token);
    if (value < minimum) {
        return minimum;
    }
    if (value > maximum) {
        return maximum;
    }
    return value;
}

COLORREF ParseColorValue(std::wstring_view token, COLORREF fallback) {
    if (token.empty()) {
        return fallback;
    }

    unsigned int parsed = 0;
    const std::wstring buffer(token);
    if (swscanf_s(buffer.c_str(), L"%x", &parsed) != 1) {
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

constexpr COLORREF kDefaultGlowPrimaryColor = RGB(0, 120, 215);
constexpr COLORREF kDefaultGlowSecondaryColor = RGB(0, 153, 255);

struct GlowSurfaceMapping {
    const wchar_t* token;
    GlowSurfaceOptions GlowSurfacePalette::*member;
    bool supportsExplorerAccent;
};

constexpr std::array<GlowSurfaceMapping, 6> kGlowSurfaceMappings = {{{L"header", &GlowSurfacePalette::header, false},
                                                                     {L"list_view", &GlowSurfacePalette::listView, true},
                                                                     {L"direct_ui", &GlowSurfacePalette::directUi, true},
                                                                     {L"toolbar", &GlowSurfacePalette::toolbar, false},
                                                                     {L"rebar", &GlowSurfacePalette::rebar, false},
                                                                     {L"edits", &GlowSurfacePalette::edits, false}}};

const GlowSurfaceMapping* FindGlowSurfaceMapping(std::wstring_view token, size_t* index) {
    if (token.empty()) {
        return nullptr;
    }
    for (size_t i = 0; i < kGlowSurfaceMappings.size(); ++i) {
        if (EqualsIgnoreCase(token, kGlowSurfaceMappings[i].token)) {
            if (index) {
                *index = i;
            }
            return &kGlowSurfaceMappings[i];
        }
    }
    return nullptr;
}

GlowSurfaceOptions* GetGlowSurfaceOptions(ShellTabsOptions* options, const GlowSurfaceMapping& mapping) {
    if (!options) {
        return nullptr;
    }
    return &(options->glowPalette.*(mapping.member));
}

const GlowSurfaceOptions* GetGlowSurfaceOptions(const ShellTabsOptions* options,
                                                const GlowSurfaceMapping& mapping) {
    if (!options) {
        return nullptr;
    }
    return &(options->glowPalette.*(mapping.member));
}

GlowSurfaceMode ParseGlowMode(std::wstring_view token, GlowSurfaceMode fallback) {
    if (token.empty()) {
        return fallback;
    }
    if (EqualsIgnoreCase(token, L"accent")) {
        return GlowSurfaceMode::kExplorerAccent;
    }
    if (EqualsIgnoreCase(token, L"solid")) {
        return GlowSurfaceMode::kSolid;
    }
    if (EqualsIgnoreCase(token, L"gradient")) {
        return GlowSurfaceMode::kGradient;
    }
    return fallback;
}

const wchar_t* GlowModeToString(GlowSurfaceMode mode) {
    switch (mode) {
        case GlowSurfaceMode::kExplorerAccent:
            return L"accent";
        case GlowSurfaceMode::kSolid:
            return L"solid";
        case GlowSurfaceMode::kGradient:
            return L"gradient";
        default:
            return L"gradient";
    }
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

void UpdateGlowPaletteFromLegacySettings(ShellTabsOptions& options) {
    const bool gradient = options.useNeonGlowGradient;
    const COLORREF primary = options.neonGlowPrimaryColor;
    const COLORREF secondary = options.useNeonGlowGradient ? options.neonGlowSecondaryColor : primary;

    const auto applyLegacy = [&](GlowSurfaceOptions& surface, bool allowAccent) {
        if (allowAccent && options.useExplorerAccentColors) {
            surface.mode = GlowSurfaceMode::kExplorerAccent;
        } else if (gradient) {
            surface.mode = GlowSurfaceMode::kGradient;
        } else {
            surface.mode = GlowSurfaceMode::kSolid;
        }
        surface.solidColor = primary;
        surface.gradientStartColor = primary;
        surface.gradientEndColor = gradient ? secondary : primary;
    };

    applyLegacy(options.glowPalette.header, false);
    applyLegacy(options.glowPalette.toolbar, false);
    applyLegacy(options.glowPalette.rebar, false);
    applyLegacy(options.glowPalette.edits, false);
    applyLegacy(options.glowPalette.listView, true);
    applyLegacy(options.glowPalette.directUi, true);
}

void UpdateLegacyGlowSettingsFromPalette(ShellTabsOptions& options) {
    const GlowSurfaceOptions& header = options.glowPalette.header;

    options.useExplorerAccentColors =
        options.glowPalette.listView.mode == GlowSurfaceMode::kExplorerAccent;

    options.useNeonGlowGradient = (header.mode == GlowSurfaceMode::kGradient);
    options.neonGlowPrimaryColor = header.solidColor;
    options.neonGlowSecondaryColor = options.useNeonGlowGradient ? header.gradientEndColor : header.solidColor;

    options.useCustomNeonGlowColors =
        header.solidColor != kDefaultGlowPrimaryColor ||
        header.gradientStartColor != kDefaultGlowPrimaryColor ||
        header.gradientEndColor != kDefaultGlowSecondaryColor;
}

OptionsStore& OptionsStore::Instance() {
    static OptionsStore store;
    return store;
}

bool OptionsStore::EnsureLoaded(std::wstring* errorContext) const {
    if (m_loaded) {
        if (errorContext) {
            errorContext->clear();
        }
        return true;
    }
    return const_cast<OptionsStore*>(this)->Load(errorContext);
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

bool OptionsStore::Load(std::wstring* errorContext) {
    m_options = {};

    m_storagePath = ResolveStoragePath();
    if (m_storagePath.empty()) {
        m_loaded = false;
        if (errorContext) {
            *errorContext = L"Options store path unavailable";
        }
        return false;
    }

    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(m_storagePath, &content, &fileExists)) {
        m_loaded = false;
        const DWORD readError = GetLastError();
        if (errorContext) {
            std::wstring message = L"Failed to read ";
            message += m_storagePath;
            *errorContext = FormatLoadError(std::move(message), readError);
        }
        return false;
    }

    const std::wstring storageDirectory = GetShellTabsDataDirectory();

    if (!fileExists || content.empty()) {
        m_loaded = true;
        if (errorContext) {
            errorContext->clear();
        }
        return true;
    }

    int version = 1;
    std::array<bool, kGlowSurfaceMappings.size()> glowSurfaceSpecified{};
    glowSurfaceSpecified.fill(false);
    bool anyGlowSurfaceToken = false;
    ParseConfigLines(content, kCommentChar, L'|', [&](const std::vector<std::wstring_view>& tokens) {
        if (tokens.empty()) {
            return true;
        }

        const std::wstring_view header = tokens[0];
        if (header == kVersionToken) {
            if (tokens.size() >= 2) {
                version = ParseInt(tokens[1]);
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

        if (header == kGlowEnabledToken) {
            if (tokens.size() >= 2) {
                m_options.enableNeonGlow = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kGlowGradientToken) {
            if (tokens.size() >= 2) {
                m_options.useNeonGlowGradient = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kGlowColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomNeonGlowColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.neonGlowPrimaryColor =
                    ParseColorValue(tokens[2], m_options.neonGlowPrimaryColor);
            }
            if (tokens.size() >= 4) {
                m_options.neonGlowSecondaryColor =
                    ParseColorValue(tokens[3], m_options.neonGlowSecondaryColor);
            }
            return true;
        }

        if (header == kGlowSurfaceToken) {
            if (tokens.size() >= 2) {
                size_t surfaceIndex = 0;
                const GlowSurfaceMapping* mapping = FindGlowSurfaceMapping(tokens[1], &surfaceIndex);
                if (mapping) {
                    if (GlowSurfaceOptions* surface = GetGlowSurfaceOptions(&m_options, *mapping)) {
                        anyGlowSurfaceToken = true;
                        glowSurfaceSpecified[surfaceIndex] = true;
                        if (tokens.size() >= 3) {
                            surface->mode = ParseGlowMode(tokens[2], surface->mode);
                        }
                        if (tokens.size() >= 4) {
                            surface->solidColor = ParseColorValue(tokens[3], surface->solidColor);
                        }
                        if (tokens.size() >= 5) {
                            surface->gradientStartColor =
                                ParseColorValue(tokens[4], surface->gradientStartColor);
                        }
                        if (tokens.size() >= 6) {
                            surface->gradientEndColor =
                                ParseColorValue(tokens[5], surface->gradientEndColor);
                        }
                        if (tokens.size() >= 7) {
                            surface->enabled = ParseBool(tokens[6]);
                        }
                    }
                }
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

        if (header == kExplorerAccentColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useExplorerAccentColors = ParseBool(tokens[1]);
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

    if (anyGlowSurfaceToken) {
        ShellTabsOptions fallbackOptions = m_options;
        UpdateLegacyGlowSettingsFromPalette(fallbackOptions);
        UpdateGlowPaletteFromLegacySettings(fallbackOptions);

        for (size_t i = 0; i < glowSurfaceSpecified.size(); ++i) {
            if (glowSurfaceSpecified[i]) {
                continue;
            }
            const GlowSurfaceMapping& mapping = kGlowSurfaceMappings[i];
            GlowSurfaceOptions* target = GetGlowSurfaceOptions(&m_options, mapping);
            const GlowSurfaceOptions* fallback = GetGlowSurfaceOptions(&fallbackOptions, mapping);
            if (target && fallback) {
                *target = *fallback;
            }
        }

        UpdateLegacyGlowSettingsFromPalette(m_options);
    } else {
        UpdateGlowPaletteFromLegacySettings(m_options);
    }

    m_loaded = true;
    if (errorContext) {
        errorContext->clear();
    }
    return true;
}

bool OptionsStore::Save() const {
    if (!EnsureLoaded()) {
        return false;
    }

    ShellTabsOptions persistedOptions = m_options;
    UpdateLegacyGlowSettingsFromPalette(persistedOptions);
    const_cast<OptionsStore*>(this)->m_options = persistedOptions;
    const ShellTabsOptions& options = const_cast<OptionsStore*>(this)->m_options;

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
    content += options.reopenOnCrash ? L"1" : L"0";
    content += L"\n";
    content += kPersistToken;
    content += L"|";
    content += options.persistGroupPaths ? L"1" : L"0";
    content += L"\n";
    content += kNewTabTemplateToken;
    content += L"|";
    content += NewTabTemplateToString(options.newTabTemplate);
    content += L"\n";
    content += kNewTabCustomPathToken;
    content += L"|";
    content += Trim(options.newTabCustomPath);
    content += L"\n";
    content += kNewTabSavedGroupToken;
    content += L"|";
    content += Trim(options.newTabSavedGroup);
    content += L"\n";
    content += kBreadcrumbGradientToken;
    content += L"|";
    content += options.enableBreadcrumbGradient ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbFontGradientToken;
    content += L"|";
    content += options.enableBreadcrumbFontGradient ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbGradientTransparencyToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbGradientTransparency, 0, 100));
    content += L"\n";
    content += kBreadcrumbFontBrightnessToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbFontBrightness, 0, 100));
    content += L"\n";
    content += kBreadcrumbHighlightAlphaMultiplierToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbHighlightAlphaMultiplier, 0, 200));
    content += L"\n";
    content += kBreadcrumbDropdownAlphaMultiplierToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbDropdownAlphaMultiplier, 0, 200));
    content += L"\n";
    content += kBreadcrumbGradientColorsToken;
    content += L"|";
    content += options.useCustomBreadcrumbGradientColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.breadcrumbGradientStartColor);
    content += L"|";
    content += ColorToHexString(options.breadcrumbGradientEndColor);
    content += L"\n";
    content += kBreadcrumbFontGradientColorsToken;
    content += L"|";
    content += options.useCustomBreadcrumbFontColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.breadcrumbFontGradientStartColor);
    content += L"|";
    content += ColorToHexString(options.breadcrumbFontGradientEndColor);
    content += L"\n";
    content += kProgressGradientColorsToken;
    content += L"|";
    content += options.useCustomProgressBarGradientColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.progressBarGradientStartColor);
    content += L"|";
    content += ColorToHexString(options.progressBarGradientEndColor);
    content += L"\n";
    content += kGlowEnabledToken;
    content += L"|";
    content += options.enableNeonGlow ? L"1" : L"0";
    content += L"\n";
    content += kGlowGradientToken;
    content += L"|";
    content += options.useNeonGlowGradient ? L"1" : L"0";
    content += L"\n";
    content += kGlowColorsToken;
    content += L"|";
    content += options.useCustomNeonGlowColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.neonGlowPrimaryColor);
    content += L"|";
    content += ColorToHexString(options.neonGlowSecondaryColor);
    content += L"\n";
    for (const auto& mapping : kGlowSurfaceMappings) {
        const GlowSurfaceOptions* surface = GetGlowSurfaceOptions(&options, mapping);
        if (!surface) {
            continue;
        }
        content += kGlowSurfaceToken;
        content += L"|";
        content += mapping.token;
        content += L"|";
        content += GlowModeToString(surface->mode);
        content += L"|";
        content += ColorToHexString(surface->solidColor);
        content += L"|";
        content += ColorToHexString(surface->gradientStartColor);
        content += L"|";
        content += ColorToHexString(surface->gradientEndColor);
        content += L"|";
        content += surface->enabled ? L"1" : L"0";
        content += L"\n";
    }
    content += kTabSelectedColorToken;
    content += L"|";
    content += options.useCustomTabSelectedColor ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.customTabSelectedColor);
    content += L"\n";
    content += kTabUnselectedColorToken;
    content += L"|";
    content += options.useCustomTabUnselectedColor ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.customTabUnselectedColor);
    content += L"\n";

    content += kExplorerAccentColorsToken;
    content += L"|";
    content += options.useExplorerAccentColors ? L"1" : L"0";
    content += L"\n";

    content += kFolderBackgroundsEnabledToken;
    content += L"|";
    content += options.enableFolderBackgrounds ? L"1" : L"0";
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

    appendCachedImageLine(kFolderBackgroundUniversalToken, options.universalFolderBackgroundImage.cachedImagePath,
                          options.universalFolderBackgroundImage.displayName);

    for (const auto& entry : options.folderBackgroundEntries) {
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
    content += DockModeToString(options.tabDockMode);
    content += L"\n";

    if (!WriteUtf8File(m_storagePath, content)) {
        return false;
    }

    UpdateCachedImageUsage(options);
    return true;
}

bool operator==(const CachedImageMetadata& left, const CachedImageMetadata& right) noexcept {
    return left.cachedImagePath == right.cachedImagePath && left.displayName == right.displayName;
}

bool operator==(const FolderBackgroundEntry& left, const FolderBackgroundEntry& right) noexcept {
    return left.folderPath == right.folderPath && left.image == right.image;
}

bool operator==(const GlowSurfaceOptions& left, const GlowSurfaceOptions& right) noexcept {
    return left.enabled == right.enabled && left.mode == right.mode && left.solidColor == right.solidColor &&
           left.gradientStartColor == right.gradientStartColor &&
           left.gradientEndColor == right.gradientEndColor;
}

bool operator==(const GlowSurfacePalette& left, const GlowSurfacePalette& right) noexcept {
    return left.header == right.header && left.listView == right.listView &&
           left.directUi == right.directUi && left.toolbar == right.toolbar && left.rebar == right.rebar &&
           left.edits == right.edits;
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
           left.enableNeonGlow == right.enableNeonGlow &&
           left.useNeonGlowGradient == right.useNeonGlowGradient &&
           left.useCustomNeonGlowColors == right.useCustomNeonGlowColors &&
           left.neonGlowPrimaryColor == right.neonGlowPrimaryColor &&
           left.neonGlowSecondaryColor == right.neonGlowSecondaryColor &&
           left.useCustomTabSelectedColor == right.useCustomTabSelectedColor &&
           left.customTabSelectedColor == right.customTabSelectedColor &&
           left.useCustomTabUnselectedColor == right.useCustomTabUnselectedColor &&
           left.customTabUnselectedColor == right.customTabUnselectedColor &&
           left.useExplorerAccentColors == right.useExplorerAccentColors &&
           left.glowPalette == right.glowPalette &&
           left.enableFolderBackgrounds == right.enableFolderBackgrounds &&
           left.universalFolderBackgroundImage == right.universalFolderBackgroundImage &&
           left.folderBackgroundEntries == right.folderBackgroundEntries &&
           left.tabDockMode == right.tabDockMode &&
           left.newTabTemplate == right.newTabTemplate &&
           left.newTabCustomPath == right.newTabCustomPath &&
           left.newTabSavedGroup == right.newTabSavedGroup;
}

}  // namespace shelltabs

