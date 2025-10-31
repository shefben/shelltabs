#pragma once

#include <windows.h>

#include <string>
#include <vector>

namespace shelltabs {

struct CachedImageMetadata {
    std::wstring cachedImagePath;
    std::wstring displayName;
};

struct FolderBackgroundEntry {
    std::wstring folderPath;
    CachedImageMetadata image;
};

enum class TabBandDockMode {
    kAutomatic = 0,
    kTop,
    kBottom,
    kLeft,
    kRight,
};

enum class NewTabTemplate {
    kDuplicateCurrent = 0,
    kThisPc,
    kCustomPath,
    kSavedGroup,
};

struct ShellTabsOptions {
    bool reopenOnCrash = false;
    bool persistGroupPaths = false;
    bool enableBreadcrumbGradient = false;
    bool enableBreadcrumbFontGradient = false;
    int breadcrumbGradientTransparency = 45;  // percentage [0, 100]
    int breadcrumbFontBrightness = 85;        // percentage [0, 100]
    int breadcrumbHighlightAlphaMultiplier = 100;   // percentage [0, 200]
    int breadcrumbDropdownAlphaMultiplier = 100;    // percentage [0, 200]
    bool useCustomBreadcrumbGradientColors = false;
    COLORREF breadcrumbGradientStartColor = RGB(255, 59, 48);
    COLORREF breadcrumbGradientEndColor = RGB(175, 82, 222);
    bool useCustomBreadcrumbFontColors = false;
    COLORREF breadcrumbFontGradientStartColor = RGB(255, 255, 255);
    COLORREF breadcrumbFontGradientEndColor = RGB(255, 255, 255);
    bool useCustomProgressBarGradientColors = false;
    COLORREF progressBarGradientStartColor = RGB(0, 120, 215);
    COLORREF progressBarGradientEndColor = RGB(0, 153, 255);
    bool enableNeonGlow = false;
    bool useNeonGlowGradient = false;
    bool useCustomNeonGlowColors = false;
    COLORREF neonGlowPrimaryColor = RGB(0, 120, 215);
    COLORREF neonGlowSecondaryColor = RGB(0, 153, 255);
    bool useCustomTabSelectedColor = false;
    COLORREF customTabSelectedColor = RGB(0, 120, 215);
    bool useCustomTabUnselectedColor = false;
    COLORREF customTabUnselectedColor = RGB(200, 200, 200);
    bool useExplorerAccentColors = true;
    bool enableFolderBackgrounds = false;
    CachedImageMetadata universalFolderBackgroundImage;
    std::vector<FolderBackgroundEntry> folderBackgroundEntries;
    TabBandDockMode tabDockMode = TabBandDockMode::kAutomatic;
    NewTabTemplate newTabTemplate = NewTabTemplate::kDuplicateCurrent;
    std::wstring newTabCustomPath;
    std::wstring newTabSavedGroup;
};

class OptionsStore {
public:
    static OptionsStore& Instance();

    bool Load();
    bool Save() const;

    const ShellTabsOptions& Get() const noexcept { return m_options; }
    void Set(const ShellTabsOptions& options) { m_options = options; m_loaded = true; }

private:
    OptionsStore() = default;

    bool EnsureLoaded() const;
    std::wstring ResolveStoragePath() const;

    mutable bool m_loaded = false;
    mutable std::wstring m_storagePath;
    mutable ShellTabsOptions m_options;
};

bool operator==(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept;
inline bool operator!=(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept {
    return !(left == right);
}

bool operator==(const CachedImageMetadata& left, const CachedImageMetadata& right) noexcept;
inline bool operator!=(const CachedImageMetadata& left, const CachedImageMetadata& right) noexcept {
    return !(left == right);
}

bool operator==(const FolderBackgroundEntry& left, const FolderBackgroundEntry& right) noexcept;
inline bool operator!=(const FolderBackgroundEntry& left, const FolderBackgroundEntry& right) noexcept {
    return !(left == right);
}

}  // namespace shelltabs

