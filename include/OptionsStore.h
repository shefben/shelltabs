#pragma once

#include <windows.h>

#include "IconCache.h"

#include <string>
#include <vector>

namespace shelltabs {

enum class ContextMenuInsertionAnchor {
    kDefault = 0,
    kTop,
    kBottom,
    kBeforeShellItems,
    kAfterShellItems,
};

enum class ContextMenuItemType {
    kCommand = 0,
    kSubmenu,
    kSeparator,
};

struct ContextMenuSelectionRule {
    int minimumSelection = 0;
    int maximumSelection = 0;  // 0 == unlimited
};

struct ContextMenuItemScope {
    bool includeAllFiles = true;
    bool includeAllFolders = true;
    std::vector<std::wstring> extensions;
};

struct ContextMenuItem {
    ContextMenuItemType type = ContextMenuItemType::kCommand;
    std::wstring label;
    std::wstring iconSource;
    std::wstring commandTemplate;
    ContextMenuItemScope scope;
    ContextMenuSelectionRule selection;
    ContextMenuInsertionAnchor anchor = ContextMenuInsertionAnchor::kDefault;
    std::vector<ContextMenuItem> children;
};

IconCache::Reference ResolveContextMenuIcon(const std::wstring& iconSource, UINT iconFlags);
std::wstring NormalizeContextMenuIconSource(const std::wstring& iconSource);
std::vector<std::wstring> NormalizeContextMenuExtensions(const std::vector<std::wstring>& extensions);

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

enum class GlowSurfaceMode {
    kExplorerAccent = 0,
    kSolid,
    kGradient,
};

struct GlowSurfaceOptions {
    bool enabled = true;
    GlowSurfaceMode mode = GlowSurfaceMode::kGradient;
    COLORREF solidColor = RGB(0, 120, 215);
    COLORREF gradientStartColor = RGB(0, 120, 215);
    COLORREF gradientEndColor = RGB(0, 153, 255);

    GlowSurfaceOptions() = default;
    explicit GlowSurfaceOptions(GlowSurfaceMode initialMode) : mode(initialMode) {}
};

struct GlowSurfacePalette {
    GlowSurfaceOptions header{};
    GlowSurfaceOptions listView{GlowSurfaceMode::kExplorerAccent};
    GlowSurfaceOptions directUi{GlowSurfaceMode::kExplorerAccent};
    GlowSurfaceOptions toolbar{};
    GlowSurfaceOptions rebar{};
    GlowSurfaceOptions edits{};
    GlowSurfaceOptions scrollbars{GlowSurfaceMode::kExplorerAccent};
    GlowSurfaceOptions popupMenus{GlowSurfaceMode::kExplorerAccent};
    GlowSurfaceOptions tooltips{GlowSurfaceMode::kExplorerAccent};
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
    GlowSurfacePalette glowPalette{};
    bool enableFolderBackgrounds = false;
    CachedImageMetadata universalFolderBackgroundImage;
    std::vector<FolderBackgroundEntry> folderBackgroundEntries;
    std::vector<ContextMenuItem> contextMenuItems;
    TabBandDockMode tabDockMode = TabBandDockMode::kAutomatic;
    NewTabTemplate newTabTemplate = NewTabTemplate::kDuplicateCurrent;
    std::wstring newTabCustomPath;
    std::wstring newTabSavedGroup;
};

class OptionsStore {
public:
    static OptionsStore& Instance();

    bool Load(std::wstring* errorContext = nullptr);
    bool Save() const;

    const ShellTabsOptions& Get() const noexcept { return m_options; }
    void Set(const ShellTabsOptions& options) { m_options = options; m_loaded = true; }

private:
    OptionsStore() = default;

    bool EnsureLoaded(std::wstring* errorContext = nullptr) const;
    std::wstring ResolveStoragePath() const;

    mutable bool m_loaded = false;
    mutable std::wstring m_storagePath;
    mutable ShellTabsOptions m_options;
};

bool operator==(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept;
inline bool operator!=(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept {
    return !(left == right);
}

bool operator==(const GlowSurfaceOptions& left, const GlowSurfaceOptions& right) noexcept;
inline bool operator!=(const GlowSurfaceOptions& left, const GlowSurfaceOptions& right) noexcept {
    return !(left == right);
}

bool operator==(const GlowSurfacePalette& left, const GlowSurfacePalette& right) noexcept;
inline bool operator!=(const GlowSurfacePalette& left, const GlowSurfacePalette& right) noexcept {
    return !(left == right);
}

bool operator==(const ContextMenuSelectionRule& left, const ContextMenuSelectionRule& right) noexcept;
inline bool operator!=(const ContextMenuSelectionRule& left, const ContextMenuSelectionRule& right) noexcept {
    return !(left == right);
}

bool operator==(const ContextMenuItemScope& left, const ContextMenuItemScope& right) noexcept;
inline bool operator!=(const ContextMenuItemScope& left, const ContextMenuItemScope& right) noexcept {
    return !(left == right);
}

bool operator==(const ContextMenuItem& left, const ContextMenuItem& right) noexcept;
inline bool operator!=(const ContextMenuItem& left, const ContextMenuItem& right) noexcept {
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

void UpdateGlowPaletteFromLegacySettings(ShellTabsOptions& options);
void UpdateLegacyGlowSettingsFromPalette(ShellTabsOptions& options);

}  // namespace shelltabs

