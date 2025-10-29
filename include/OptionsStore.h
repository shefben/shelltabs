#pragma once

#include <windows.h>

#include <string>

namespace shelltabs {

struct ShellTabsOptions {
    bool reopenOnCrash = false;
    bool persistGroupPaths = false;
    bool enableBreadcrumbGradient = false;
    bool enableBreadcrumbFontGradient = false;
    int breadcrumbGradientTransparency = 45;  // percentage [0, 100]
    int breadcrumbFontBrightness = 85;        // percentage [0, 100]
    bool useCustomBreadcrumbGradientColors = false;
    COLORREF breadcrumbGradientStartColor = RGB(255, 59, 48);
    COLORREF breadcrumbGradientEndColor = RGB(175, 82, 222);
    bool useCustomBreadcrumbFontColors = false;
    COLORREF breadcrumbFontGradientStartColor = RGB(255, 255, 255);
    COLORREF breadcrumbFontGradientEndColor = RGB(255, 255, 255);
    bool useCustomProgressBarGradientColors = false;
    COLORREF progressBarGradientStartColor = RGB(0, 120, 215);
    COLORREF progressBarGradientEndColor = RGB(0, 153, 255);
    bool useCustomTabSelectedColor = false;
    COLORREF customTabSelectedColor = RGB(0, 120, 215);
    bool useCustomTabUnselectedColor = false;
    COLORREF customTabUnselectedColor = RGB(200, 200, 200);
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

}  // namespace shelltabs

