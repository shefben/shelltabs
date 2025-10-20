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
    int breadcrumbFontTransparency = 0;       // percentage [0, 100]
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

