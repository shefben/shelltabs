#pragma once

#include <windows.h>

#include <string>
#include <vector>

#include <ShObjIdl_core.h>
#include <wrl/client.h>

#include "TabManager.h"

namespace shelltabs {

class TabBand;

class TaskbarTabController {
public:
    explicit TaskbarTabController(TabBand* owner);
    ~TaskbarTabController();

    TaskbarTabController(const TaskbarTabController&) = delete;
    TaskbarTabController& operator=(const TaskbarTabController&) = delete;

    static constexpr UINT kThumbnailToolbarCommandId = 0xE171;

    static bool IsSupported() noexcept;

    void SyncFrameSummary(const std::vector<TabViewItem>& items, TabLocation active, HWND frame);
    void Reset();
    void HandleThumbnailButton(const POINT& anchor);

private:
    struct CachedTab {
        TabLocation location;
        std::wstring name;
        std::wstring tooltip;
        bool selected = false;
    };

    static bool LocationsEqual(const TabLocation& a, const TabLocation& b) noexcept;
    static bool TabsEqual(const std::vector<CachedTab>& a, const std::vector<CachedTab>& b);

    TabBand* m_owner = nullptr;
    Microsoft::WRL::ComPtr<ITaskbarList3> m_taskbar;
    HWND m_frame = nullptr;
    std::vector<CachedTab> m_cachedTabs;
    TabLocation m_activeLocation;
    std::wstring m_frameTooltip;
    HWND m_thumbButtonFrame = nullptr;
    bool m_thumbButtonAdded = false;
    HICON m_thumbButtonIcon = nullptr;

    bool EnsureTaskbar();
    void RefreshFrameTooltip(HWND frame, const std::wstring& tooltip);
    void EnsureThumbnailButton(HWND frame);
};

}  // namespace shelltabs

