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

    static bool IsSupported() noexcept;

    void SyncFrameSummary(const std::vector<TabViewItem>& items, TabLocation active, HWND frame);
    void Reset();

private:
    struct CachedTab;

    TabBand* m_owner = nullptr;
    Microsoft::WRL::ComPtr<ITaskbarList3> m_taskbar;
    HWND m_frame = nullptr;
    std::vector<CachedTab> m_cachedTabs;
    TabLocation m_activeLocation;
    std::wstring m_frameTooltip;

    bool EnsureTaskbar();
    void RefreshFrameTooltip(HWND frame, const std::wstring& tooltip);
}; 

}  // namespace shelltabs

