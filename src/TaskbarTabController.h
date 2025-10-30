#pragma once

#include <windows.h>

#include <memory>
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

    void SyncTabs(const std::vector<TabViewItem>& items, TabLocation active, HWND frame);
    void Reset();

    void OnProxyActivated(const TabLocation& location);
    void OnProxyCommand(const TabLocation& location);

private:
    struct Proxy;

    TabBand* m_owner = nullptr;
    Microsoft::WRL::ComPtr<ITaskbarList3> m_taskbar;
    HWND m_frame = nullptr;
    std::vector<std::unique_ptr<Proxy>> m_proxies;

    Proxy* FindProxy(TabLocation location) const;
    bool EnsureTaskbar();
    void DestroyAllProxies();
    void DestroyProxy(Proxy* proxy);
    Proxy* CreateProxy(TabLocation location, HWND frame);
    void UpdateProxy(Proxy* proxy, const TabViewItem& item);
};

}  // namespace shelltabs

