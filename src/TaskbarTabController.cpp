#include "TaskbarTabController.h"

#include <VersionHelpers.h>

#include <algorithm>
#include <utility>

#include "Logging.h"
#include "TabBand.h"
#include "TabBandWindow.h"
#include "TaskbarProxyWindow.h"

namespace shelltabs {

namespace {

struct ProxyLocationEquals {
    bool operator()(const TabLocation& a, const TabLocation& b) const noexcept {
        return a.groupIndex == b.groupIndex && a.tabIndex == b.tabIndex;
    }
};

}  // namespace

struct TaskbarTabController::Proxy {
    HWND hwnd = nullptr;
    TabLocation location;
    std::wstring title;
    std::wstring tooltip;
    HICON smallIcon = nullptr;
    HICON largeIcon = nullptr;
    bool registered = false;
};

TaskbarTabController::TaskbarTabController(TabBand* owner) : m_owner(owner) {}

TaskbarTabController::~TaskbarTabController() { Reset(); }

bool TaskbarTabController::IsSupported() noexcept { return IsWindows7OrGreater() != FALSE; }

bool TaskbarTabController::EnsureTaskbar() {
    if (m_taskbar) {
        return true;
    }
    if (!IsSupported()) {
        return false;
    }

    Microsoft::WRL::ComPtr<ITaskbarList3> taskbar;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&taskbar));
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TaskbarTabController CoCreateInstance failed (hr=0x%08X)",
                   static_cast<unsigned int>(hr));
        return false;
    }
    hr = taskbar->HrInit();
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TaskbarTabController HrInit failed (hr=0x%08X)",
                   static_cast<unsigned int>(hr));
        return false;
    }
    m_taskbar = std::move(taskbar);
    return true;
}

TaskbarTabController::Proxy* TaskbarTabController::FindProxy(TabLocation location) const {
    auto it = std::find_if(m_proxies.begin(), m_proxies.end(), [&](const auto& candidate) {
        return ProxyLocationEquals{}(candidate->location, location);
    });
    return it != m_proxies.end() ? it->get() : nullptr;
}

void TaskbarTabController::DestroyProxy(Proxy* proxy) {
    if (!proxy) {
        return;
    }

    if (m_taskbar && proxy->registered && proxy->hwnd) {
        m_taskbar->UnregisterTab(proxy->hwnd);
        proxy->registered = false;
    }

    if (proxy->hwnd) {
        if (proxy->smallIcon) {
            HICON previous = reinterpret_cast<HICON>(SendMessageW(proxy->hwnd, WM_SETICON, ICON_SMALL, 0));
            if (previous) {
                DestroyIcon(previous);
            } else {
                DestroyIcon(proxy->smallIcon);
            }
            proxy->smallIcon = nullptr;
        }
        if (proxy->largeIcon) {
            HICON previous = reinterpret_cast<HICON>(SendMessageW(proxy->hwnd, WM_SETICON, ICON_BIG, 0));
            if (previous) {
                DestroyIcon(previous);
            } else {
                DestroyIcon(proxy->largeIcon);
            }
            proxy->largeIcon = nullptr;
        }
        DestroyTaskbarProxyWindow(proxy->hwnd);
        proxy->hwnd = nullptr;
    } else {
        if (proxy->smallIcon) {
            DestroyIcon(proxy->smallIcon);
            proxy->smallIcon = nullptr;
        }
        if (proxy->largeIcon) {
            DestroyIcon(proxy->largeIcon);
            proxy->largeIcon = nullptr;
        }
    }
}

void TaskbarTabController::DestroyAllProxies() {
    for (auto& proxy : m_proxies) {
        DestroyProxy(proxy.get());
    }
    m_proxies.clear();
}

TaskbarTabController::Proxy* TaskbarTabController::CreateProxy(TabLocation location, HWND frame) {
    TaskbarProxyConfig config;
    config.controller = this;
    config.location = location;
    config.owner = frame;

    HWND hwnd = CreateTaskbarProxyWindow(config);
    if (!hwnd) {
        return nullptr;
    }

    auto proxy = std::make_unique<Proxy>();
    proxy->hwnd = hwnd;
    proxy->location = location;
    proxy->registered = false;

    if (m_taskbar) {
        HRESULT hr = m_taskbar->RegisterTab(hwnd, frame);
        if (FAILED(hr)) {
            LogMessage(LogLevel::Warning, L"TaskbarTabController RegisterTab failed (hr=0x%08X)",
                       static_cast<unsigned int>(hr));
        } else {
            proxy->registered = true;
        }
    }

    m_proxies.push_back(std::move(proxy));
    return m_proxies.back().get();
}

void TaskbarTabController::UpdateProxy(Proxy* proxy, const TabViewItem& item) {
    if (!proxy || !proxy->hwnd) {
        return;
    }

    if (!ProxyLocationEquals{}(proxy->location, item.location)) {
        proxy->location = item.location;
        UpdateTaskbarProxyLocation(proxy->hwnd, proxy->location);
    }

    if (proxy->title != item.name) {
        SetWindowTextW(proxy->hwnd, item.name.c_str());
        proxy->title = item.name;
    }

    if (proxy->tooltip != item.tooltip) {
        if (m_taskbar) {
            m_taskbar->SetThumbnailTooltip(proxy->hwnd, item.tooltip.c_str());
        }
        proxy->tooltip = item.tooltip;
    }

    auto applyIcon = [&](UINT type, HICON newIcon, HICON& storage) {
        if (storage == newIcon) {
            return;
        }
        HICON previous = reinterpret_cast<HICON>(
            SendMessageW(proxy->hwnd, WM_SETICON, type, reinterpret_cast<LPARAM>(newIcon)));
        if (previous && previous != newIcon) {
            DestroyIcon(previous);
        }
        if (storage && storage != previous && storage != newIcon) {
            DestroyIcon(storage);
        }
        storage = newIcon;
    };

    HICON smallIcon = nullptr;
    HICON largeIcon = nullptr;
    if (m_owner && m_owner->m_window) {
        smallIcon = m_owner->m_window->GetTaskbarIcon(item, true);
        largeIcon = m_owner->m_window->GetTaskbarIcon(item, false);
    }

    applyIcon(ICON_SMALL, smallIcon, proxy->smallIcon);
    applyIcon(ICON_BIG, largeIcon, proxy->largeIcon);
}

void TaskbarTabController::SyncTabs(const std::vector<TabViewItem>& items, TabLocation active, HWND frame) {
    if (!IsSupported()) {
        Reset();
        return;
    }

    if (!frame) {
        DestroyAllProxies();
        m_frame = nullptr;
        return;
    }

    if (!EnsureTaskbar()) {
        DestroyAllProxies();
        return;
    }

    if (frame != m_frame) {
        DestroyAllProxies();
        m_frame = frame;
    }

    std::vector<Proxy*> ordered;
    ordered.reserve(items.size());

    HWND previousHwnd = nullptr;

    for (const auto& item : items) {
        if (item.type != TabViewItemType::kTab) {
            continue;
        }

        Proxy* proxy = FindProxy(item.location);
        if (!proxy) {
            proxy = CreateProxy(item.location, frame);
        }
        if (!proxy || !proxy->hwnd) {
            continue;
        }

        ordered.push_back(proxy);
        UpdateProxy(proxy, item);

        if (m_taskbar && !proxy->registered) {
            HRESULT registerHr = m_taskbar->RegisterTab(proxy->hwnd, frame);
            if (FAILED(registerHr)) {
                LogMessage(LogLevel::Warning,
                           L"TaskbarTabController RegisterTab retry failed (hr=0x%08X)",
                           static_cast<unsigned int>(registerHr));
            } else {
                proxy->registered = true;
            }
        }

        if (m_taskbar) {
            HRESULT hr = m_taskbar->SetTabOrder(proxy->hwnd, previousHwnd);
            if (FAILED(hr)) {
                LogMessage(LogLevel::Warning, L"TaskbarTabController SetTabOrder failed (hr=0x%08X)",
                           static_cast<unsigned int>(hr));
            }
        }

        previousHwnd = proxy->hwnd;
    }

    for (auto it = m_proxies.begin(); it != m_proxies.end();) {
        Proxy* proxy = it->get();
        const bool keep = std::find(ordered.begin(), ordered.end(), proxy) != ordered.end();
        if (!keep) {
            DestroyProxy(proxy);
            it = m_proxies.erase(it);
        } else {
            ++it;
        }
    }

    Proxy* activeProxy = nullptr;
    if (active.IsValid()) {
        activeProxy = FindProxy(active);
    }
    if (!activeProxy) {
        auto it = std::find_if(ordered.begin(), ordered.end(), [&](Proxy* proxy) {
            return proxy && ProxyLocationEquals{}(proxy->location, active);
        });
        if (it != ordered.end()) {
            activeProxy = *it;
        }
    }

    if (!activeProxy) {
        auto it = std::find_if(items.begin(), items.end(), [](const TabViewItem& item) {
            return item.type == TabViewItemType::kTab && item.selected;
        });
        if (it != items.end()) {
            activeProxy = FindProxy(it->location);
        }
    }

    if (activeProxy && m_taskbar) {
        HRESULT hr = m_taskbar->SetTabActive(activeProxy->hwnd, frame, 0);
        if (FAILED(hr)) {
            LogMessage(LogLevel::Warning, L"TaskbarTabController SetTabActive failed (hr=0x%08X)",
                       static_cast<unsigned int>(hr));
        }
    }
}

void TaskbarTabController::Reset() {
    DestroyAllProxies();
    m_taskbar.Reset();
    m_frame = nullptr;
}

void TaskbarTabController::OnProxyActivated(const TabLocation& location) {
    if (!m_owner) {
        return;
    }
    m_owner->OnTabSelected(location);

    if (Proxy* proxy = FindProxy(location); proxy && m_taskbar && m_frame) {
        m_taskbar->SetTabActive(proxy->hwnd, m_frame, 0);
    }
}

void TaskbarTabController::OnProxyCommand(const TabLocation& location) {
    OnProxyActivated(location);
}

}  // namespace shelltabs

