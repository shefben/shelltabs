#include "DirectUiHooks.h"

#include <UIAutomation.h>
#include <UIAutomationCoreApi.h>

#include <algorithm>
#include <mutex>
#include <vector>

#include <wrl/client.h>

#include "Logging.h"

namespace shelltabs {
namespace {

Microsoft::WRL::ComPtr<IUIAutomation> GetAutomationInstance() {
    static std::once_flag automationInitFlag;
    static Microsoft::WRL::ComPtr<IUIAutomation> automation;
    static HRESULT automationInitResult = E_FAIL;

    std::call_once(automationInitFlag, []() {
        Microsoft::WRL::ComPtr<IUIAutomation> instance;
        const HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER,
                                            IID_PPV_ARGS(&instance));
        if (SUCCEEDED(hr)) {
            automation = std::move(instance);
            automationInitResult = S_OK;
        } else {
            automation.Reset();
            automationInitResult = hr;
        }
    });

    if (FAILED(automationInitResult)) {
        return nullptr;
    }

    return automation;
}

void AppendDirectUiRectangle(IUIAutomationElement* element, HWND host, const RECT& clientRect,
                             std::vector<RECT>& rectangles) {
    if (!element || !host || !IsWindow(host)) {
        return;
    }

    RECT bounding{};
    if (FAILED(element->get_CurrentBoundingRectangle(&bounding))) {
        return;
    }

    if (bounding.right <= bounding.left || bounding.bottom <= bounding.top) {
        return;
    }

    BOOL offscreen = FALSE;
    const HRESULT offscreenHr = element->get_CurrentIsOffscreen(&offscreen);
    if (FAILED(offscreenHr) || offscreen) {
        return;
    }

    POINT points[2] = {{bounding.left, bounding.top}, {bounding.right, bounding.bottom}};
    MapWindowPoints(nullptr, host, points, 2);

    RECT local = {points[0].x, points[0].y, points[1].x, points[1].y};
    RECT clipped{};
    if (!IntersectRect(&clipped, &local, &clientRect)) {
        return;
    }

    if (clipped.right <= clipped.left || clipped.bottom <= clipped.top) {
        return;
    }

    const auto duplicate = std::find_if(rectangles.begin(), rectangles.end(), [&](const RECT& existing) {
        return EqualRect(&existing, &clipped);
    });

    if (duplicate == rectangles.end()) {
        rectangles.push_back(clipped);
    }
}

void CollectDirectUiDescendants(IUIAutomationElement* element, IUIAutomationTreeWalker* walker, HWND host,
                                const RECT& clientRect, std::vector<RECT>& rectangles) {
    if (!element || !walker) {
        return;
    }

    Microsoft::WRL::ComPtr<IUIAutomationElement> child;
    HRESULT hr = walker->GetFirstChildElement(element, &child);
    if (FAILED(hr)) {
        return;
    }

    while (child) {
        AppendDirectUiRectangle(child.Get(), host, clientRect, rectangles);
        CollectDirectUiDescendants(child.Get(), walker, host, clientRect, rectangles);

        Microsoft::WRL::ComPtr<IUIAutomationElement> next;
        hr = walker->GetNextSiblingElement(child.Get(), &next);
        if (FAILED(hr)) {
            break;
        }
        child = std::move(next);
    }
}

bool EnumerateDirectUiRectangles(HWND host, const RECT& clientRect, std::vector<RECT>& rectangles) {
    rectangles.clear();

    auto automation = GetAutomationInstance();
    if (!automation) {
        return false;
    }

    Microsoft::WRL::ComPtr<IUIAutomationElement> root;
    if (FAILED(automation->ElementFromHandle(host, &root)) || !root) {
        return false;
    }

    Microsoft::WRL::ComPtr<IUIAutomationTreeWalker> walker;
    if (FAILED(automation->get_ControlViewWalker(&walker)) || !walker) {
        return false;
    }

    CollectDirectUiDescendants(root.Get(), walker.Get(), host, clientRect, rectangles);
    return !rectangles.empty();
}

}  // namespace

DirectUiHooks& DirectUiHooks::Instance() {
    static DirectUiHooks instance;
    return instance;
}

void DirectUiHooks::RegisterHost(HWND host) {
    if (!host) {
        return;
    }

    std::lock_guard<std::mutex> lock(m_lock);
    auto& entry = m_hosts[host];
    entry.hwnd = host;
    entry.windowProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(host, GWLP_WNDPROC));
    entry.classProc = reinterpret_cast<WNDPROC>(GetClassLongPtrW(host, GCLP_WNDPROC));
}

void DirectUiHooks::UnregisterHost(HWND host) {
    if (!host) {
        return;
    }

    std::lock_guard<std::mutex> lock(m_lock);
    m_hosts.erase(host);
}

bool DirectUiHooks::TryResolveHostLocked(HostEntry& entry) {
    entry.attempted = true;

    if (!entry.hwnd || !IsWindow(entry.hwnd)) {
        entry.resolved = false;
        return false;
    }

    IRawElementProviderSimple* provider = nullptr;
    const HRESULT hr = UiaHostProviderFromHwnd(entry.hwnd, &provider);
    if (SUCCEEDED(hr) && provider) {
        provider->Release();
        entry.resolved = true;
    } else {
        entry.resolved = false;
        LogMessage(LogLevel::Verbose,
                   L"DirectUiHooks failed to resolve host provider (hwnd=%p hr=0x%08X)", entry.hwnd, hr);
    }
    return entry.resolved;
}

bool DirectUiHooks::PaintHost(HWND host, const RECT& clientRect, const PaintCallback& callback) {
    if (!host || !IsWindow(host) || !callback) {
        return false;
    }

    bool resolved = false;
    {
        std::lock_guard<std::mutex> lock(m_lock);
        auto it = m_hosts.find(host);
        if (it == m_hosts.end()) {
            HostEntry entry{};
            entry.hwnd = host;
            entry.windowProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(host, GWLP_WNDPROC));
            entry.classProc = reinterpret_cast<WNDPROC>(GetClassLongPtrW(host, GCLP_WNDPROC));
            auto [insertedIt, _] = m_hosts.emplace(host, std::move(entry));
            it = insertedIt;
        }

        HostEntry& entry = it->second;
        if (!entry.attempted) {
            TryResolveHostLocked(entry);
        }
        resolved = entry.resolved;
    }

    if (!resolved) {
        return false;
    }

    std::vector<RECT> rectangles;
    if (!EnumerateDirectUiRectangles(host, clientRect, rectangles) || rectangles.empty()) {
        return false;
    }

    callback(rectangles);
    return true;
}

bool DirectUiHooks::EnumerateRectangles(HWND host, const RECT& clientRect, std::vector<RECT>& rectangles) const {
    if (!host || !IsWindow(host)) {
        rectangles.clear();
        return false;
    }

    return EnumerateDirectUiRectangles(host, clientRect, rectangles);
}

}  // namespace shelltabs

