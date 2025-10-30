#pragma once

#include <windows.h>

#include <string>
#include <vector>

#include "TabManager.h"

namespace shelltabs {

struct FrameTabEntry {
    TabLocation location;
    std::wstring name;
    std::wstring tooltip;
    bool selected = false;
};

class TaskbarProxyWindow {
public:
    using ActivationCallback = void (*)(void* context, TabLocation location);

    TaskbarProxyWindow(TabLocation location, ActivationCallback callback, void* context) noexcept;
    ~TaskbarProxyWindow();

    TaskbarProxyWindow(const TaskbarProxyWindow&) = delete;
    TaskbarProxyWindow& operator=(const TaskbarProxyWindow&) = delete;

    bool EnsureCreated(HWND frame, const FrameTabEntry& entry);
    void UpdateEntry(const FrameTabEntry& entry);
    void Destroy();

    HWND GetHwnd() const noexcept { return m_hwnd; }
    TabLocation GetLocation() const noexcept { return m_location; }

    bool IsRegistered() const noexcept { return m_registered; }
    void SetRegistered(bool registered) noexcept { m_registered = registered; }

private:
    static ATOM RegisterClass();
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    static TaskbarProxyWindow* FromHwnd(HWND hwnd) noexcept;

    void OnActivate();
    void OnCommand(WPARAM wParam, LPARAM lParam);

    ActivationCallback m_callback = nullptr;
    void* m_callbackContext = nullptr;
    TabLocation m_location{};
    HWND m_hwnd = nullptr;
    HWND m_frame = nullptr;
    bool m_registered = false;
    std::wstring m_name;
    std::wstring m_tooltip;
};

std::wstring BuildFrameTooltip(const std::vector<FrameTabEntry>& entries);

}  // namespace shelltabs

