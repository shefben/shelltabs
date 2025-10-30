#include "TaskbarProxyWindow.h"

#include <algorithm>
#include <cwchar>
#include <mutex>

#include "Logging.h"
#include "Module.h"

namespace shelltabs {

namespace {

constexpr wchar_t kProxyWindowClassName[] = L"ShellTabsTaskbarProxyWindow";

}  // namespace

TaskbarProxyWindow::TaskbarProxyWindow(TabLocation location, ActivationCallback callback, void* context) noexcept
    : m_callback(callback), m_callbackContext(context), m_location(location) {}

TaskbarProxyWindow::~TaskbarProxyWindow() { Destroy(); }

ATOM TaskbarProxyWindow::RegisterClass() {
    static ATOM atom = 0;
    static std::once_flag once;
    std::call_once(once, []() {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = &TaskbarProxyWindow::WindowProc;
        wc.cbClsExtra = 0;
        wc.cbWndExtra = 0;
        wc.hInstance = GetModuleHandleInstance();
        wc.hIcon = nullptr;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.hbrBackground = nullptr;
        wc.lpszMenuName = nullptr;
        wc.lpszClassName = kProxyWindowClassName;
        wc.hIconSm = nullptr;
        atom = RegisterClassExW(&wc);
        if (!atom) {
            LogMessage(LogLevel::Error, L"TaskbarProxyWindow RegisterClassExW failed (err=%lu)", GetLastError());
        }
    });
    return atom;
}

TaskbarProxyWindow* TaskbarProxyWindow::FromHwnd(HWND hwnd) noexcept {
    return reinterpret_cast<TaskbarProxyWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

bool TaskbarProxyWindow::EnsureCreated(HWND frame, const FrameTabEntry& entry) {
    if (!frame || !IsWindow(frame)) {
        return false;
    }

    if (m_hwnd && m_frame != frame) {
        Destroy();
    }

    if (!m_hwnd) {
        if (!RegisterClass()) {
            return false;
        }

        m_frame = frame;
        m_hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, kProxyWindowClassName, L"",
                                 WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, frame, nullptr,
                                 GetModuleHandleInstance(), this);
        if (!m_hwnd) {
            LogMessage(LogLevel::Error, L"TaskbarProxyWindow CreateWindowExW failed (err=%lu)", GetLastError());
            m_frame = nullptr;
            return false;
        }
        ShowWindow(m_hwnd, SW_HIDE);
        m_registered = false;
    }

    UpdateEntry(entry);
    return true;
}

void TaskbarProxyWindow::UpdateEntry(const FrameTabEntry& entry) {
    m_location = entry.location;
    m_name = entry.name;
    m_tooltip = entry.tooltip;
    if (m_hwnd) {
        SetWindowTextW(m_hwnd, m_name.c_str());
    }
}

void TaskbarProxyWindow::Destroy() {
    if (!m_hwnd) {
        m_registered = false;
        m_frame = nullptr;
        return;
    }
    const HWND window = m_hwnd;
    m_hwnd = nullptr;
    m_registered = false;
    m_frame = nullptr;
    DestroyWindow(window);
}

void TaskbarProxyWindow::OnActivate() {
    if (!m_callback || !m_location.IsValid()) {
        return;
    }
    m_callback(m_callbackContext, m_location);
}

void TaskbarProxyWindow::OnCommand(WPARAM, LPARAM) {
    OnActivate();
}

LRESULT CALLBACK TaskbarProxyWindow::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_NCCREATE: {
            auto* create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
            auto* instance = static_cast<TaskbarProxyWindow*>(create->lpCreateParams);
            if (!instance) {
                return FALSE;
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(instance));
            return TRUE;
        }
        case WM_NCDESTROY: {
            if (auto* instance = FromHwnd(hwnd)) {
                instance->m_hwnd = nullptr;
                instance->m_frame = nullptr;
                instance->m_registered = false;
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            break;
        }
        case WM_ACTIVATE: {
            if (auto* instance = FromHwnd(hwnd)) {
                if (LOWORD(wParam) != WA_INACTIVE) {
                    instance->OnActivate();
                }
                return 0;
            }
            break;
        }
        case WM_COMMAND: {
            if (auto* instance = FromHwnd(hwnd)) {
                instance->OnCommand(wParam, lParam);
                return 0;
            }
            break;
        }
        case WM_GETTEXT: {
            if (auto* instance = FromHwnd(hwnd)) {
                wchar_t* buffer = reinterpret_cast<wchar_t*>(lParam);
                if (!buffer || wParam == 0) {
                    return 0;
                }
                wcsncpy_s(buffer, static_cast<size_t>(wParam), instance->m_name.c_str(), _TRUNCATE);
                return static_cast<LRESULT>(wcslen(buffer));
            }
            break;
        }
        case WM_GETTEXTLENGTH: {
            if (auto* instance = FromHwnd(hwnd)) {
                return static_cast<LRESULT>(instance->m_name.size());
            }
            break;
        }
        default:
            break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

std::wstring BuildFrameTooltip(const std::vector<FrameTabEntry>& entries) {
    if (entries.empty()) {
        return std::wstring();
    }

    auto selected = std::find_if(entries.begin(), entries.end(), [](const FrameTabEntry& entry) {
        return entry.selected;
    });
    if (selected != entries.end()) {
        if (!selected->tooltip.empty()) {
            return selected->tooltip;
        }
        if (!selected->name.empty()) {
            return selected->name;
        }
    }

    for (const auto& entry : entries) {
        if (!entry.tooltip.empty()) {
            return entry.tooltip;
        }
    }

    return entries.front().name;
}

}  // namespace shelltabs

