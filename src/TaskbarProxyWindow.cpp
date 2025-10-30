#include "TaskbarProxyWindow.h"

#include <windowsx.h>

#include <memory>

#include "Module.h"
#include "TaskbarTabController.h"

namespace shelltabs {

namespace {
constexpr wchar_t kProxyWindowClassName[] = L"ShellTabs.TaskbarProxy";

struct ProxyWindowState {
    TaskbarTabController* controller = nullptr;
    TabLocation location;
};

ATOM EnsureProxyWindowClass() {
    static ATOM atom = 0;
    if (atom) {
        return atom;
    }

    WNDCLASSEXW cls{};
    cls.cbSize = sizeof(cls);
    cls.style = CS_DBLCLKS;
    cls.lpfnWndProc = [](HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) -> LRESULT {
        auto* state = reinterpret_cast<ProxyWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        switch (message) {
            case WM_NCCREATE: {
                auto* create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
                if (!create) {
                    return FALSE;
                }
                auto* config = static_cast<const TaskbarProxyConfig*>(create->lpCreateParams);
                if (!config || !config->controller) {
                    return FALSE;
                }
                auto newState = std::make_unique<ProxyWindowState>();
                newState->controller = config->controller;
                newState->location = config->location;
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(newState.release()));
                return TRUE;
            }
            case WM_ACTIVATE: {
                if (LOWORD(wParam) != WA_INACTIVE && state && state->controller) {
                    state->controller->OnProxyActivated(state->location);
                }
                break;
            }
            case WM_COMMAND: {
                if (state && state->controller) {
                    state->controller->OnProxyCommand(state->location);
                }
                break;
            }
            case WM_NCDESTROY: {
                if (state) {
                    std::unique_ptr<ProxyWindowState> cleanup(state);
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
                break;
            }
        }
        return DefWindowProcW(hwnd, message, wParam, lParam);
    };
    cls.cbWndExtra = 0;
    cls.cbClsExtra = 0;
    cls.hInstance = GetModuleHandleInstance();
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    cls.hbrBackground = nullptr;
    cls.lpszClassName = kProxyWindowClassName;

    atom = RegisterClassExW(&cls);
    return atom;
}

ProxyWindowState* GetProxyWindowState(HWND hwnd) {
    return reinterpret_cast<ProxyWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

}  // namespace

HWND CreateTaskbarProxyWindow(const TaskbarProxyConfig& config) {
    if (!config.controller) {
        return nullptr;
    }
    if (!EnsureProxyWindowClass()) {
        return nullptr;
    }

    return CreateWindowExW(WS_EX_TOOLWINDOW, kProxyWindowClassName, L"", WS_POPUP,
                           CW_USEDEFAULT, CW_USEDEFAULT, 1, 1, config.owner, nullptr,
                           GetModuleHandleInstance(), const_cast<TaskbarProxyConfig*>(&config));
}

void DestroyTaskbarProxyWindow(HWND hwnd) {
    if (!hwnd) {
        return;
    }
    DestroyWindow(hwnd);
}

void UpdateTaskbarProxyLocation(HWND hwnd, const TabLocation& location) {
    if (!hwnd) {
        return;
    }
    if (auto* state = GetProxyWindowState(hwnd)) {
        state->location = location;
    }
}

}  // namespace shelltabs

