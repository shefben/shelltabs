#include "ThemeNotifier.h"

#include <windows.h>
#include <wtsapi32.h>

#include <atomic>

namespace {

LRESULT CALLBACK TestWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_DESTROY) {
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

HWND CreateTestWindow() {
    WNDCLASSW wc{};
    wc.lpfnWndProc = TestWindowProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"ShellTabsThemeTestWindow";
    RegisterClassW(&wc);
    return CreateWindowExW(0, wc.lpszClassName, L"", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
                           CW_USEDEFAULT, CW_USEDEFAULT, nullptr, nullptr, wc.hInstance, nullptr);
}

void PumpMessagesOnce() {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

}  // namespace

int wmain() {
    HWND hwnd = CreateTestWindow();
    if (!hwnd) {
        return 1;
    }

    std::atomic<int> callbackCount{0};
    shelltabs::ThemeNotifier notifier;
    if (!notifier.Initialize(hwnd, [&]() {
            callbackCount.fetch_add(1, std::memory_order_relaxed);
        })) {
        DestroyWindow(hwnd);
        return 1;
    }

    notifier.SimulateColorChangeForTest();
    PumpMessagesOnce();

    notifier.SimulateSessionEventForTest(WTS_SESSION_UNLOCK);
    PumpMessagesOnce();

    bool success = callbackCount.load(std::memory_order_relaxed) >= 2;

    notifier.Shutdown();
    DestroyWindow(hwnd);
    return success ? 0 : 1;
}

