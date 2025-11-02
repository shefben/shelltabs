#include "DpiUtils.h"

#include <windows.h>

namespace {

using GetDpiForWindowFunction = UINT(WINAPI*)(HWND);
using SetThreadDpiAwarenessContextFunction = DPI_AWARENESS_CONTEXT(WINAPI*)(DPI_AWARENESS_CONTEXT);

GetDpiForWindowFunction ResolveGetDpiForWindow() {
    static const GetDpiForWindowFunction function =
        reinterpret_cast<GetDpiForWindowFunction>(GetProcAddress(GetModuleHandleW(L"user32"), "GetDpiForWindow"));
    return function;
}

SetThreadDpiAwarenessContextFunction ResolveSetThreadDpiAwarenessContext() {
    static const SetThreadDpiAwarenessContextFunction function =
        reinterpret_cast<SetThreadDpiAwarenessContextFunction>(
            GetProcAddress(GetModuleHandleW(L"user32"), "SetThreadDpiAwarenessContext"));
    return function;
}

}  // namespace

namespace shelltabs {

UINT GetWindowDpi(HWND hwnd) {
    if (auto* function = ResolveGetDpiForWindow()) {
        const UINT dpi = function(hwnd);
        if (dpi != 0) {
            return dpi;
        }
    }

    HDC localDc = hwnd ? GetDC(hwnd) : GetDC(nullptr);
    if (!localDc) {
        return 96u;
    }

    const int dpi = GetDeviceCaps(localDc, LOGPIXELSX);
    if (hwnd) {
        ReleaseDC(hwnd, localDc);
    } else {
        ReleaseDC(nullptr, localDc);
    }

    return dpi > 0 ? static_cast<UINT>(dpi) : 96u;
}

ScopedThreadDpiAwarenessContext::ScopedThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT desiredContext,
                                                                 bool enabled) {
    auto* function = ResolveSetThreadDpiAwarenessContext();
    if (!enabled || !function || !desiredContext) {
        return;
    }

    m_setThreadDpiAwarenessContext = function;
    m_previousContext = m_setThreadDpiAwarenessContext(desiredContext);
}

ScopedThreadDpiAwarenessContext::~ScopedThreadDpiAwarenessContext() {
    if (m_setThreadDpiAwarenessContext && m_previousContext) {
        m_setThreadDpiAwarenessContext(m_previousContext);
    }
}

}  // namespace shelltabs

