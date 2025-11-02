#pragma once

#include <windows.h>

namespace shelltabs {

using SetThreadDpiAwarenessContextFunction = DPI_AWARENESS_CONTEXT(WINAPI*)(DPI_AWARENESS_CONTEXT);

UINT GetWindowDpi(HWND hwnd);

class ScopedThreadDpiAwarenessContext {
public:
    ScopedThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT desiredContext, bool enabled);
    ScopedThreadDpiAwarenessContext(const ScopedThreadDpiAwarenessContext&) = delete;
    ScopedThreadDpiAwarenessContext& operator=(const ScopedThreadDpiAwarenessContext&) = delete;
    ~ScopedThreadDpiAwarenessContext();

private:
    DPI_AWARENESS_CONTEXT m_previousContext = nullptr;
    SetThreadDpiAwarenessContextFunction m_setThreadDpiAwarenessContext = nullptr;
};

}  // namespace shelltabs

