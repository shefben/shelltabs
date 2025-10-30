#include "PreviewOverlay.h"

#include <algorithm>

#include <CommCtrl.h>

#include "Module.h"

namespace shelltabs {
namespace {

constexpr wchar_t kPreviewWindowClassName[] = L"ShellTabsPreviewWindow";

}  // namespace

PreviewOverlay::~PreviewOverlay() { Destroy(); }

ATOM PreviewOverlay::EnsureWindowClass() {
    static ATOM atom = 0;
    if (atom != 0) {
        return atom;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc = DefWindowProcW;
    wc.hInstance = GetModuleHandleInstance();
    wc.lpszClassName = kPreviewWindowClassName;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);

    atom = RegisterClassW(&wc);
    if (atom == 0 && GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
        atom = 1;  // sentinel for already registered
    }

    return atom;
}

HWND PreviewOverlay::EnsureWindow(HWND owner) {
    if (m_window && IsWindow(m_window)) {
        if (owner && GetParent(m_window) != owner) {
            SetWindowLongPtrW(m_window, GWLP_HWNDPARENT, reinterpret_cast<LONG_PTR>(owner));
        }
        m_owner = owner;
        return m_window;
    }

    if (!EnsureWindowClass()) {
        return nullptr;
    }

    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_NOACTIVATE, kPreviewWindowClassName,
                                L"", WS_POPUP, 0, 0, 0, 0, owner, nullptr, GetModuleHandleInstance(), nullptr);
    if (!hwnd) {
        return nullptr;
    }

    m_window = hwnd;
    m_owner = owner;
    return m_window;
}

bool PreviewOverlay::Show(HWND owner, HBITMAP bitmap, const SIZE& size, const POINT& screenPt) {
    if (!bitmap) {
        return false;
    }

    HWND window = EnsureWindow(owner);
    if (!window) {
        return false;
    }

    HDC screenDC = GetDC(nullptr);
    if (!screenDC) {
        return false;
    }

    HDC memDC = CreateCompatibleDC(screenDC);
    if (!memDC) {
        ReleaseDC(nullptr, screenDC);
        return false;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, bitmap);
    POINT dest = screenPt;
    dest.x += 16;
    dest.y += 16;
    POINT src{0, 0};
    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 255;
    blend.AlphaFormat = AC_SRC_ALPHA;

    UpdateLayeredWindow(window, screenDC, &dest, &size, memDC, &src, 0, &blend, ULW_ALPHA);

    SelectObject(memDC, oldBitmap);
    DeleteDC(memDC);
    ReleaseDC(nullptr, screenDC);

    m_bitmapSize = size;
    m_visible = true;
    ShowWindow(window, SW_SHOWNOACTIVATE);
    return true;
}

void PreviewOverlay::Hide(bool destroyWindow) {
    if (!m_window || !IsWindow(m_window)) {
        m_visible = false;
        return;
    }

    ShowWindow(m_window, SW_HIDE);
    if (destroyWindow) {
        DestroyWindow(m_window);
        m_window = nullptr;
    }
    m_visible = false;
}

void PreviewOverlay::Destroy() {
    Hide(true);
    m_owner = nullptr;
    m_bitmapSize = {};
}

void PreviewOverlay::PositionRelativeToRect(const RECT& anchorRectScreen, const POINT& cursorScreenPt) {
    if (!m_window || !IsWindow(m_window)) {
        return;
    }

    const int width = m_bitmapSize.cx;
    const int height = m_bitmapSize.cy;
    int x = anchorRectScreen.left + ((anchorRectScreen.right - anchorRectScreen.left) - width) / 2;
    int y = anchorRectScreen.bottom + 8;

    HMONITOR monitor = MonitorFromPoint(cursorScreenPt, MONITOR_DEFAULTTONEAREST);
    MONITORINFO monitorInfo{sizeof(monitorInfo)};
    if (GetMonitorInfoW(monitor, &monitorInfo)) {
        x = std::clamp<int>(x, monitorInfo.rcWork.left + 4, monitorInfo.rcWork.right - width - 4);
        if (y + height > monitorInfo.rcWork.bottom) {
            y = anchorRectScreen.top - height - 8;
        }
        if (y < monitorInfo.rcWork.top + 4) {
            y = monitorInfo.rcWork.top + 4;
        }
    }

    SetWindowPos(m_window, HWND_TOPMOST, x, y, width, height, SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_SHOWWINDOW);
}

}  // namespace shelltabs

