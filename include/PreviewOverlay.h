#pragma once

#include <windows.h>

namespace shelltabs {

class PreviewOverlay {
public:
    PreviewOverlay() = default;
    ~PreviewOverlay();

    PreviewOverlay(const PreviewOverlay&) = delete;
    PreviewOverlay& operator=(const PreviewOverlay&) = delete;

    void SetOwner(HWND owner) noexcept { m_owner = owner; }
    bool Show(HWND owner, HBITMAP bitmap, const SIZE& size, const POINT& screenPt);
    void Hide(bool destroyWindow);
    void Destroy();

    HWND GetWindow() const noexcept { return m_window; }
    SIZE GetBitmapSize() const noexcept { return m_bitmapSize; }
    bool IsVisible() const noexcept { return m_visible; }

    void PositionRelativeToRect(const RECT& anchorRectScreen, const POINT& cursorScreenPt);

private:
    HWND EnsureWindow(HWND owner);
    static ATOM EnsureWindowClass();

    HWND m_owner = nullptr;
    HWND m_window = nullptr;
    SIZE m_bitmapSize{};
    bool m_visible = false;
};

}  // namespace shelltabs

