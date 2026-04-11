#pragma once

#include <windows.h>
#include <gdiplus.h>
#include <string>
#include <vector>
#include <functional>

#include "TaskbarTabProvider.h"

namespace shelltabs {

class TaskbarTabListPopup {
public:
    TaskbarTabListPopup();
    ~TaskbarTabListPopup();

    void Create();
    void Destroy();

    using TabActivatedCallback = std::function<void(HWND explorerHwnd, int groupIndex, int tabIndex)>;
    void SetTabActivatedCallback(TabActivatedCallback cb);

    void ShowForWindow(const TaskbarTabSnapshot& snapshot, const RECT& thumbnailRect, HWND taskbarHwnd);
    void Hide();
    bool IsVisible() const;
    HWND GetHwnd() const { return m_hwnd; }

private:
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    void OnPaint();
    void OnMouseMove(POINT clientPt);
    void OnMouseLeave();
    void OnLButtonDown(POINT clientPt);

    void PaintToLayered();
    void DrawTabEntry(Gdiplus::Graphics& g, int index, const TaskbarTabEntry& entry,
                      const Gdiplus::RectF& bounds, bool isHovered, bool isSelected);
    Gdiplus::Color GetBackgroundColor() const;
    Gdiplus::Color GetTextColor() const;
    Gdiplus::Color GetSelectedColor() const;
    Gdiplus::Color GetHoverColor() const;
    Gdiplus::Color GetAccentColor() const;

    struct TabHitRect {
        Gdiplus::RectF bounds;
        int groupIndex;
        int tabIndex;
    };
    std::vector<TabHitRect> m_hitRects;
    int m_hoveredIndex = -1;

    RECT CalculatePopupRect(const RECT& thumbnailRect, HWND taskbarHwnd, int itemCount) const;
    int GetTaskbarEdge(HWND taskbarHwnd) const;

    HWND m_hwnd = nullptr;
    TaskbarTabSnapshot m_snapshot;
    HWND m_explorerHwnd = nullptr;
    TabActivatedCallback m_callback;
    bool m_trackingMouse = false;
    bool m_isDarkMode = false;
    int m_dpi = 96;

    int ItemHeight() const;
    int ItemPaddingX() const;
    int IconSize() const;
    int PopupWidth() const;
    int PopupPadding() const;
    int CornerRadius() const;
    int AccentBarWidth() const;

    static ATOM s_windowClass;
    static constexpr int kBaseItemHeight = 28;
    static constexpr int kBaseItemPadX = 8;
    static constexpr int kBaseIconSize = 16;
    static constexpr int kBasePopupWidth = 240;
    static constexpr int kBasePopupPadding = 4;
    static constexpr int kBaseCornerRadius = 8;
    static constexpr int kBaseAccentBar = 3;
    static constexpr int kMaxVisibleItems = 20;
};

}  // namespace shelltabs
