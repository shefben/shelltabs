#pragma once

#include <windows.h>

#ifndef _WIN32_IE
#define _WIN32_IE 0x0601
#elif _WIN32_IE < 0x0601
#undef _WIN32_IE
#define _WIN32_IE 0x0601
#endif

#include <CommCtrl.h>
#include <OleIdl.h>
#include <uxtheme.h>

#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "TabManager.h"

namespace shelltabs {

class TabBand;

constexpr UINT WM_SHELLTABS_CLOSETAB = WM_APP + 42;
constexpr UINT WM_SHELLTABS_DEFER_NAVIGATE = WM_APP + 43;
constexpr UINT WM_SHELLTABS_REFRESH_COLORIZER = WM_APP + 44;
constexpr UINT WM_SHELLTABS_REFRESH_GIT_STATUS = WM_APP + 45;
constexpr UINT WM_SHELLTABS_ENABLE_GIT_STATUS = WM_APP + 46;

class TabBandWindow {
public:
    explicit TabBandWindow(TabBand* owner);
    ~TabBandWindow();

    HWND Create(HWND parent);
    void Destroy();

    HWND GetHwnd() const noexcept { return m_hwnd; }

    void Show(bool show);
    void SetTabs(const std::vector<TabViewItem>& items);
    bool HasFocus() const;
    void FocusTab();

private:
    HWND m_hwnd = nullptr;
    HWND m_toolbar = nullptr;
    TabBand* m_owner = nullptr;
    std::vector<TabViewItem> m_tabData;
    HIMAGELIST m_imageList = nullptr;
    std::unordered_map<int, TabLocation> m_commandMap;
    std::unordered_map<int, size_t> m_commandToIndex;
    int m_nextCommandId = 41000;

    void EnsureToolbar();
    void DestroyToolbar();
    void RebuildToolbar();
    void ClearToolbar();
    void ClearImageList();
    int AppendImage(HICON icon);
    void UpdateCheckedState();
    void HandleToolbarCommand(int commandId);
    void HandleContextMenu(int commandId, const POINT& screenPt);
    void HandleMiddleClick(int commandId);
    void HandleLButtonDown(int commandId);
    void HandleFilesDropped(TabLocation location, const std::vector<std::wstring>& paths, bool move);
    void HandleMouseMove(const POINT& screenPt);
    void HandleLButtonUp(const POINT& screenPt);
    void HandleTooltipRequest(NMTTDISPINFOW* info);
    void RelayFocusToToolbar();
    int CommandIdFromButtonIndex(int index) const;
    TabLocation LocationForCommand(int commandId) const;
    const TabViewItem* ItemForCommand(int commandId) const;
    LRESULT HandleToolbarCustomDraw(NMTBCUSTOMDRAW* customDraw);
    void UpdateTheme();
    void ApplyThemeToToolbar();
    bool PaintHostBackground(HDC dc) const;
    bool PaintToolbarBackground(HWND hwnd, HDC dc) const;
    bool ShouldUpdateThemeForSettingChange(LPARAM lParam) const;
    bool IsDarkModePreferred() const;
    void RegisterDropTarget();
    void RevokeDropTarget();
    void BeginDrag(int commandId, const POINT& screenPt);
    void UpdateDrag(const POINT& screenPt);
    void EndDrag(const POINT& screenPt, bool canceled);
    void CancelDrag();
    TabLocation ComputeTabInsertLocation(const POINT& clientPt) const;
    int ComputeGroupInsertIndex(const POINT& clientPt) const;
    const TabViewItem* ItemFromPoint(const POINT& screenPt) const;
    TabLocation TabLocationFromPoint(const POINT& screenPt) const;

    struct DragState {
        bool tracking = false;
        bool dragging = false;
        bool isGroup = false;
        int commandId = -1;
        TabLocation tabLocation{};
        int groupIndex = -1;
        POINT startPoint{};
    };

    struct ToolbarTheme {
        COLORREF background = RGB(249, 249, 249);
        COLORREF hover = RGB(229, 229, 229);
        COLORREF pressed = RGB(212, 212, 212);
        COLORREF checked = RGB(200, 200, 200);
        COLORREF text = RGB(32, 32, 32);
        COLORREF textDisabled = RGB(150, 150, 150);
        COLORREF groupHeaderBackground = RGB(240, 240, 240);
        COLORREF groupHeaderHover = RGB(225, 225, 225);
        COLORREF groupHeaderText = RGB(96, 96, 96);
        COLORREF highlight = RGB(0, 120, 215);
        COLORREF border = RGB(200, 200, 200);
        COLORREF separator = RGB(220, 220, 220);

        bool operator==(const ToolbarTheme& other) const noexcept {
            return background == other.background && hover == other.hover && pressed == other.pressed &&
                   checked == other.checked && text == other.text && textDisabled == other.textDisabled &&
                   groupHeaderBackground == other.groupHeaderBackground &&
                   groupHeaderHover == other.groupHeaderHover && groupHeaderText == other.groupHeaderText &&
                   highlight == other.highlight && border == other.border && separator == other.separator;
        }

        bool operator!=(const ToolbarTheme& other) const noexcept { return !(*this == other); }
    };

    ToolbarTheme CalculateTheme(bool darkMode) const;
    static void FillRectColor(HDC dc, const RECT& rect, COLORREF color);
    static void FrameRectColor(HDC dc, const RECT& rect, COLORREF color);

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                    DWORD_PTR refData);
    static LRESULT CALLBACK ToolbarWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                           DWORD_PTR refData);

    ToolbarTheme m_theme{};
    bool m_darkModeEnabled = false;
    DragState m_dragState{};
    bool m_ignoreNextCommand = false;
    int m_ignoredCommandId = -1;
    IDropTarget* m_dropTarget = nullptr;
};

}  // namespace shelltabs
