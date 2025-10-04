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
#include <wrl/client.h>

#include "TabManager.h"

namespace shelltabs {

class TabBand;
class TabToolbarDropTarget;

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
    friend class TabToolbarDropTarget;

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
    void ConfigureToolbarMetrics();
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
    void ApplyThemeToRibbonAncestors();
    bool PaintHostBackground(HDC dc) const;
    bool PaintToolbarBackground(HWND hwnd, HDC dc) const;
    bool ShouldUpdateThemeForSettingChange(LPARAM lParam) const;
    bool ExplorerHostPrefersDarkMode() const;
    bool IsDarkModePreferred() const;
    bool IsAmbientDark() const;
    void RegisterDropTarget();
    void RevokeDropTarget();
    void BeginDrag(int commandId, const POINT& screenPt);
    void UpdateDrag(const POINT& screenPt);
    void EndDrag(const POINT& screenPt, bool canceled);
    void CancelDrag();
    bool StartDragVisual(const POINT& screenPt);
    void DestroyDragImage();
    bool HandleShellContextMenuMessage(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* result);
    void ResetContextMenuState();
    bool IsPointInsideToolbar(const POINT& screenPt) const;
    TabLocation ComputeTabInsertLocation(const POINT& clientPt) const;
    int ComputeGroupInsertIndex(const POINT& clientPt) const;
    const TabViewItem* ItemFromPoint(const POINT& screenPt) const;
    TabLocation TabLocationFromPoint(const POINT& screenPt) const;
    UINT CurrentDpi() const;
    int GroupIndicatorWidth() const;
    int GroupIndicatorSpacing() const;
    int GroupIndicatorVisualWidth() const;
    COLORREF GroupIndicatorColor(const TabViewItem& item) const;
    int TabHorizontalPadding() const;
    int IconTextSpacing() const;
    int CloseButtonSpacing() const;
    int CloseButtonSize() const;
    RECT CloseButtonRect(const RECT& buttonRect) const;
    bool GetButtonRect(int commandId, RECT* rect) const;
    int GetButtonImage(int commandId) const;
    void InvalidateButton(int commandId) const;
    bool IsPointInCloseButton(int commandId, const POINT& screenPt, RECT* closeRectOut = nullptr) const;
    void ResetCloseTracking();
    void ResetCommandIgnore();
    void CloseTabCommand(int commandId);
    bool TryHandleCloseClick(const POINT& screenPt);
    int CalculateTabButtonWidth(const TabViewItem& item) const;
    int CalculateGroupHeaderWidth(const TabViewItem& item) const;
    int MeasureTabTextWidth(const std::wstring& text) const;
    std::wstring DisplayLabelForItem(const TabViewItem& item) const;
    void UpdateInsertMark(const POINT& screenPt);
    void ClearInsertMark();

    struct DragState {
        bool tracking = false;
        bool dragging = false;
        bool isGroup = false;
        int commandId = -1;
        TabLocation tabLocation{};
        int groupIndex = -1;
        POINT startPoint{};
        HIMAGELIST dragImage = nullptr;
        bool dragImageVisible = false;
        HWND dragImageWindow = nullptr;
        bool suppressCancel = false;
    };

    struct CloseButtonState {
        bool tracking = false;
        bool hot = false;
        int commandId = -1;
        RECT rect{};
    };

    struct ShellContextMenuState {
        Microsoft::WRL::ComPtr<IContextMenu> menu;
        Microsoft::WRL::ComPtr<IContextMenu2> menu2;
        Microsoft::WRL::ComPtr<IContextMenu3> menu3;
        HMENU menuHandle = nullptr;
        HMENU explorerSubMenu = nullptr;
        UINT idFirst = 0;
        UINT idLast = 0;
        TabLocation location{};
        POINT invokePoint{};

        bool IsActive() const noexcept { return menu || menu2 || menu3; }
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
    CloseButtonState m_closeState{};
    ShellContextMenuState m_contextMenuState{};
};

}  // namespace shelltabs
