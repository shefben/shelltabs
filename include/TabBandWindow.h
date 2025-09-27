#pragma once

#include <windows.h>
#include <uxtheme.h>

#include <limits>
#include <string>
#include <utility>
#include <vector>

#include <wrl/client.h>

#include "TabManager.h"

namespace shelltabs {

class TabBand;

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

    struct HitInfo {
        bool hit = false;
        size_t itemIndex = 0;
        TabViewItemType type = TabViewItemType::kGroupHeader;
        TabLocation location;
        bool before = false;
        bool after = false;
        bool closeButton = false;
    };

    struct DropTarget {
        bool active = false;
        bool outside = false;
        bool group = false;
        int groupIndex = -1;
        int tabIndex = -1;
        int indicatorX = -1;
        bool newGroup = false;
        bool floating = false;
    };

private:
    struct VisualItem {
        TabViewItem data;
        RECT bounds{};
        bool firstInGroup = false;
        int badgeWidth = 0;
        HICON icon = nullptr;
        int iconWidth = 0;
        int iconHeight = 0;
        bool hasGroupHeader = false;
        TabViewItem groupHeader{};
        bool collapsedPlaceholder = false;
        bool indicatorHandle = false;
        size_t index = 0;
    };

    struct GroupOutline {
        int groupIndex = -1;
        RECT bounds{};
        COLORREF color = RGB(0, 0, 0);
        bool initialized = false;
        bool visible = false;
    };

    struct ExplorerContext {
        Microsoft::WRL::ComPtr<IContextMenu> menu;
        Microsoft::WRL::ComPtr<IContextMenu2> menu2;
        Microsoft::WRL::ComPtr<IContextMenu3> menu3;
        TabLocation location;
        UINT idFirst = 0;
        UINT idLast = 0;
    };

    struct DragState {
        bool tracking = false;
        bool dragging = false;
        HitInfo origin;
        POINT start{};
        DropTarget target{};
        POINT current{};
        bool hasCurrent = false;
        bool originSelected = false;
        bool closeClick = false;
        size_t closeItemIndex = 0;
        TabLocation closeLocation{};
        HWND overlay = nullptr;
        bool overlayVisible = false;
    };

    struct ExternalDropState {
        bool active = false;
        DropTarget target{};
        TabBandWindow* source = nullptr;
    };

    struct ThemePalette {
        COLORREF rebarBackground = 0;
        COLORREF borderTop = 0;
        COLORREF borderBottom = 0;
        COLORREF tabBase = 0;
        COLORREF tabSelectedBase = 0;
        COLORREF tabText = 0;
        COLORREF tabSelectedText = 0;
        COLORREF groupBase = 0;
        COLORREF groupText = 0;
        bool tabTextValid = false;
        bool tabSelectedTextValid = false;
        bool groupTextValid = false;
    };

    HWND m_hwnd = nullptr;
    HWND m_newTabButton = nullptr;
    HWND m_parentRebar = nullptr;
    TabBand* m_owner = nullptr;

    RECT m_clientRect{};
    std::vector<TabViewItem> m_tabData;
    std::vector<VisualItem> m_items;
    DragState m_drag;
    HitInfo m_contextHit;
    std::vector<std::pair<UINT, TabLocation>> m_hiddenTabCommands;
    std::vector<std::pair<UINT, std::wstring>> m_savedGroupCommands;
    ExplorerContext m_explorerContext;
    POINT m_lastContextPoint{};
    HTHEME m_tabTheme = nullptr;
    HTHEME m_rebarTheme = nullptr;
    HTHEME m_windowTheme = nullptr;
    bool m_darkMode = false;
    bool m_refreshingTheme = false;
    bool m_windowDarkModeInitialized = false;
    bool m_windowDarkModeValue = false;
    bool m_buttonDarkModeInitialized = false;
    bool m_buttonDarkModeValue = false;
    COLORREF m_accentColor = RGB(0, 120, 215);
    ExternalDropState m_externalDrop;
    ThemePalette m_themePalette;
    int m_toolbarGripWidth = 14;
    size_t m_hotCloseIndex = std::numeric_limits<size_t>::max();
    bool m_mouseTracking = false;
    int m_rebarBandIndex = -1;

    void Layout(int width, int height);
    void RebuildLayout();
    void Draw(HDC dc) const;
    void PaintSurface(HDC dc, const RECT& windowRect) const;
    void DrawBackground(HDC dc, const RECT& bounds) const;
    void DrawGroupHeader(HDC dc, const VisualItem& item) const;
    void DrawTab(HDC dc, const VisualItem& item) const;
    void DrawGroupOutlines(HDC dc, const std::vector<GroupOutline>& outlines) const;
    void DrawDropIndicator(HDC dc) const;
    void DrawDragVisual(HDC dc) const;
    void ClearVisualItems();
    void ClearExplorerContext();
    HICON LoadItemIcon(const TabViewItem& item) const;
    bool HandleExplorerMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT* result);
    void EnsureMouseTracking();
    void UpdateCloseButtonHover(const POINT& pt);
    void ClearCloseButtonHover();

    void EnsureRebarIntegration();
    void RefreshRebarMetrics();
    int FindRebarBandIndex() const;
    static bool IsRebarWindow(HWND hwnd);

    void HandleCommand(WPARAM wParam, LPARAM lParam);
    bool HandleMouseDown(const POINT& pt);
    bool HandleMouseUp(const POINT& pt);
    bool HandleMouseMove(const POINT& pt);
    bool HandleDoubleClick(const POINT& pt);
    void HandleFileDrop(HDROP drop);
    void CancelDrag();
    void UpdateDropTarget(const POINT& pt);
    void CompleteDrop();
    DropTarget ComputeDropTarget(const POINT& pt, const HitInfo& origin) const;
    void UpdateExternalDrag(const POINT& screenPt);
    bool TryCompleteExternalDrop();
    void HandleExternalDragUpdate();
    void HandleExternalDragLeave();
    void HandleExternalDropExecute();
    void RequestSelection(const HitInfo& hit);
    HitInfo HitTest(const POINT& pt) const;
    void ShowContextMenu(const POINT& pt);
    void PopulateHiddenTabsMenu(HMENU menu, int groupIndex);
    void PopulateSavedGroupsMenu(HMENU parent, bool addSeparator);
    int ResolveInsertGroupIndex() const;
    int GroupCount() const;
    const VisualItem* FindLastGroupHeader() const;
    const VisualItem* FindVisualForHit(const HitInfo& hit) const;
    int MeasureBadgeWidth(const TabViewItem& item, HDC dc) const;
    std::wstring BuildGitBadgeText(const TabViewItem& item) const;
    COLORREF ResolveTabBackground(const TabViewItem& item) const;
    COLORREF ResolveGroupBackground(const TabViewItem& item) const;
    COLORREF ResolveTextColor(COLORREF background) const;
    COLORREF ResolveTabTextColor(bool selected, COLORREF background) const;
    COLORREF ResolveGroupTextColor(const TabViewItem& item, COLORREF background) const;
    std::vector<GroupOutline> BuildGroupOutlines() const;
    RECT ComputeCloseButtonRect(const VisualItem& item) const;
    HBITMAP CreateDragVisualBitmap(const VisualItem& item, SIZE* size) const;
    void UpdateDragOverlay(const POINT& clientPt, const POINT& screenPt);
    void HideDragOverlay(bool destroy);
    void RefreshTheme();
    void CloseThemeHandles();
    void UpdateNewTabButtonTheme();
    bool IsSystemDarkMode() const;
    void UpdateAccentColor();
    void ResetThemePalette();
    void UpdateThemePalette();
    void UpdateToolbarMetrics();

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
};

constexpr UINT WM_SHELLTABS_CLOSETAB = WM_APP + 42;
constexpr UINT WM_SHELLTABS_DEFER_NAVIGATE = WM_APP + 43;
constexpr UINT WM_SHELLTABS_REFRESH_COLORIZER = WM_APP + 44;
constexpr UINT WM_SHELLTABS_REFRESH_GIT_STATUS = WM_APP + 45;
constexpr UINT WM_SHELLTABS_ENABLE_GIT_STATUS = WM_APP + 46;

enum : UINT_PTR {
    IDC_NEW_TAB = 1001,
    IDM_CLOSE_TAB = 40001,
    IDM_HIDE_TAB = 40002,
    IDM_DETACH_TAB = 40003,
    IDM_CLONE_TAB = 40004,
    IDM_TOGGLE_ISLAND = 40010,
    IDM_UNHIDE_ALL = 40011,
    IDM_NEW_ISLAND = 40012,
    IDM_DETACH_ISLAND = 40013,
    IDM_TOGGLE_SPLIT = 40014,
    IDM_SET_SPLIT_SECONDARY = 40015,
    IDM_CLEAR_SPLIT_SECONDARY = 40016,
    IDM_SWAP_SPLIT = 40017,
    IDM_OPEN_TERMINAL = 40018,
    IDM_OPEN_VSCODE = 40019,
    IDM_COPY_PATH = 40020,
    IDM_TOGGLE_ISLAND_HEADER = 40021,
    IDM_CREATE_SAVED_GROUP = 40022,
    IDM_HIDDEN_TAB_BASE = 41000,
    IDM_EXPLORER_CONTEXT_BASE = 42000,
    IDM_EXPLORER_CONTEXT_LAST = 42999,
    IDM_LOAD_SAVED_GROUP_BASE = 43000,
    IDM_LOAD_SAVED_GROUP_LAST = 43999,
};

}  // namespace shelltabs
