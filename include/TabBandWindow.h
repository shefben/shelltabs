#pragma once

// Ensure Windows 7+ APIs so INamespaceTreeControlCustomDraw and NSTCCUSTOMDRAW::clrText are available.
#if !defined(_WIN32_WINNT) || _WIN32_WINNT < 0x0601
#undef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif
#include <sdkddkver.h>  // NTDDI_* macros
#if !defined(NTDDI_VERSION) || NTDDI_VERSION < NTDDI_WIN7
#undef NTDDI_VERSION
#define NTDDI_VERSION NTDDI_WIN7
#endif

#include <windows.h>
#include <uxtheme.h>

#include <limits>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include <commctrl.h>
#include <wrl/client.h>
#include <shobjidl.h>

#include "OptionsStore.h"
#include "PreviewOverlay.h"
#include "TabManager.h"
#include "resource.h"


namespace shelltabs {

class TabBand;
class ExplorerWindowHook;

class TabBandWindow {
public:
    explicit TabBandWindow(TabBand* owner);
    ~TabBandWindow();

    HWND Create(HWND parent);
    void Destroy();

    HWND GetHwnd() const noexcept { return m_hwnd; }
	STDMETHOD(SetSite)(IUnknown* pUnkSite);
	STDMETHOD(GetSite)(REFIID riid, void** ppvSite);
    void Show(bool show);
    void SetTabs(const std::vector<TabViewItem>& items);
    const std::vector<TabViewItem>& GetTabData() const noexcept { return m_tabData; }
    bool HasFocus() const;
    void FocusTab();
    void RefreshTheme();

    void SetPreferredDockMode(TabBandDockMode mode);
    TabBandDockMode GetCurrentDockMode() const noexcept { return m_currentDockMode; }
    static uint32_t GetAvailableDockMask();

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
        int row = 0;
    };

    struct GroupOutline {
        int groupIndex = -1;
        int row = 0;
        RECT bounds{};
        COLORREF color = RGB(0, 0, 0);
        bool initialized = false;
        bool visible = false;
        TabGroupOutlineStyle style = TabGroupOutlineStyle::kSolid;
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
    bool m_rebarZOrderTop = false;
    Microsoft::WRL::ComPtr<IDropTarget> m_dropTarget;
    HitInfo m_dropHoverHit;
    bool m_dropHoverHasFileData = false;
    bool m_dropHoverTimerActive = false;
    COLORREF m_progressStartColor = RGB(0, 120, 215);
    COLORREF m_progressEndColor = RGB(0, 153, 255);
    PreviewOverlay m_previewOverlay;
    size_t m_previewItemIndex = std::numeric_limits<size_t>::max();
    bool m_previewVisible = false;
    uint64_t m_previewRequestId = 0;
    POINT m_previewAnchorPoint{};
    UINT m_shellNotifyMessage = 0;
    ULONG m_shellNotifyId = 0;
    bool m_progressTimerActive = false;
        int m_lastRowCount = 1;  // tracks wrapped rows for height calc
        // track if we've installed the subclass
        bool m_rebarSubclassed = false;
	struct EmptyIslandPlus {
		int   groupIndex = -1;
		RECT  rect{};   // click target for "+"
	};

	// Render-time cache of empty-island "+" hit targets
	std::vector<EmptyIslandPlus> m_emptyIslandPlusButtons;
	// Site/browser for current Explorer window
        Microsoft::WRL::ComPtr<IUnknown> m_siteUnknown;
        Microsoft::WRL::ComPtr<IServiceProvider> m_siteSp;
        std::shared_ptr<ExplorerWindowHook> m_windowHook;

        TabBandDockMode m_preferredDockMode = TabBandDockMode::kAutomatic;
        TabBandDockMode m_currentDockMode = TabBandDockMode::kAutomatic;

        // Utilities
        // Helpers
	bool FindEmptyIslandPlusAt(POINT pt, int* outGroupIndex) const;
	void DrawEmptyIslandPluses(HDC dc) const;

	// Rebar background control
	void InstallRebarDarkSubclass();
	static LRESULT CALLBACK RebarSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);

    void Layout(int width, int height);
    void RebuildLayout();
    void Draw(HDC dc) const;
    void PaintSurface(HDC dc, const RECT& windowRect) const;
    void DrawBackground(HDC dc, const RECT& bounds) const;
    void DrawGroupHeader(HDC dc, const VisualItem& item) const;
    void DrawTab(HDC dc, const VisualItem& item) const;
    void DrawGroupOutlines(HDC dc, const std::vector<GroupOutline>& outlines) const;
    void DrawTabProgress(HDC dc, const VisualItem& item, int left, int right, const RECT& tabRect,
                         COLORREF background) const;
    void DrawDropIndicator(HDC dc) const;
    void DrawDragVisual(HDC dc) const;
    void ClearVisualItems();
    void ClearExplorerContext();
    HICON LoadItemIcon(const TabViewItem& item, UINT iconFlags) const;
    bool HandleExplorerMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT* result);
    void EnsureMouseTracking(const POINT& pt);
    void UpdateHoverPreview(const POINT& pt);
    void HandleMouseHover(const POINT& pt);
    void ShowPreviewForItem(size_t index, const POINT& screenPt);
    void HidePreviewWindow(bool destroy);
    void PositionPreviewWindow(const VisualItem& item, const POINT& screenPt);
    void HandlePreviewReady(uint64_t requestId);
    void CancelPreviewRequest();
    void RefreshProgressState();
    void UpdateProgressAnimationState();
    bool AnyProgressActive() const;
    void HandleProgressTimer();
    void RegisterShellNotifications();
    void UnregisterShellNotifications();
    void OnShellNotify(WPARAM wParam, LPARAM lParam);
    void UpdateCloseButtonHover(const POINT& pt);
    void ClearCloseButtonHover();

    void EnsureRebarIntegration();
    void RefreshRebarMetrics();
    int FindRebarBandIndex() const;
    static bool IsRebarWindow(HWND hwnd);
    bool DrawRebarBackground(HDC dc, const RECT& bounds) const;
    void OnParentRebarMetricsChanged();
    void EnsureToolbarZOrder();
    void UpdateRebarColors();
        void AdjustBandHeightToRow();
    bool BandHasRebarGrip() const;

    void HandleCommand(WPARAM wParam, LPARAM lParam);
    bool HandleMouseDown(const POINT& pt);
    bool HandleMouseUp(const POINT& pt);
    bool HandleMouseMove(const POINT& pt);
    bool HandleDoubleClick(const POINT& pt);
    void HandleFileDrop(HDROP drop, bool ownsHandle);
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
    bool HasAnyTabs() const;
    int ResolveInsertGroupIndex() const;
    int GroupCount() const;
    const VisualItem* FindLastGroupHeader() const;
    const VisualItem* FindVisualForHit(const HitInfo& hit) const;
    COLORREF ResolveTabBackground(const TabViewItem& item) const;
    COLORREF ResolveGroupBackground(const TabViewItem& item) const;
    COLORREF ResolveTextColor(COLORREF background) const;
    COLORREF ResolveTabTextColor(bool selected, COLORREF background) const;
    void ApplyOptionColorOverrides();
    COLORREF ResolveGroupTextColor(const TabViewItem& item, COLORREF background) const;
    std::vector<GroupOutline> BuildGroupOutlines() const;
    RECT ComputeCloseButtonRect(const VisualItem& item) const;
    HBITMAP CreateDragVisualBitmap(const VisualItem& item, SIZE* size) const;
    void UpdateDragOverlay(const POINT& clientPt, const POINT& screenPt);
    void HideDragOverlay(bool destroy);
    void CloseThemeHandles();
    void UpdateNewTabButtonTheme();
    bool IsSystemDarkMode() const;
    void UpdateAccentColor();
    void ResetThemePalette();
    void UpdateThemePalette();
    void UpdateToolbarMetrics();

    void UpdateDropHoverState(const HitInfo& hit, bool hasFileData);
    void ClearDropHoverState();
    void StartDropHoverTimer();
    void CancelDropHoverTimer();
    void OnDropHoverTimer();
    bool IsSameHit(const HitInfo& a, const HitInfo& b) const;
    bool IsSelectedTabHit(const HitInfo& hit) const;
    bool HasFileDropData(IDataObject* dataObject) const;
    DWORD ComputeFileDropEffect(DWORD keyState, bool hasFileData) const;

    HRESULT OnNativeDragEnter(IDataObject* dataObject, DWORD keyState, const POINTL& point, DWORD* effect);
    HRESULT OnNativeDragOver(DWORD keyState, const POINTL& point, DWORD* effect);
    HRESULT OnNativeDragLeave();
    HRESULT OnNativeDrop(IDataObject* dataObject, DWORD keyState, const POINTL& point, DWORD* effect);

    class BandDropTarget;
    friend class BandDropTarget;

    static constexpr UINT_PTR kDropHoverTimerId = 0x5348;  // 'SH'
    static constexpr UINT_PTR kSessionFlushTimerId = 0x5346;  // 'SF'
    static constexpr UINT_PTR kProgressTimerId = 0x5349;   // 'SI'

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
};

constexpr UINT WM_SHELLTABS_CLOSETAB = WM_APP + 42;
constexpr UINT WM_SHELLTABS_DEFER_NAVIGATE = WM_APP + 43;
constexpr UINT WM_SHELLTABS_PREVIEW_READY = WM_APP + 64;
constexpr ULONG_PTR SHELLTABS_COPYDATA_OPEN_FOLDER = 'STNT';

}  // namespace shelltabs
