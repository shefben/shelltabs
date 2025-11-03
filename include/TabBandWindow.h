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
#include <unordered_map>
#include <functional>
#include <optional>

#include <commctrl.h>
#include <wrl/client.h>
#include <shobjidl.h>

#include "OptionsStore.h"
#include "PreviewOverlay.h"
#include "IconCache.h"
#include "TabManager.h"
#include "resource.h"
#include "ThemeNotifier.h"


namespace shelltabs {

class TabBand;
class ExplorerWindowHook;
struct GlowColorSet;

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
    void OnSavedGroupsChanged();

    void SetPreferredDockMode(TabBandDockMode mode);
    TabBandDockMode GetCurrentDockMode() const noexcept { return m_currentDockMode; }
    static uint32_t GetAvailableDockMask();
    static constexpr UINT_PTR SessionFlushTimerId() noexcept { return kSessionFlushTimerId; }

    enum class HitType {
        kNone,
        kWhitespace,
        kGroupHeader,
        kTab,
        kNewTab,
    };

    struct HitInfo {
        bool hit = false;
        size_t itemIndex = 0;
        HitType type = HitType::kNone;
        TabLocation location;
        bool before = false;
        bool after = false;
        bool closeButton = false;

        bool IsWhitespace() const noexcept {
            return hit && (type == HitType::kWhitespace || type == HitType::kNewTab);
        }
        bool IsTab() const noexcept { return hit && type == HitType::kTab && location.IsValid(); }
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
        uint64_t stableId = 0;
        RECT bounds{};
        bool firstInGroup = false;
        int badgeWidth = 0;
        IconCache::Reference icon;
        int iconWidth = 0;
        int iconHeight = 0;
        bool hasGroupHeader = false;
        TabViewItem groupHeader{};
        bool collapsedPlaceholder = false;
        bool indicatorHandle = false;
        size_t index = 0;
        int row = 0;
        size_t reuseSourceIndex = std::numeric_limits<size_t>::max();
        bool reusedIconMetrics = false;
    };

    struct TabLocationHash {
        size_t operator()(const TabLocation& location) const noexcept {
            const size_t group = std::hash<int>{}(location.groupIndex);
            const size_t tab = std::hash<int>{}(location.tabIndex);
            return group ^ (tab + 0x9e3779b97f4a7c15ULL + (group << 6) + (group >> 2));
        }
    };

    struct TabLocationEqual {
        bool operator()(const TabLocation& lhs, const TabLocation& rhs) const noexcept {
            return lhs.groupIndex == rhs.groupIndex && lhs.tabIndex == rhs.tabIndex;
        }
    };

    struct TabPaintMetrics {
        RECT itemBounds{};
        RECT tabBounds{};
        RECT closeButton{};
        int textLeft = 0;
        int textRight = 0;
        int iconLeft = 0;
        int iconWidth = 0;
        int iconHeight = 0;
        int islandIndicator = 0;
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

    struct LayoutResult {
        std::vector<VisualItem> items;
        int rowCount = 0;
        RECT newTabBounds{};
        bool newTabVisible = false;
    };

    struct VisualItemReuseContext {
        std::vector<VisualItem>* source = nullptr;
        std::unordered_map<uint64_t, std::vector<size_t>> indexByKey;
        std::vector<bool> reserved;
    };

    struct LayoutDiffStats {
        size_t inserted = 0;
        size_t removed = 0;
        size_t moved = 0;
        size_t updated = 0;
        std::vector<RECT> invalidRects;
        std::vector<size_t> removedIndices;
        std::vector<size_t> matchedOldIndices;
    };

    struct RedrawMetrics {
        double incrementalTotalMs = 0.0;
        uint64_t incrementalCount = 0;
        double fullTotalMs = 0.0;
        uint64_t fullCount = 0;
        double lastDurationMs = 0.0;
        bool lastWasIncremental = false;
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
        RECT indicatorRect{};
        RECT previewRect{};
    };

    struct ExternalDropState {
        bool active = false;
        DropTarget target{};
        TabBandWindow* source = nullptr;
        RECT indicatorRect{};
        RECT previewRect{};
    };

    struct BrushHandle {
        BrushHandle() = default;
        explicit BrushHandle(HBRUSH value) noexcept : handle(value) {}
        BrushHandle(const BrushHandle&) = delete;
        BrushHandle& operator=(const BrushHandle&) = delete;
        BrushHandle(BrushHandle&& other) noexcept : handle(other.handle) { other.handle = nullptr; }
        BrushHandle& operator=(BrushHandle&& other) noexcept {
            if (this != &other) {
                Reset();
                handle = other.handle;
                other.handle = nullptr;
            }
            return *this;
        }
        ~BrushHandle() { Reset(); }

        void Reset(HBRUSH value = nullptr) noexcept {
            if (handle && handle != value) {
                DeleteObject(handle);
            }
            handle = value;
        }

        [[nodiscard]] HBRUSH Get() const noexcept { return handle; }
        explicit operator bool() const noexcept { return handle != nullptr; }

    private:
        HBRUSH handle = nullptr;
    };

    struct PenHandle {
        PenHandle() = default;
        explicit PenHandle(HPEN value) noexcept : handle(value) {}
        PenHandle(const PenHandle&) = delete;
        PenHandle& operator=(const PenHandle&) = delete;
        PenHandle(PenHandle&& other) noexcept : handle(other.handle) { other.handle = nullptr; }
        PenHandle& operator=(PenHandle&& other) noexcept {
            if (this != &other) {
                Reset();
                handle = other.handle;
                other.handle = nullptr;
            }
            return *this;
        }
        ~PenHandle() { Reset(); }

        void Reset(HPEN value = nullptr) noexcept {
            if (handle && handle != value) {
                DeleteObject(handle);
            }
            handle = value;
        }

        [[nodiscard]] HPEN Get() const noexcept { return handle; }
        explicit operator bool() const noexcept { return handle != nullptr; }

    private:
        HPEN handle = nullptr;
    };

    struct PenKey {
        COLORREF color = 0;
        int width = 1;
        int style = PS_SOLID;

        bool operator==(const PenKey& other) const noexcept {
            return color == other.color && width == other.width && style == other.style;
        }
    };

    struct PenKeyHash {
        size_t operator()(const PenKey& key) const noexcept {
            size_t hash = std::hash<DWORD>{}(key.color);
            hash ^= std::hash<int>{}(key.width) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
            hash ^= std::hash<int>{}(key.style) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
            return hash;
        }
    };

    struct CachedGroupOutlines {
        std::vector<GroupOutline> outlines;
        bool valid = false;
    };

    struct RebarColorScheme {
        COLORREF background = CLR_DEFAULT;
        COLORREF foreground = CLR_DEFAULT;

        bool operator==(const RebarColorScheme& other) const noexcept {
            return background == other.background && foreground == other.foreground;
        }

        bool operator!=(const RebarColorScheme& other) const noexcept { return !(*this == other); }
    };

    struct ThemePalette {
        COLORREF rebarGradientTop = 0;
        COLORREF rebarGradientBottom = 0;
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
        bool rebarGradientValid = false;
    };

    HWND m_hwnd = nullptr;
    HWND m_newTabButton = nullptr;
    HWND m_parentRebar = nullptr;
    HWND m_parentFrame = nullptr;
    TabBand* m_owner = nullptr;

    RECT m_clientRect{};
    std::vector<TabViewItem> m_tabData;
    std::unordered_map<TabLocation, size_t, TabLocationHash, TabLocationEqual> m_tabLocationIndex;
    uint32_t m_tabLayoutVersion = 0;
    std::vector<VisualItem> m_items;
    std::vector<RECT> m_progressRects;
    std::vector<size_t> m_activeProgressIndices;
    size_t m_activeProgressCount = 0;
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
    bool m_highContrast = false;
    bool m_refreshingTheme = false;
    bool m_windowDarkModeInitialized = false;
    bool m_windowDarkModeValue = false;
    bool m_newTabButtonHot = false;
    bool m_newTabButtonPressed = false;
    bool m_newTabButtonKeyboardPressed = false;
    bool m_newTabButtonTrackingMouse = false;
    bool m_newTabButtonCommandPending = false;
    COLORREF m_accentColor = RGB(0, 120, 215);
    ExternalDropState m_externalDrop;
    ThemePalette m_themePalette;
    int m_toolbarGripWidth = 14;
    size_t m_hotCloseIndex = std::numeric_limits<size_t>::max();
    bool m_mouseTracking = false;
    int m_rebarBandIndex = -1;
    bool m_rebarZOrderTop = false;
    Microsoft::WRL::ComPtr<IDropTarget> m_dropTarget;
    bool m_dropTargetRegistered = false;
    bool m_dropTargetRegistrationPending = false;
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
    ThemeNotifier m_themeNotifier;
    ThemeNotifier::ThemeColors m_themeColors;
        bool m_rebarSubclassed = false;
        struct EmptyIslandPlus {
                int   groupIndex = -1;
                RECT  plus{};
                RECT  placeholder{};
        };
        bool m_rebarIntegrationDirty = true;
        HWND m_lastIntegratedRebar = nullptr;
        HWND m_lastIntegratedFrame = nullptr;
        std::optional<RebarColorScheme> m_lastRebarColors;
        bool m_rebarNeedsRepaint = false;

	// Render-time cache of empty-island "+" hit targets
        std::vector<EmptyIslandPlus> m_emptyIslandPlusButtons;
        RECT m_newTabBounds{};
	// Site/browser for current Explorer window
        Microsoft::WRL::ComPtr<IUnknown> m_siteUnknown;
        Microsoft::WRL::ComPtr<IServiceProvider> m_siteSp;
        std::shared_ptr<ExplorerWindowHook> m_windowHook;

        HDC m_backBufferDC = nullptr;
        HBITMAP m_backBufferBitmap = nullptr;
        HGDIOBJ m_backBufferOldBitmap = nullptr;
        SIZE m_backBufferSize{};

        TabBandDockMode m_preferredDockMode = TabBandDockMode::kAutomatic;
        TabBandDockMode m_currentDockMode = TabBandDockMode::kAutomatic;
        bool m_nextRedrawIncremental = false;
        RedrawMetrics m_redrawMetrics{};
        int m_lastAppliedRowCount = 0;
        bool m_closeButtonSizeCached = false;
        int m_cachedCloseButtonSize = 0;
        UINT m_cachedCloseButtonDpi = 0;

        mutable CachedGroupOutlines m_groupOutlineCache;
        mutable std::unordered_map<COLORREF, BrushHandle> m_brushCache;
        mutable std::unordered_map<PenKey, PenHandle, PenKeyHash> m_penCache;

        // Utilities
        // Helpers
        bool FindEmptyIslandPlusAt(POINT pt, int* outGroupIndex) const;
        void DrawEmptyIslandPluses(HDC dc) const;
        LayoutResult BuildLayoutItems(const std::vector<TabViewItem>& items,
                                      VisualItemReuseContext* reuseContext = nullptr);
        LayoutDiffStats ComputeLayoutDiff(std::vector<VisualItem>& oldItems,
                                          std::vector<VisualItem>& newItems) const;
        void ApplyPreservedVisualItems(const std::vector<VisualItem>& preserved,
                                       std::vector<VisualItem>& current,
                                       const LayoutDiffStats& diff) const;
        void DestroyVisualItemResources(std::vector<VisualItem>& items);
        void RecordRedrawDuration(double milliseconds, bool incremental);

	// Rebar background control
	void InstallRebarDarkSubclass();
	static LRESULT CALLBACK RebarSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);

    void Layout(int width, int height);
    void RebuildLayout();
    void Draw(HDC dc);
    void PaintSurface(HDC dc, const RECT& windowRect) const;
    void DrawBackground(HDC dc, const RECT& bounds) const;
    bool DrawRebarThemePart(HDC dc, const RECT& bounds, int partId, int stateId, bool suppressFallback,
                            const GlowColorSet* overrideColors = nullptr) const;
    GlowColorSet BuildRebarGlowColors(const ThemePalette& palette) const;
    void DrawGroupHeader(HDC dc, const VisualItem& item) const;
    void DrawTab(HDC dc, const VisualItem& item) const;
    void DrawGroupOutlines(HDC dc, const std::vector<GroupOutline>& outlines) const;
    void DrawTabProgress(HDC dc, const VisualItem& item, const TabPaintMetrics& metrics,
                         COLORREF background) const;
    void DrawDropIndicator(HDC dc) const;
    void DrawDragVisual(HDC dc) const;
    void ClearVisualItems();
    void ReleaseBackBuffer();
    void ClearExplorerContext();
    IconCache::Reference LoadItemIcon(const TabViewItem& item, UINT iconFlags) const;
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
    void RefreshProgressState(const std::vector<TabLocation>& prioritizedTabs);
    void RefreshProgressState(const TabProgressUpdatePayload* payload);
    void RefreshProgressState(const std::vector<TabLocation>& prioritizedTabs,
                              const TabProgressUpdatePayload* payload);
    void UpdateProgressAnimationState();
    bool AnyProgressActive() const;
    void HandleProgressTimer();
    TabManager* ResolveManager() const noexcept;
    void RegisterShellNotifications();
    void UnregisterShellNotifications();
    void OnShellNotify(WPARAM wParam, LPARAM lParam);
    void UpdateCloseButtonHover(const POINT& pt);
    void ClearCloseButtonHover();

    void InvalidateRebarIntegration();
    bool NeedsRebarIntegration() const;
    void EnsureRebarIntegration();
    void RefreshRebarMetrics();
    void FlushRebarRepaint();
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
    RECT ComputeDropIndicatorRect(const DropTarget& target) const;
    RECT ComputeDropPreviewRect(const DropTarget& target) const;
    bool TryGetGroupBounds(int groupIndex, RECT* bounds) const;
    bool TryGetTabBounds(int groupIndex, int tabIndex, RECT* bounds) const;
    void InvalidateDropRegions(const RECT& previousIndicator, const RECT& currentIndicator,
                               const RECT& previousPreview, const RECT& currentPreview);
    void ApplyDropTargetChange(const DropTarget& previous, const DropTarget& current,
                               RECT& indicatorRectStorage, RECT& previewRectStorage);
    void ApplyInternalDropTarget(const DropTarget& previous, const DropTarget& current);
    void ApplyExternalDropTarget(const DropTarget& previous, const DropTarget& current, TabBandWindow* sourceWindow);
    void UpdateDropTarget(const POINT& pt);
    void CompleteDrop();
    DropTarget ComputeDropTarget(const POINT& pt, const HitInfo& origin) const;
    int ComputeIndicatorXForInsertion(int groupIndex, int tabIndex) const;
    void AdjustDropTargetForPinned(const HitInfo& origin, DropTarget& target) const;
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
    size_t FindTabDataIndex(TabLocation location) const;
    size_t FindGroupHeaderIndex(int groupIndex) const;
    void RebuildTabLocationIndex();
    TabPaintMetrics ComputeTabPaintMetrics(const VisualItem& item) const;
    bool ComputeProgressBounds(const VisualItem& item, const TabPaintMetrics& metrics, RECT* out) const;
    void EnsureProgressRectCache();
    void RebuildProgressRectCache();
    void RecomputeActiveProgressCount();
    void InvalidateProgressForIndices(const std::vector<size_t>& indices);
    void InvalidateActiveProgress();
    COLORREF ResolveTabBackground(const TabViewItem& item) const;
    COLORREF ResolveGroupBackground(const TabViewItem& item) const;
    COLORREF ResolveTextColor(COLORREF background) const;
    COLORREF ResolveTabTextColor(bool selected, COLORREF background) const;
    void ApplyOptionColorOverrides();
    COLORREF ResolveGroupTextColor(const TabViewItem& item, COLORREF background) const;
    const std::vector<GroupOutline>& BuildGroupOutlines() const;
    void InvalidateGroupOutlineCache();
    void RebuildGroupOutlineCache() const;
    std::vector<GroupOutline> ComputeGroupOutlines() const;
    bool DropPreviewAffectsIndicators(const DropTarget& target) const;
    void OnDropPreviewTargetChanged(const DropTarget& previous, const DropTarget& current);
    RECT ComputeCloseButtonRect(const VisualItem& item) const;
    HBITMAP CreateDragVisualBitmap(const VisualItem& item, SIZE* size) const;
    void UpdateDragOverlay(const POINT& clientPt, const POINT& screenPt);
    void HideDragOverlay(bool destroy);
    void CloseThemeHandles();
    void ResetCloseButtonMetrics();
    void UpdateNewTabButtonTheme();
    void PaintNewTabButton(HWND hwnd, HDC dc) const;
    void HandleNewTabButtonMouseMove(HWND hwnd);
    void HandleNewTabButtonMouseLeave(HWND hwnd);
    void HandleNewTabButtonLButtonDown(HWND hwnd);
    void HandleNewTabButtonLButtonUp(HWND hwnd, POINT pt);
    void HandleNewTabButtonCaptureLost();
    void HandleNewTabButtonFocusChanged(HWND hwnd, bool focused);
    void HandleNewTabButtonKeyDown(HWND hwnd, UINT key, bool repeat);
    void HandleNewTabButtonKeyUp(HWND hwnd, UINT key);
    void TriggerNewTabButtonAction();
    bool IsSystemDarkMode() const;
    void UpdateAccentColor();
    void ResetThemePalette();
    void DrawPinnedGlyph(HDC dc, const RECT& tabRect, int x, COLORREF color) const;
    void UpdateThemePalette();
    void UpdateToolbarMetrics();
    void HandleDpiChanged(UINT dpiX, UINT dpiY, const RECT* suggestedRect);
    [[nodiscard]] HBRUSH GetCachedBrush(COLORREF color) const;
    [[nodiscard]] HPEN GetCachedPen(COLORREF color, int width = 1, int style = PS_SOLID) const;
    void ClearGdiCache();

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

    void EnsureDropTargetRegistered();
    void ScheduleDropTargetRegistrationRetry();

    class BandDropTarget;
    friend class BandDropTarget;
    friend LRESULT CALLBACK NewTabButtonWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
#if defined(SHELLTABS_BUILD_TESTS)
    friend struct TabBandWindowDiffTestHarness;
#endif

    static constexpr UINT_PTR kDropHoverTimerId = 0x5348;  // 'SH'
    static constexpr UINT_PTR kSessionFlushTimerId = 0x5346;  // 'SF'
    static constexpr UINT_PTR kProgressTimerId = 0x5349;   // 'SI'

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
};

constexpr UINT WM_SHELLTABS_CLOSETAB = WM_APP + 42;
constexpr UINT WM_SHELLTABS_DEFER_NAVIGATE = WM_APP + 43;
constexpr UINT WM_SHELLTABS_PREVIEW_READY = WM_APP + 64;
constexpr UINT WM_SHELLTABS_REGISTER_DRAGDROP = WM_APP + 65;
constexpr ULONG_PTR SHELLTABS_COPYDATA_OPEN_FOLDER = 'STNT';

}  // namespace shelltabs
