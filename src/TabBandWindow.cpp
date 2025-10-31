#ifndef NOMINMAX
#define NOMINMAX
#endif
#include "TabBandWindow.h"

#include <algorithm>  // for 2-arg std::max/std::min

#include <chrono>
#include <cmath>
#include <cstdint>
#include <functional>
#include <memory>
#include <mutex>
#include <optional>
#include <limits>
#include <string>
#include <unordered_map>
#include <vector>
#include <atomic>

#include <CommCtrl.h>
#include <windowsx.h>
#include <shellapi.h>
#include <ShlObj.h>
#include <shobjidl_core.h>
#include <shlguid.h>
#include <Shlwapi.h>
#include <Ole2.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <winreg.h>
#include <dwmapi.h>
#include <cwchar>

#include "Logging.h"
#include "Module.h"
#include "OptionsStore.h"
#include "ShellTabsMessages.h"
#include "TabBand.h"
#include "PreviewCache.h"
#include "Utilities.h"

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
using Microsoft::WRL::ComPtr;

namespace shelltabs {

namespace {

std::atomic<uint32_t> g_availableDockMask{0};

TabBandDockMode DockModeFromRebarStyle(DWORD style) {
    if ((style & CCS_VERT) != 0) {
        if ((style & CCS_RIGHT) != 0) {
            return TabBandDockMode::kRight;
        }
        return TabBandDockMode::kLeft;
    }
    if ((style & CCS_BOTTOM) != 0) {
        return TabBandDockMode::kBottom;
    }
    return TabBandDockMode::kTop;
}

size_t HashCombine(size_t seed, size_t value) noexcept {
    return seed ^ (value + 0x9e3779b97f4a7c15ull + (seed << 6) + (seed >> 2));
}

size_t HashWideString(const std::wstring& value) noexcept {
    return std::hash<std::wstring>{}(value);
}

struct TabViewItemKey {
    TabViewItemType type = TabViewItemType::kGroupHeader;
    uint64_t ordinal = 0;
    std::wstring savedGroupId;
    std::wstring path;
    std::wstring name;
};

struct TabViewItemKeyHash {
    size_t operator()(const TabViewItemKey& key) const noexcept {
        size_t hash = static_cast<size_t>(key.type);
        hash = HashCombine(hash, static_cast<size_t>(key.ordinal));
        hash = HashCombine(hash, HashWideString(key.savedGroupId));
        hash = HashCombine(hash, HashWideString(key.path));
        hash = HashCombine(hash, HashWideString(key.name));
        return hash;
    }
};

struct TabViewItemKeyEqual {
    bool operator()(const TabViewItemKey& a, const TabViewItemKey& b) const noexcept {
        return a.type == b.type && a.ordinal == b.ordinal && a.savedGroupId == b.savedGroupId &&
               a.path == b.path && a.name == b.name;
    }
};

TabViewItemKey MakeKey(const TabViewItem& item) {
    TabViewItemKey key;
    key.type = item.type;
    if (item.type == TabViewItemType::kTab) {
        key.ordinal = item.activationOrdinal != 0 ? item.activationOrdinal : item.lastActivatedTick;
        if (key.ordinal == 0 && item.pidl) {
            key.ordinal = reinterpret_cast<uint64_t>(item.pidl);
        }
        key.path = item.path;
        key.name = item.name;
    } else {
        key.ordinal = static_cast<uint64_t>(item.location.groupIndex);
        key.savedGroupId = item.savedGroupId;
        key.name = item.name;
    }
    return key;
}

RECT NormalizeRect(const RECT& rect) noexcept {
    RECT normalized = rect;
    if (normalized.left > normalized.right) {
        std::swap(normalized.left, normalized.right);
    }
    if (normalized.top > normalized.bottom) {
        std::swap(normalized.top, normalized.bottom);
    }
    return normalized;
}

bool ClipRectToClient(const RECT& rect, const RECT& client, RECT* clipped) noexcept {
    if (!clipped) {
        return false;
    }
    RECT normalized = NormalizeRect(rect);
    RECT intersection{};
    if (!IntersectRect(&intersection, &normalized, &client)) {
        return false;
    }
    if (IsRectEmpty(&intersection)) {
        return false;
    }
    *clipped = intersection;
    return true;
}

bool EquivalentTabViewItem(const TabViewItem& a, const TabViewItem& b) noexcept {
    return a.type == b.type && a.name == b.name && a.tooltip == b.tooltip &&
           a.selected == b.selected && a.collapsed == b.collapsed && a.totalTabs == b.totalTabs &&
           a.visibleTabs == b.visibleTabs && a.hiddenTabs == b.hiddenTabs &&
           a.hasCustomOutline == b.hasCustomOutline && a.outlineColor == b.outlineColor &&
           a.outlineStyle == b.outlineStyle && a.headerVisible == b.headerVisible &&
           a.isSavedGroup == b.isSavedGroup && a.progress == b.progress;
}

bool EquivalentVisualMetadata(const TabBandWindow::VisualItem& a,
                              const TabBandWindow::VisualItem& b) noexcept {
    if (a.firstInGroup != b.firstInGroup) {
        return false;
    }
    if (a.badgeWidth != b.badgeWidth) {
        return false;
    }
    if (a.hasGroupHeader != b.hasGroupHeader) {
        return false;
    }
    if (a.collapsedPlaceholder != b.collapsedPlaceholder) {
        return false;
    }
    if (a.indicatorHandle != b.indicatorHandle) {
        return false;
    }
    if (a.hasGroupHeader && !EquivalentTabViewItem(a.groupHeader, b.groupHeader)) {
        return false;
    }
    return true;
}

uint32_t DockModeToMask(TabBandDockMode mode) {
    return 1u << static_cast<uint32_t>(mode);
}

void UpdateAvailableDockMaskFromFrame(HWND frame) {
    if (!frame) {
        return;
    }

    uint32_t mask = 0;
    EnumChildWindows(
        frame,
        [](HWND hwnd, LPARAM param) -> BOOL {
            wchar_t className[64] = {};
            if (GetClassNameW(hwnd, className, ARRAYSIZE(className)) == 0) {
                return TRUE;
            }
            if (wcscmp(className, L"ReBarWindow32") != 0) {
                return TRUE;
            }

            const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
            const auto mode = DockModeFromRebarStyle(style);
            if (mode != TabBandDockMode::kAutomatic) {
                *reinterpret_cast<uint32_t*>(param) |= DockModeToMask(mode);
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&mask));

    if (mask != 0) {
        g_availableDockMask.store(mask, std::memory_order_release);
    }
}
// Older Windows SDKs used by consumers of the project might not expose the
// SID_SDataObject symbol (the service identifier for the current data object).
// Define the GUID locally so the build remains compatible with those SDKs.
constexpr GUID kSidDataObject = {0x000214e8, 0x0000, 0x0000,
                                 {0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

const wchar_t kWindowClassName[] = L"ShellTabsBandWindow";
constexpr size_t kInvalidIndex = std::numeric_limits<size_t>::max();
constexpr int kButtonWidth = 19;
constexpr int kButtonHeight = 19;
constexpr int kButtonMargin = 2;
constexpr int kItemMinWidth = 60;
constexpr int kGroupMinWidth = 90;
constexpr int kGroupGap = 4;   // gap between “islands” (groups)
constexpr int kTabGap   = 4;   // gap between adjacent tabs
constexpr int kPaddingX = 12;
constexpr int kGroupPaddingX = 16;
constexpr int kToolbarGripWidth = 14;
constexpr int kDragThreshold = 4;
constexpr int kTabCornerRadius = 8;
constexpr int kGroupCornerRadius = 10;
constexpr int kGroupOutlineThickness = 2;
constexpr int kIconGap = 6;
constexpr int kIslandIndicatorWidth = 5;
constexpr int kIslandOutlineThickness = 1;
constexpr int kCloseButtonSize = 14;
constexpr int kCloseButtonEdgePadding = 6;
constexpr int kCloseButtonSpacing = 6;
constexpr int kCloseButtonVerticalPadding = 3;
constexpr int kDropPreviewOffset = 12;
// Small placeholder for empty island content
constexpr int kEmptyIslandBodyMinWidth = 24; // enough space for a centered "+"
constexpr int kEmptyIslandBodyMaxWidth = 32; // clamp empty outline length
constexpr int kEmptyPlusSize = 14; // glyph size

const wchar_t kThemePreferenceKey[] =
L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
const wchar_t kThemePreferenceValue[] = L"AppsUseLightTheme";

constexpr UINT WM_SHELLTABS_EXTERNAL_DRAG = WM_APP + 60;
constexpr UINT WM_SHELLTABS_EXTERNAL_DRAG_LEAVE = WM_APP + 61;
constexpr UINT WM_SHELLTABS_EXTERNAL_DROP = WM_APP + 62;
const wchar_t kOverlayWindowClassName[] = L"ShellTabsDragOverlay";
constexpr UINT kPreviewHoverTime = 400;
constexpr ULONGLONG kProgressStaleTimeoutMs = 3000;

// How many rows of tabs max
static constexpr int kMaxTabRows = 5;

// Vertical gap between rows (keep it tight)
static constexpr int kRowGap = 2;

struct WindowRegistry {
    std::mutex mutex;
    std::unordered_map<HWND, TabBandWindow*> windows;
};

WindowRegistry& GetWindowRegistry() {
    static WindowRegistry registry;
    return registry;
}

void RegisterWindow(HWND hwnd, TabBandWindow* window) {
    if (!hwnd || !window) {
        return;
    }
    auto& registry = GetWindowRegistry();
    std::scoped_lock lock(registry.mutex);
    registry.windows[hwnd] = window;
}

void UnregisterWindow(HWND hwnd, TabBandWindow* window) {
    if (!hwnd) {
        return;
    }
    auto& registry = GetWindowRegistry();
    std::scoped_lock lock(registry.mutex);
    auto it = registry.windows.find(hwnd);
    if (it != registry.windows.end() && (!window || it->second == window)) {
        registry.windows.erase(it);
    }
}

TabBandWindow* LookupWindow(HWND hwnd) {
    if (!hwnd) {
        return nullptr;
    }
    auto& registry = GetWindowRegistry();
    std::scoped_lock lock(registry.mutex);
    auto it = registry.windows.find(hwnd);
    if (it == registry.windows.end()) {
        return nullptr;
    }
    return it->second;
}

TabBandWindow* FindWindowFromPoint(const POINT& screenPt) {
    HWND target = WindowFromPoint(screenPt);
    while (target) {
        if (auto* window = LookupWindow(target)) {
            return window;
        }
        target = GetParent(target);
    }
    return nullptr;
}

void DispatchExternalMessage(HWND hwnd, UINT message) {
    if (!hwnd) {
        return;
    }
    SendMessageTimeoutW(hwnd, message, 0, 0, SMTO_ABORTIFHUNG | SMTO_BLOCK, 50, nullptr);
}

struct TransferPayload {
    enum class Type {
        None,
        Tab,
        Group,
    } type = Type::None;
    TabBand* source = nullptr;
    TabBand* target = nullptr;
    bool select = false;
    int targetGroupIndex = -1;
    int targetTabIndex = -1;
    bool createGroup = false;
    bool headerVisible = true;
    TabInfo tab;
    TabGroup group;
};

struct SharedDragState {
    std::mutex mutex;
    TabBandWindow* source = nullptr;
    TabBandWindow* hover = nullptr;
    POINT screen{};
    TabBandWindow::HitInfo origin;
    bool targetValid = false;
    TabBandWindow::DropTarget target;
    std::unique_ptr<TransferPayload> payload;
};

SharedDragState& GetSharedDragState() {
    static SharedDragState state;
    return state;
}

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

COLORREF GetGroupColor(bool selected) {
    return selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_BTNFACE);
}

COLORREF GetTabColor(bool selected) {
    return selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_WINDOW);
}

COLORREF GetTabTextColor(bool selected) {
    return selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_WINDOWTEXT);
}

COLORREF LightenColor(COLORREF color, double factor) {
    factor = std::clamp(factor, 0.0, 1.0);
    const int r = static_cast<int>(GetRValue(color) + (255 - GetRValue(color)) * factor);
    const int g = static_cast<int>(GetGValue(color) + (255 - GetGValue(color)) * factor);
    const int b = static_cast<int>(GetBValue(color) + (255 - GetBValue(color)) * factor);
    return RGB(std::clamp(r, 0, 255), std::clamp(g, 0, 255), std::clamp(b, 0, 255));
}

COLORREF DarkenColor(COLORREF color, double factor) {
    factor = std::clamp(factor, 0.0, 1.0);
    const int r = static_cast<int>(GetRValue(color) * (1.0 - factor));
    const int g = static_cast<int>(GetGValue(color) * (1.0 - factor));
    const int b = static_cast<int>(GetBValue(color) * (1.0 - factor));
    return RGB(std::clamp(r, 0, 255), std::clamp(g, 0, 255), std::clamp(b, 0, 255));
}

COLORREF BlendColors(COLORREF base, COLORREF accent, double ratio) {
    ratio = std::clamp(ratio, 0.0, 1.0);
    const double inverse = 1.0 - ratio;
    const int r = static_cast<int>(GetRValue(base) * inverse + GetRValue(accent) * ratio);
    const int g = static_cast<int>(GetGValue(base) * inverse + GetGValue(accent) * ratio);
    const int b = static_cast<int>(GetBValue(base) * inverse + GetBValue(accent) * ratio);
    return RGB(std::clamp(r, 0, 255), std::clamp(g, 0, 255), std::clamp(b, 0, 255));
}

double ComputeLuminance(COLORREF color) {
    const double r = GetRValue(color) / 255.0;
    const double g = GetGValue(color) / 255.0;
    const double b = GetBValue(color) / 255.0;
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

COLORREF AdjustForDarkTone(COLORREF color, double baseFactor, bool darkMode) {
    if (!darkMode) {
        return color;
    }
    double factor = std::clamp(baseFactor, 0.0, 1.0);
    const double luminance = ComputeLuminance(color);
    if (luminance > 0.3) {
        factor = std::clamp(factor + (luminance - 0.3) * 1.1, factor, 0.8);
    }
    return BlendColors(color, RGB(0, 0, 0), factor);
}

COLORREF ResolveIndicatorColor(const TabViewItem* header, const TabViewItem& tab) {
    if (header) {
        if (header->hasCustomOutline) {
            return header->outlineColor;
        }
    }
    if (tab.hasCustomOutline) {
        return tab.outlineColor;
    }
    return GetSysColor(COLOR_HOTLIGHT);
}

HFONT GetDefaultFont() {
    return static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
}

void ApplyImmersiveDarkMode(HWND hwnd, bool enabled) {
    if (!hwnd) {
        return;
    }

    using DwmSetWindowAttributeFn = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFn setWindowAttribute = []() -> DwmSetWindowAttributeFn {
        HMODULE module = GetModuleHandleW(L"dwmapi.dll");
        if (!module) {
            module = LoadLibraryW(L"dwmapi.dll");
        }
        if (!module) {
            return nullptr;
        }
        return reinterpret_cast<DwmSetWindowAttributeFn>(
            GetProcAddress(module, "DwmSetWindowAttribute"));
    }();

    if (!setWindowAttribute) {
        return;
    }

    const BOOL value = enabled ? TRUE : FALSE;
    setWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &value, sizeof(value));
}

HWND CreateDragOverlayWindow() {
    static bool registered = false;
    if (!registered) {
        WNDCLASSW wc{};
        wc.lpfnWndProc = DefWindowProcW;
        wc.hInstance = GetModuleHandleInstance();
        wc.lpszClassName = kOverlayWindowClassName;
        wc.hCursor = nullptr;
        if (!RegisterClassW(&wc)) {
            if (GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
                return nullptr;
            }
        }
        registered = true;
    }

    return CreateWindowExW(WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT | WS_EX_TOPMOST,
                           kOverlayWindowClassName, L"", WS_POPUP, 0, 0, 0, 0, nullptr, nullptr,
                           GetModuleHandleInstance(), nullptr);
}

}  // namespace

class TabBandWindow::BandDropTarget : public IDropTarget {
public:
    explicit BandDropTarget(TabBandWindow* owner) : m_refCount(1), m_owner(owner) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDropTarget) {
            *ppv = static_cast<IDropTarget*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return ++m_refCount; }

    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --m_refCount;
        if (count == 0) {
            delete this;
        }
        return count;
    }

    IFACEMETHODIMP DragEnter(IDataObject* dataObject, DWORD keyState, POINTL point, DWORD* effect) override {
        if (!m_owner) {
            if (effect) {
                *effect = DROPEFFECT_NONE;
            }
            return E_FAIL;
        }
        return m_owner->OnNativeDragEnter(dataObject, keyState, point, effect);
    }

    IFACEMETHODIMP DragOver(DWORD keyState, POINTL point, DWORD* effect) override {
        if (!m_owner) {
            if (effect) {
                *effect = DROPEFFECT_NONE;
            }
            return E_FAIL;
        }
        return m_owner->OnNativeDragOver(keyState, point, effect);
    }

    IFACEMETHODIMP DragLeave() override {
        if (!m_owner) {
            return E_FAIL;
        }
        return m_owner->OnNativeDragLeave();
    }

    IFACEMETHODIMP Drop(IDataObject* dataObject, DWORD keyState, POINTL point, DWORD* effect) override {
        if (!m_owner) {
            if (effect) {
                *effect = DROPEFFECT_NONE;
            }
            return E_FAIL;
        }
        return m_owner->OnNativeDrop(dataObject, keyState, point, effect);
    }

private:
    std::atomic<ULONG> m_refCount;
    TabBandWindow* m_owner = nullptr;
};

TabBandWindow::TabBandWindow(TabBand* owner) : m_owner(owner) {
    ResetThemePalette();
    m_dropHoverHit = {};
    m_dropHoverHasFileData = false;
    m_dropHoverTimerActive = false;
}

TabBandWindow::~TabBandWindow() { Destroy(); }

void TabBandWindow::SetPreferredDockMode(TabBandDockMode mode) {
    m_preferredDockMode = mode;
    EnsureRebarIntegration();
}

uint32_t TabBandWindow::GetAvailableDockMask() {
    uint32_t mask = g_availableDockMask.load(std::memory_order_acquire);
    if (mask == 0) {
        mask |= DockModeToMask(TabBandDockMode::kTop);
        mask |= DockModeToMask(TabBandDockMode::kBottom);
    }
    return mask;
}

HWND TabBandWindow::Create(HWND parent) {
    if (m_hwnd) {
        return m_hwnd;
    }

    WNDCLASSW wc{};
    wc.style = 0;
    wc.lpfnWndProc = TabBandWindow::WndProc;
    wc.hInstance = GetModuleHandleInstance();
    wc.lpszClassName = kWindowClassName;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

    static ATOM atom = RegisterClassW(&wc);
    if (!atom && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        return nullptr;
    }

    m_hwnd = CreateWindowExW(0, kWindowClassName, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP,
                             0, 0, 0, 0, parent, nullptr, GetModuleHandleInstance(), this);

    if (m_hwnd) {
        RegisterWindow(m_hwnd, this);
        EnsureRebarIntegration();
        if (!m_dropTarget) {
            m_dropTarget.Attach(new BandDropTarget(this));
        }
        if (m_dropTarget) {
            const HRESULT hr = RegisterDragDrop(m_hwnd, m_dropTarget.Get());
            if (FAILED(hr)) {
                m_dropTarget.Reset();
            }
        }
        RegisterShellNotifications();
        TabManager::Get().RegisterProgressListener(m_hwnd);
        UpdateProgressAnimationState();
    }

    return m_hwnd;
}

void TabBandWindow::Destroy() {
    CancelDrag();
    ClearExplorerContext();
    ClearVisualItems();
    CloseThemeHandles();
    ClearDropHoverState();
    HidePreviewWindow(true);
    if (m_hwnd) {
        TabManager::Get().UnregisterProgressListener(m_hwnd);
    }
    UnregisterShellNotifications();
    if (m_hwnd && m_progressTimerActive) {
        KillTimer(m_hwnd, kProgressTimerId);
        m_progressTimerActive = false;
    }
    if (m_hwnd && m_dropTarget) {
        RevokeDragDrop(m_hwnd);
    }
    m_dropTarget.Reset();
    m_darkMode = false;
    m_refreshingTheme = false;
    m_windowDarkModeInitialized = false;
    m_windowDarkModeValue = false;
    m_buttonDarkModeInitialized = false;
    m_buttonDarkModeValue = false;
    ResetThemePalette();

    if (m_newTabButton) {
        DestroyWindow(m_newTabButton);
        m_newTabButton = nullptr;
    }
    if (m_hwnd) {
        UnregisterWindow(m_hwnd, this);
        DestroyWindow(m_hwnd);
        m_hwnd = nullptr;
    }
    m_parentRebar = nullptr;
    m_rebarBandIndex = -1;
    m_tabData.clear();
    m_nextRedrawIncremental = false;
    m_redrawMetrics = {};
    m_lastAppliedRowCount = 0;
}

void TabBandWindow::Show(bool show) {
    if (!m_hwnd) {
        return;
    }
    ShowWindow(m_hwnd, show ? SW_SHOW : SW_HIDE);
}

void TabBandWindow::SetTabs(const std::vector<TabViewItem>& items) {
    m_tabData = items;
    m_contextHit = {};
    ClearExplorerContext();

    if (!m_hwnd) {
        DestroyVisualItemResources(m_items);
        m_items.clear();
        m_emptyIslandPlusButtons.clear();
        m_nextRedrawIncremental = false;
        m_lastAppliedRowCount = 0;
        return;
    }

    std::vector<VisualItem> oldItems;
    oldItems.swap(m_items);

    HideDragOverlay(true);
    HidePreviewWindow(false);
    m_drag = {};
    m_emptyIslandPlusButtons.clear();

    LayoutResult layout = BuildLayoutItems(items);

    LayoutDiffStats diff = ComputeLayoutDiff(oldItems, layout.items);
    const int normalizedRowCount = layout.rowCount > 0 ? layout.rowCount : std::max(m_lastRowCount, 1);
    const bool rowCountChanged = normalizedRowCount != m_lastRowCount;

    m_items = std::move(layout.items);
    m_lastRowCount = normalizedRowCount;

    bool topologyChanged = diff.inserted > 0 || diff.removed > 0 || rowCountChanged;
    if (topologyChanged) {
        InvalidateRect(m_hwnd, nullptr, FALSE);
        m_nextRedrawIncremental = false;
    } else if (!diff.invalidRects.empty()) {
        for (const auto& rect : diff.invalidRects) {
            InvalidateRect(m_hwnd, &rect, FALSE);
        }
        m_nextRedrawIncremental = true;
    } else {
        m_nextRedrawIncremental = false;
    }

    if (rowCountChanged) {
        AdjustBandHeightToRow();
    }

    UpdateProgressAnimationState();

    DestroyVisualItemResources(oldItems);

    if (diff.inserted > 0 || diff.removed > 0 || diff.moved > 0 || diff.updated > 0) {
        LogMessage(LogLevel::Info,
                   L"Tab diff: +%zu -%zu move=%zu update=%zu rows=%d incremental=%ls",
                   diff.inserted, diff.removed, diff.moved, diff.updated, m_lastRowCount,
                   m_nextRedrawIncremental ? L"true" : L"false");
    }
}

bool TabBandWindow::HasFocus() const {
    if (!m_hwnd) {
        return false;
    }
    const HWND focus = GetFocus();
    return focus == m_hwnd;
}

void TabBandWindow::FocusTab() {
    if (m_hwnd) {
        SetFocus(m_hwnd);
    }
}
STDMETHODIMP TabBandWindow::SetSite(IUnknown* pUnkSite) {
        if (!pUnkSite) {
                m_siteSp.Reset();
                return S_OK;
        }

        Microsoft::WRL::ComPtr<IServiceProvider> sp;
        if (FAILED(pUnkSite->QueryInterface(IID_PPV_ARGS(&sp)))) {
                return E_NOINTERFACE;
        }

        m_siteSp = sp;

        return S_OK;
}
STDMETHODIMP TabBandWindow::GetSite(REFIID riid, void** ppvSite) {
	if (!ppvSite) return E_POINTER;
	*ppvSite = nullptr;

	if (!m_siteSp) return E_FAIL;

	// Return the requested interface from the stored site.
	return m_siteSp->QueryInterface(riid, ppvSite);
}


void TabBandWindow::Layout(int width, int height) {
    m_clientRect = {0, 0, width, height};

    int buttonHeight = std::max(0, height - kButtonMargin * 2);
    if (buttonHeight > kButtonHeight) {
        buttonHeight = kButtonHeight;
    }
    if (buttonHeight == 0 && height > 0) {
        buttonHeight = std::min(height, kButtonHeight);
    }

    int buttonWidth = 0;
    if (buttonHeight > 0) {
        buttonWidth = std::min(kButtonWidth, buttonHeight);
    }
    const int maxAvailableWidth = std::max(0, width - kButtonMargin);
    if (buttonWidth == 0 && maxAvailableWidth > 0) {
        buttonWidth = std::min(kButtonWidth, maxAvailableWidth);
    }
    if (buttonWidth == 0 && width > 0) {
        buttonWidth = std::min(kButtonWidth, width);
    }
    const int buttonX = std::max(0, width - buttonWidth - kButtonMargin);
    const int buttonY = std::max(0, (height - buttonHeight) / 2);
    if (m_newTabButton) {
        MoveWindow(m_newTabButton, buttonX, buttonY, buttonWidth, buttonHeight, TRUE);
    }

    m_clientRect.right = std::max(0, buttonX - kButtonMargin);
    RebuildLayout();
}

void TabBandWindow::DestroyVisualItemResources(std::vector<VisualItem>& items) {
    for (auto& item : items) {
        if (item.icon) {
            DestroyIcon(item.icon);
            item.icon = nullptr;
        }
    }
}

void TabBandWindow::ClearVisualItems() {
    HideDragOverlay(true);
    HidePreviewWindow(false);

    DestroyVisualItemResources(m_items);

    m_items.clear();
    m_drag = {};
    m_contextHit = {};
    m_emptyIslandPlusButtons.clear();
}

TabBandWindow::LayoutResult TabBandWindow::BuildLayoutItems(const std::vector<TabViewItem>& items) {
    LayoutResult result;
    if (!m_hwnd) {
        return result;
    }

    m_emptyIslandPlusButtons.clear();

    RECT bounds = m_clientRect;
    const int availableWidth = bounds.right - bounds.left;
    if (availableWidth <= 0) {
        return result;
    }

    HDC dc = GetDC(m_hwnd);
    if (!dc) {
        return result;
    }
    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(dc, font));

        const int baseIconWidth = std::max(GetSystemMetrics(SM_CXSMICON), 16);
        const int baseIconHeight = std::max(GetSystemMetrics(SM_CYSMICON), 16);

        TEXTMETRIC tm{};
        GetTextMetrics(dc, &tm);

        int rowHeight = static_cast<int>(tm.tmHeight);
        if (rowHeight > 0) {
                rowHeight += 6;  // give text breathing room
        } else {
                rowHeight = baseIconHeight + 8;
        }
        rowHeight = std::max(rowHeight, baseIconHeight + 8);
        rowHeight = std::max(rowHeight, kCloseButtonSize + kCloseButtonVerticalPadding * 2 + 4);
        rowHeight = std::max(rowHeight, kButtonHeight - kButtonMargin);
        rowHeight = std::max(rowHeight, 24);

        const int bandWidth = bounds.right - bounds.left;
        const int gripWidth = std::clamp(m_toolbarGripWidth, 0, std::max(0, bandWidth));
        int x = bounds.left + gripWidth - 3;   // DO NOT TOUCH

        // row layout
        const int startY = bounds.top + 2;
        const int maxX = bounds.right;

        int row = 0;
        int maxRowUsed = 0;
        auto rowTop = [&](int r) { return startY + r * (rowHeight + kRowGap); };
        auto rowBottom = [&](int r) { return rowTop(r) + rowHeight; };

        auto try_wrap = [&]() {
                if (row + 1 < kMaxTabRows) {
                        ++row;
                        if (row > maxRowUsed) {
                                maxRowUsed = row;
                        }
                        x = bounds.left + gripWidth - 3;  // DO NOT TOUCH
                        return true;
                }
                return false;
        };

	int currentGroup = -1;
	TabViewItem currentHeader{};
	bool headerMetadata = false;
	bool expectFirstTab = false;
	bool pendingIndicator = false;
	TabViewItem indicatorHeader{};

        result.items.reserve(items.size() + 8);

        for (const auto& item : items) {
		if (item.type == TabViewItemType::kGroupHeader) {
			pendingIndicator = false;
			currentGroup = item.location.groupIndex;
			currentHeader = item;
			headerMetadata = true;
			expectFirstTab = true;

			const bool collapsed = item.collapsed;
			const bool hasVisibleTabs = item.visibleTabs > 0;
			if (!item.headerVisible && !collapsed && hasVisibleTabs) {
				indicatorHeader = item;
				pendingIndicator = true;
				continue;
			}

			if (currentGroup >= 0 && x > bounds.left) {
				x += kGroupGap;
			}

			int width = kIslandIndicatorWidth;
			if (x + width > maxX) {
				if (!try_wrap()) break;
			}

			// Emit the island's indicator handle
			VisualItem visual;
			visual.data = item;
			visual.firstInGroup = true;
			visual.collapsedPlaceholder = collapsed;
			visual.indicatorHandle = true;
                        visual.bounds = { x, rowTop(row), x + width, rowBottom(row) };
                        visual.row = row;
                        result.items.emplace_back(std::move(visual));
                        x += width;

                        // NEW: empty island => reserve a tiny body and register a '+' target
                        if (item.headerVisible && !collapsed && !hasVisibleTabs) {
                                const int remaining = maxX - x;
                                if (remaining > 0) {
                                        const int placeholderWidth = std::min(remaining, kEmptyIslandBodyMaxWidth);
                                        if (placeholderWidth > 0) {
                                                RECT placeholder{ x, rowTop(row), x + placeholderWidth, rowBottom(row) };

                                                // Emit a synthetic body so the island outline has a region to hug
                                                VisualItem emptyBody;
                                                emptyBody.data = item;               // tie to this group header
                                                emptyBody.hasGroupHeader = true;
                                                emptyBody.groupHeader = currentHeader;
                                                emptyBody.bounds = placeholder;
                                                emptyBody.row = row;
                                                result.items.emplace_back(std::move(emptyBody));

                                                const int bodyWidth = placeholder.right - placeholder.left;
                                                if (bodyWidth >= 4) {
                                                        const int h = placeholder.bottom - placeholder.top;
                                                        const int maxCentered = std::max(bodyWidth - 4, 0);
                                                        int size = std::min(kEmptyPlusSize, maxCentered);
                                                        if (size < 8) {
                                                                size = std::max(4, maxCentered);
                                                        }
                                                        const int plusLeft = placeholder.left + (bodyWidth - size) / 2;
                                                        RECT plus{
                                                                plusLeft,
                                                                placeholder.top + (h - size) / 2,
                                                                plusLeft + size,
                                                                placeholder.top + (h - size) / 2 + size
                                                        };
                                                        m_emptyIslandPlusButtons.push_back({ currentGroup, plus });
                                                }

                                                x = placeholder.right;
                                        }
                                }
                        }
			continue;
		}

		VisualItem visual;
		visual.data = item;

		if (currentGroup != item.location.groupIndex) {
			currentGroup = item.location.groupIndex;
			headerMetadata = false;
			expectFirstTab = true;
			if (!result.items.empty()) {
				x += kGroupGap;
			}
			pendingIndicator = false;
		}
		else if (!expectFirstTab) {
			x += kTabGap;
		}

		if (expectFirstTab) {
			visual.firstInGroup = true;
			expectFirstTab = false;
		}
		visual.hasGroupHeader = headerMetadata;
		if (visual.hasGroupHeader) {
			visual.groupHeader = currentHeader;
		}
                if (pendingIndicator && visual.firstInGroup) {
                        visual.hasGroupHeader = true;
                        visual.groupHeader = indicatorHeader;
                        visual.indicatorHandle = indicatorHeader.headerVisible;
                        pendingIndicator = false;
                        headerMetadata = true;
                }

		SIZE textSize{ 0, 0 };
		if (!item.name.empty()) {
			GetTextExtentPoint32W(dc, item.name.c_str(),
				static_cast<int>(item.name.size()), &textSize);
		}

		int width = textSize.cx + kPaddingX * 2;
		width = std::max(width, kItemMinWidth);

                visual.badgeWidth = 0;

                visual.icon = LoadItemIcon(item, SHGFI_SMALLICON);
		if (visual.icon) {
			visual.iconWidth = baseIconWidth;
			visual.iconHeight = baseIconHeight;
			ICONINFO iconInfo{};
			if (GetIconInfo(visual.icon, &iconInfo)) {
				BITMAP bitmap{};
				if (iconInfo.hbmColor &&
					GetObject(iconInfo.hbmColor, sizeof(bitmap), &bitmap) == sizeof(bitmap)) {
					visual.iconWidth = bitmap.bmWidth;
					visual.iconHeight = bitmap.bmHeight;
				}
				else if (iconInfo.hbmMask &&
					GetObject(iconInfo.hbmMask, sizeof(bitmap), &bitmap) == sizeof(bitmap)) {
					visual.iconWidth = bitmap.bmWidth;
					visual.iconHeight = bitmap.bmHeight / 2;
				}
				if (iconInfo.hbmColor) DeleteObject(iconInfo.hbmColor);
				if (iconInfo.hbmMask)  DeleteObject(iconInfo.hbmMask);
			}
			if (visual.iconWidth <= 0) visual.iconWidth = baseIconWidth;
			if (visual.iconHeight <= 0) visual.iconHeight = baseIconHeight;
			width += visual.iconWidth + kIconGap;
		}

		width += kCloseButtonSize + kCloseButtonEdgePadding + kCloseButtonSpacing;

                bool wrapped = false;
                if (x + width > maxX) {
                        if (!try_wrap()) {
                                width = std::max(40, maxX - x);
                                if (width <= 0) break;
                        } else {
                                wrapped = true;
                        }
                }

                if (wrapped && visual.firstInGroup) {
                        if (!result.items.empty()) {
                                VisualItem& previous = result.items.back();
                                if (previous.indicatorHandle &&
                                        previous.data.location.groupIndex == item.location.groupIndex) {
                                        const int indicatorWidth = previous.bounds.right - previous.bounds.left;
                                        previous.bounds.left = x;
                                        previous.bounds.right = x + indicatorWidth;
                                        previous.bounds.top = rowTop(row);
                                        previous.bounds.bottom = rowBottom(row);
                                        previous.row = row;
                                        x += indicatorWidth;
                                }
                        }
                }

                width = std::clamp(width, 40, maxX - x);

                visual.bounds = { x, rowTop(row), x + width, rowBottom(row) };
                visual.row = row;
                visual.index = result.items.size();
                result.items.emplace_back(std::move(visual));
                x += width;
        }

        if (row > maxRowUsed) {
                maxRowUsed = row;
        }
        result.rowCount = std::clamp(maxRowUsed + 1, 1, kMaxTabRows);

        if (oldFont) SelectObject(dc, oldFont);
        ReleaseDC(m_hwnd, dc);
    return result;
}

void TabBandWindow::RebuildLayout() {
    if (!m_hwnd) {
        DestroyVisualItemResources(m_items);
        m_items.clear();
        m_emptyIslandPlusButtons.clear();
        m_nextRedrawIncremental = false;
        return;
    }

    std::vector<VisualItem> oldItems;
    oldItems.swap(m_items);

    HideDragOverlay(true);
    HidePreviewWindow(false);
    m_drag = {};
    m_contextHit = {};
    m_emptyIslandPlusButtons.clear();

    LayoutResult layout = BuildLayoutItems(m_tabData);
    m_items = std::move(layout.items);

    const int normalizedRowCount = layout.rowCount > 0 ? layout.rowCount : std::max(m_lastRowCount, 1);
    const bool rowChanged = normalizedRowCount != m_lastRowCount;
    m_lastRowCount = normalizedRowCount;

    DestroyVisualItemResources(oldItems);

    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }

    m_nextRedrawIncremental = false;

    if (rowChanged) {
        AdjustBandHeightToRow();
    }
}

TabBandWindow::LayoutDiffStats TabBandWindow::ComputeLayoutDiff(
    const std::vector<VisualItem>& oldItems, const std::vector<VisualItem>& newItems) const {
    LayoutDiffStats stats;

    if (!m_hwnd) {
        stats.inserted = newItems.size();
        stats.removed = oldItems.size();
        return stats;
    }

    std::unordered_map<TabViewItemKey, std::vector<size_t>, TabViewItemKeyHash, TabViewItemKeyEqual> oldMap;
    oldMap.reserve(oldItems.size());
    for (size_t i = 0; i < oldItems.size(); ++i) {
        oldMap[MakeKey(oldItems[i].data)].push_back(i);
    }

    std::vector<bool> consumed(oldItems.size(), false);

    RECT client = m_clientRect;
    if (client.right <= client.left || client.bottom <= client.top) {
        if (m_hwnd) {
            GetClientRect(m_hwnd, &client);
        }
    }

    auto enqueueRect = [&](const RECT& rect) {
        RECT clipped{};
        if (ClipRectToClient(rect, client, &clipped)) {
            stats.invalidRects.push_back(clipped);
        }
    };

    for (const auto& item : newItems) {
        const auto key = MakeKey(item.data);
        auto it = oldMap.find(key);
        if (it == oldMap.end() || it->second.empty()) {
            ++stats.inserted;
            enqueueRect(item.bounds);
            continue;
        }

        const size_t oldIndex = it->second.back();
        it->second.pop_back();
        consumed[oldIndex] = true;
        const VisualItem& oldItem = oldItems[oldIndex];

        const bool moved = !EqualRect(&oldItem.bounds, &item.bounds);
        const bool contentChanged = !EquivalentTabViewItem(oldItem.data, item.data) ||
                                    !EquivalentVisualMetadata(oldItem, item);

        if (moved) {
            ++stats.moved;
        }
        if (contentChanged) {
            ++stats.updated;
        }
        if (moved || contentChanged) {
            RECT unionRect{};
            RECT oldRect = NormalizeRect(oldItem.bounds);
            RECT newRect = NormalizeRect(item.bounds);
            UnionRect(&unionRect, &oldRect, &newRect);
            enqueueRect(unionRect);
        }
    }

    for (size_t i = 0; i < consumed.size(); ++i) {
        if (!consumed[i]) {
            ++stats.removed;
            enqueueRect(oldItems[i].bounds);
        }
    }

    return stats;
}



static bool QueryUserDarkMode() {
	// Win10/11: HKCU\...\Personalize\AppsUseLightTheme (0=dark, 1=light)
	// Win7/8: key/value not present -> treat as light (false).
	DWORD v = 1;
	DWORD sz = sizeof(v);
	const wchar_t* k = L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
	const wchar_t* n = L"AppsUseLightTheme";
	LSTATUS st = RegGetValueW(HKEY_CURRENT_USER, k, n, RRF_RT_DWORD, nullptr, &v, &sz);
	if (st == ERROR_SUCCESS) return v == 0;
	return false; // legacy OS: no dark mode
}

bool TabBandWindow::BandHasRebarGrip() const {
	if (!m_parentRebar || m_rebarBandIndex < 0) return false;
	REBARBANDINFOW bi{ sizeof(bi) };
	bi.fMask = RBBIM_STYLE;
	if (!SendMessageW(m_parentRebar, RB_GETBANDINFO, m_rebarBandIndex, reinterpret_cast<LPARAM>(&bi)))
		return false;
	const bool noGrip = (bi.fStyle & RBBS_NOGRIPPER) != 0;
	const bool always = (bi.fStyle & RBBS_GRIPPERALWAYS) != 0;
	return !noGrip && always;
}

void TabBandWindow::UpdateRebarColors() {
	if (!m_parentRebar || !IsWindow(m_parentRebar)) return;

	// Revert to defaults so bands are transparent to the bar we paint.
	// Do NOT set RB_SETBKCOLOR or a custom COLORSCHEME in dark;
	// themed rebars ignore parts of it and it triggers extra invalidation.
	const int count = (int)SendMessageW(m_parentRebar, RB_GETBANDCOUNT, 0, 0);
	for (int i = 0; i < count; ++i) {
		REBARBANDINFOW bi{ sizeof(bi) };
		bi.fMask = RBBIM_COLORS;
		bi.clrBack = CLR_DEFAULT; // transparent to bar
		bi.clrFore = CLR_DEFAULT;
		SendMessageW(m_parentRebar, RB_SETBANDINFO, i, (LPARAM)&bi);
	}

	RedrawWindow(m_parentRebar, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
}


void TabBandWindow::DrawBackground(HDC dc, const RECT& bounds) const {
	if (!dc) return;
	if (bounds.right <= bounds.left || bounds.bottom <= bounds.top) return;

	const_cast<TabBandWindow*>(this)->EnsureRebarIntegration();

	bool backgroundDrawn = false;

	// Only let the parent paint if we're NOT in dark mode.
	if (!m_darkMode && m_hwnd) {
		RECT clip = bounds;
		const int saved = SaveDC(dc);
		if (saved != 0) {
			HWND parent = GetParent(m_hwnd);
			if (parent) {
				POINT origin{ 0, 0 };
				MapWindowPoints(m_hwnd, parent, &origin, 1);
				SetWindowOrgEx(dc, origin.x, origin.y, nullptr);
			}
			else {
				POINT screen{ 0, 0 };
				ClientToScreen(m_hwnd, &screen);
				SetWindowOrgEx(dc, screen.x, screen.y, nullptr);
			}
			HRESULT hr = DrawThemeParentBackgroundEx(m_hwnd, dc, DTBG_CLIPRECT, &clip);
			RestoreDC(dc, saved);
			if (SUCCEEDED(hr)) backgroundDrawn = true;
		}
		else if (SUCCEEDED(DrawThemeParentBackgroundEx(m_hwnd, dc, DTBG_CLIPRECT, &clip))) {
			backgroundDrawn = true;
		}
	}

	// After: never draw the themed rebar surfaces when dark.
	if (!backgroundDrawn && m_rebarTheme && !m_darkMode) {
		RECT fillRect = bounds;
		if (SUCCEEDED(DrawThemeBackground(m_rebarTheme, dc, RP_BACKGROUND, 0, &fillRect, nullptr))) {
			backgroundDrawn = true;
		}
	}
	if (!backgroundDrawn && m_rebarTheme && !m_darkMode) {
		RECT fillRect = bounds;
		if (SUCCEEDED(DrawThemeBackground(m_rebarTheme, dc, RP_BAND, 0, &fillRect, nullptr))) {
			backgroundDrawn = true;
		}
	}

	// Fallback fill (your code) now actually runs in dark mode:
	if (!backgroundDrawn) {
		const COLORREF fallback = m_themePalette.rebarBackground;
		HBRUSH b = CreateSolidBrush(fallback);
		if (b) { FillRect(dc, &bounds, b); DeleteObject(b); }
		else { FillRect(dc, &bounds, GetSysColorBrush(COLOR_BTNFACE)); }
		backgroundDrawn = true;
	}

	const int bandWidth = static_cast<int>(bounds.right - bounds.left);
	const int gripWidth = std::clamp(m_toolbarGripWidth, 0, std::max(0, bandWidth));
	if (m_rebarTheme && gripWidth > 0 && !BandHasRebarGrip()) {
		RECT gripRect{ bounds.left, bounds.top, bounds.left + gripWidth, bounds.bottom };
		if (gripRect.right > gripRect.left) {
			if (FAILED(DrawThemeBackground(m_rebarTheme, dc, RP_GRIPPER, 0, &gripRect, nullptr))) {
				DrawThemeBackground(m_rebarTheme, dc, RP_GRIPPERVERT, 0, &gripRect, nullptr);
			}
		}
	}
}


void TabBandWindow::Draw(HDC dc) const {
    if (!dc) {
        return;
    }

    RECT windowRect = m_clientRect;
    if (m_hwnd) {
        GetClientRect(m_hwnd, &windowRect);
    }

    const int width = windowRect.right - windowRect.left;
    const int height = windowRect.bottom - windowRect.top;
    if (width <= 0 || height <= 0) {
        return;
    }

    const bool incremental = m_nextRedrawIncremental;
    struct DrawMetricsGuard {
        TabBandWindow* owner;
        bool incremental;
        std::chrono::steady_clock::time_point start;
        DrawMetricsGuard(TabBandWindow* o, bool inc)
            : owner(o), incremental(inc), start(std::chrono::steady_clock::now()) {}
        ~DrawMetricsGuard() {
            if (!owner) {
                return;
            }
            const auto end = std::chrono::steady_clock::now();
            const double ms = std::chrono::duration<double, std::milli>(end - start).count();
            owner->RecordRedrawDuration(ms, incremental);
            owner->m_nextRedrawIncremental = false;
        }
    } guard(const_cast<TabBandWindow*>(this), incremental);

    HDC memDC = CreateCompatibleDC(dc);
    if (!memDC) {
        PaintSurface(dc, windowRect);
        return;
    }

    HBITMAP buffer = CreateCompatibleBitmap(dc, width, height);
    if (!buffer) {
        DeleteDC(memDC);
        PaintSurface(dc, windowRect);
        return;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, buffer);
    RECT localRect{0, 0, width, height};
    PaintSurface(memDC, localRect);
    BitBlt(dc, windowRect.left, windowRect.top, width, height, memDC, 0, 0, SRCCOPY);
    SelectObject(memDC, oldBitmap);
    DeleteObject(buffer);
    DeleteDC(memDC);
}

void TabBandWindow::RecordRedrawDuration(double milliseconds, bool incremental) {
    m_redrawMetrics.lastDurationMs = milliseconds;
    m_redrawMetrics.lastWasIncremental = incremental;

    if (incremental) {
        m_redrawMetrics.incrementalTotalMs += milliseconds;
        ++m_redrawMetrics.incrementalCount;
    } else {
        m_redrawMetrics.fullTotalMs += milliseconds;
        ++m_redrawMetrics.fullCount;
    }

    const uint64_t totalSamples = m_redrawMetrics.incrementalCount + m_redrawMetrics.fullCount;
    if (totalSamples > 0 && (totalSamples % 60) == 0) {
        const double incrementalAvg = m_redrawMetrics.incrementalCount > 0
                                          ? m_redrawMetrics.incrementalTotalMs /
                                                static_cast<double>(m_redrawMetrics.incrementalCount)
                                          : 0.0;
        const double fullAvg = m_redrawMetrics.fullCount > 0
                                   ? m_redrawMetrics.fullTotalMs /
                                         static_cast<double>(m_redrawMetrics.fullCount)
                                   : 0.0;

        LogMessage(LogLevel::Info,
                   L"Tab redraw metrics: incremental %.2f ms (%llu), full %.2f ms (%llu), last %.2f ms (%ls)",
                   incrementalAvg,
                   static_cast<unsigned long long>(m_redrawMetrics.incrementalCount), fullAvg,
                   static_cast<unsigned long long>(m_redrawMetrics.fullCount), milliseconds,
                   incremental ? L"incremental" : L"full");
    }
}

void TabBandWindow::DrawEmptyIslandPluses(HDC dc) const {
	if (!dc) return;

	HPEN pen = CreatePen(PS_SOLID, 2,
		m_themePalette.tabTextValid ? m_themePalette.tabText : RGB(220, 220, 220));
	if (!pen) return;

	HGDIOBJ old = SelectObject(dc, pen);

	for (const auto& b : m_emptyIslandPlusButtons) {
		// force integer math, no template weirdness, no Windows macros
		const int w = static_cast<int>(b.rect.right - b.rect.left);
		const int h = static_cast<int>(b.rect.bottom - b.rect.top);
		const int cx = b.rect.left + w / 2;
		const int cy = b.rect.top + h / 2;

		// radius = min(w, h)/2 - 1, clamped to >= 1, done without std::min/std::max
		int d = (w < h) ? w : h;
		int r = (d / 2) - 1;
		if (r < 1) r = 1;

		// horizontal
		MoveToEx(dc, cx - r, cy, nullptr);
		LineTo(dc, cx + r, cy);
		// vertical
		MoveToEx(dc, cx, cy - r, nullptr);
		LineTo(dc, cx, cy + r);
	}

	SelectObject(dc, old);
	DeleteObject(pen);
}


void TabBandWindow::PaintSurface(HDC dc, const RECT& windowRect) const {
    if (!dc) {
        return;
    }

    DrawBackground(dc, windowRect);

    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(dc, font));
    SetBkMode(dc, TRANSPARENT);

    auto outlines = BuildGroupOutlines();

    const DropTarget* previewTarget = nullptr;
    if (m_drag.dragging && m_drag.target.active && !m_drag.target.outside) {
        previewTarget = &m_drag.target;
    } else if (m_externalDrop.active && m_externalDrop.target.active && !m_externalDrop.target.outside) {
        previewTarget = &m_externalDrop.target;
    }

    const int previewOffset = previewTarget ? kDropPreviewOffset : 0;
    int previewGroupIndex = -1;
    int previewTabIndex = -1;
    bool previewForGroup = false;
    if (previewTarget && previewOffset > 0) {
        previewGroupIndex = previewTarget->groupIndex;
        if (previewTarget->group) {
            previewForGroup = true;
        } else {
            previewTabIndex = previewTarget->tabIndex;
        }
    }

    bool previewGroupShifted = false;
    bool previewTabShifted = false;

    for (const auto& item : m_items) {
        VisualItem drawItem = item;
        if (previewOffset > 0 && previewTarget) {
            bool shift = false;
            if (previewForGroup && previewGroupIndex >= 0) {
                if (drawItem.data.location.groupIndex == previewGroupIndex) {
                    shift = true;
                    previewGroupShifted = true;
                }
            } else if (!previewForGroup && previewGroupIndex >= 0 && previewTabIndex >= 0) {
                if (drawItem.data.type == TabViewItemType::kTab &&
                    drawItem.data.location.groupIndex == previewGroupIndex &&
                    drawItem.data.location.tabIndex == previewTabIndex) {
                    shift = true;
                    previewTabShifted = true;
                }
            }
            if (shift) {
                OffsetRect(&drawItem.bounds, previewOffset, 0);
            }
        }

        if (drawItem.data.type == TabViewItemType::kGroupHeader) {
            DrawGroupHeader(dc, drawItem);
        } else {
            DrawTab(dc, drawItem);
        }
    }

    if (previewOffset > 0 && previewTarget) {
        for (auto& outline : outlines) {
            if (!outline.initialized || !outline.visible) {
                continue;
            }
            if (previewForGroup && previewGroupShifted && outline.groupIndex == previewGroupIndex) {
                OffsetRect(&outline.bounds, previewOffset, 0);
            } else if (!previewForGroup && previewTabShifted && outline.groupIndex == previewGroupIndex) {
                outline.bounds.right += previewOffset;
            }
        }
    }

	DrawGroupOutlines(dc, outlines);
	DrawDropIndicator(dc);
	DrawDragVisual(dc);

	// draw the '+' on empty islands last so it’s on top
	DrawEmptyIslandPluses(dc);

	if (oldFont) SelectObject(dc, oldFont);

}

COLORREF TabBandWindow::ResolveTabBackground(const TabViewItem& item) const {
    COLORREF base = item.selected ? m_themePalette.tabSelectedBase : m_themePalette.tabBase;
    if (item.selected) {
        base = BlendColors(base, m_accentColor, m_darkMode ? 0.45 : 0.35);
    }
    if (item.hasCustomOutline) {
        base = BlendColors(base, item.outlineColor, m_darkMode ? 0.35 : 0.25);
    }
    return base;
}

COLORREF TabBandWindow::ResolveGroupBackground(const TabViewItem& item) const {
    COLORREF base = m_themePalette.groupBase;
    if (item.selected) {
        base = BlendColors(base, m_accentColor, m_darkMode ? 0.4 : 0.25);
    }
    if (item.hasCustomOutline) {
        base = BlendColors(base, item.outlineColor, m_darkMode ? 0.35 : 0.25);
    }
    return base;
}

COLORREF TabBandWindow::ResolveTextColor(COLORREF background) const {
    return ComputeLuminance(background) > 0.55 ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

COLORREF TabBandWindow::ResolveTabTextColor(bool selected, COLORREF background) const {
    if (selected) {
        if (m_themePalette.tabSelectedTextValid) {
            return m_themePalette.tabSelectedText;
        }
        return GetSysColor(COLOR_HIGHLIGHTTEXT);
    }
    if (m_themePalette.tabTextValid) {
        return m_themePalette.tabText;
    }
    return ResolveTextColor(background);
}

COLORREF TabBandWindow::ResolveGroupTextColor(const TabViewItem& item, COLORREF background) const {
    if (item.selected && m_themePalette.tabSelectedTextValid) {
        return m_themePalette.tabSelectedText;
    }
    if (m_themePalette.groupTextValid) {
        return m_themePalette.groupText;
    }
    if (item.selected) {
        return GetSysColor(COLOR_HIGHLIGHTTEXT);
    }
    return ResolveTextColor(background);
}

bool TabBandWindow::FindEmptyIslandPlusAt(POINT pt, int* outGroupIndex) const {
	for (const auto& b : m_emptyIslandPlusButtons) {
		if (PtInRect(&b.rect, pt)) {
			if (outGroupIndex) *outGroupIndex = b.groupIndex;
			return true;
		}
	}
	return false;
}



std::vector<TabBandWindow::GroupOutline> TabBandWindow::BuildGroupOutlines() const {
	struct OutlineKey {
		int groupIndex;
		int row;
	};

	struct OutlineKeyHasher {
		size_t operator()(const OutlineKey& key) const noexcept {
			return (static_cast<size_t>(key.groupIndex) << 16) ^ static_cast<size_t>(key.row & 0xFFFF);
		}
	};

	struct OutlineKeyEqual {
		bool operator()(const OutlineKey& left, const OutlineKey& right) const noexcept {
			return left.groupIndex == right.groupIndex && left.row == right.row;
		}
	};

	std::unordered_map<OutlineKey, GroupOutline, OutlineKeyHasher, OutlineKeyEqual> outlines;

        auto accumulate = [&](const VisualItem& item, const RECT& bounds, COLORREF color, bool headerVisible,
                               bool updateColor) {
                OutlineKey key{ item.data.location.groupIndex, item.row };
                auto& outline = outlines[key];
                if (!outline.initialized) {
                        outline.groupIndex = key.groupIndex;
                        outline.row = key.row;
                        outline.bounds = bounds;
                        outline.color = color;
                        outline.initialized = true;
                        outline.visible = headerVisible;
                        outline.style = item.data.outlineStyle;
                } else {
                        outline.bounds.left = std::min(outline.bounds.left, bounds.left);
                        outline.bounds.top = std::min(outline.bounds.top, bounds.top);
                        outline.bounds.right = std::max(outline.bounds.right, bounds.right);
                        outline.bounds.bottom = std::max(outline.bounds.bottom, bounds.bottom);
			if (updateColor) {
				outline.color = color;
			}
			outline.visible = outline.visible || headerVisible;
		}
	};

	// 1) Grow outlines from real tabs (existing behavior)
	for (const auto& item : m_items) {
		if (item.data.type != TabViewItemType::kTab) continue;
		if (item.data.location.groupIndex < 0)	continue;
		if (!item.data.headerVisible)		continue;

		RECT rect = item.bounds;
		if (item.indicatorHandle) {
			rect.left = std::max(m_clientRect.left, rect.left - kIslandIndicatorWidth);
		}

		COLORREF outlineColor = ResolveIndicatorColor(item.hasGroupHeader ? &item.groupHeader : nullptr, item.data);
		if (item.data.selected) {
			outlineColor = DarkenColor(outlineColor, 0.2);
		}

		accumulate(item, rect, outlineColor, item.data.headerVisible, true);
	}

	// 2) Include visible indicators/placeholder bodies so outlines hug the handle too
	for (const auto& item : m_items) {
		if (item.data.type != TabViewItemType::kGroupHeader) continue;
		if (item.data.location.groupIndex < 0)	continue;
		if (!item.data.headerVisible || item.collapsedPlaceholder) continue;

		RECT rect = item.bounds;
		if (item.indicatorHandle) {
			rect.left = std::max(rect.left, m_clientRect.left);
			rect.right = std::max(rect.right, rect.left + kIslandIndicatorWidth);
		}

		accumulate(item, rect, ResolveIndicatorColor(&item.data, item.data), item.data.headerVisible, false);
	}

	// 3) Ensure empty islands still get a small outline body after the indicator
	for (const auto& item : m_items) {
	        if (item.data.type != TabViewItemType::kGroupHeader) continue;
	        if (!item.indicatorHandle) continue;

	        const int gi = item.data.location.groupIndex;
	        if (gi < 0) continue;
	        if (!item.data.headerVisible || item.collapsedPlaceholder) continue;

	        // NOTE: visibleTabs is a member of TabViewItem, not VisualItem
	        if (item.data.visibleTabs > 0) continue;

	        // Synthesize a tiny body area to the right of the indicator so the island outline has width.
	        const RECT body = item.bounds;          // indicator rect
	        const LONG left = body.right;           // start immediately after indicator
	        const LONG available = std::max<LONG>(0, m_clientRect.right - left);
	        LONG width = std::min<LONG>(available, kEmptyIslandBodyMaxWidth);
	        if (width < kEmptyIslandBodyMinWidth) {
	                width = std::max<LONG>(width, static_cast<LONG>(kEmptyIslandBodyMinWidth));
	        }
	        const LONG right = left + width;

	        RECT rect{
	                std::max<LONG>(m_clientRect.left,  left - kIslandIndicatorWidth),
	                std::max<LONG>(m_clientRect.top,   body.top),
	                std::min<LONG>(m_clientRect.right, right),
	                std::min<LONG>(m_clientRect.bottom, body.bottom)
	        };

	        accumulate(item, rect, ResolveIndicatorColor(&item.data, item.data), item.data.headerVisible, true);
	}
	std::vector<GroupOutline> result;
	result.reserve(outlines.size());
	for (auto& entry : outlines) {
		if (entry.second.initialized && entry.second.visible) {
			result.emplace_back(entry.second);
		}
	}
	std::sort(result.begin(), result.end(),
		[](const GroupOutline& a, const GroupOutline& b) {
			if (a.bounds.top == b.bounds.top) {
				if (a.bounds.left == b.bounds.left) {
					return a.groupIndex < b.groupIndex;
				}
				return a.bounds.left < b.bounds.left;
			}
			return a.bounds.top < b.bounds.top;
		});
	return result;
}

void TabBandWindow::DrawGroupOutlines(HDC dc, const std::vector<GroupOutline>& outlines) const {
    const auto createPenForOutline = [](const GroupOutline& outline) -> HPEN {
        DWORD baseStyle = PS_SOLID;
        switch (outline.style) {
            case TabGroupOutlineStyle::kDashed:
                baseStyle = PS_DASH;
                break;
            case TabGroupOutlineStyle::kDotted:
                baseStyle = PS_DOT;
                break;
            case TabGroupOutlineStyle::kSolid:
            default:
                baseStyle = PS_SOLID;
                break;
        }

        if (baseStyle == PS_SOLID) {
            return CreatePen(PS_SOLID, kIslandOutlineThickness, outline.color);
        }

        LOGBRUSH brush{};
        brush.lbStyle = BS_SOLID;
        brush.lbColor = outline.color;
        HPEN pen = ExtCreatePen(PS_GEOMETRIC | baseStyle, std::max(1, kIslandOutlineThickness), &brush, 0, nullptr);
        if (pen) {
            return pen;
        }

        pen = CreatePen(baseStyle, 1, outline.color);
        if (pen) {
            return pen;
        }

        return CreatePen(PS_SOLID, kIslandOutlineThickness, outline.color);
    };

    for (const auto& outline : outlines) {
        if (!outline.initialized) {
            continue;
        }
        RECT rect = outline.bounds;
        rect.left = std::max(rect.left, m_clientRect.left);
        rect.top = std::max(rect.top, m_clientRect.top);
        rect.right = std::min(rect.right + 1, m_clientRect.right);
        rect.bottom = std::min(rect.bottom, m_clientRect.bottom);
        if (rect.right <= rect.left || rect.bottom <= rect.top) {
            continue;
        }

        HPEN pen = createPenForOutline(outline);
        if (!pen) {
            continue;
        }
        HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));

        const int left = rect.left;
        const int right = rect.right;
        const int top = rect.top;
        const int bottom = rect.bottom - 1;

        MoveToEx(dc, left, top, nullptr);
        LineTo(dc, right, top);
        MoveToEx(dc, left, top, nullptr);
        LineTo(dc, left, bottom);
        MoveToEx(dc, left, bottom, nullptr);
        LineTo(dc, right, bottom);
        MoveToEx(dc, right, top, nullptr);
        LineTo(dc, right, bottom);

        SelectObject(dc, oldPen);
        DeleteObject(pen);
    }
}

// TabBandWindow.cpp
LRESULT CALLBACK TabBandWindow::RebarSubclassProc(HWND hwnd, UINT msg,
	WPARAM wParam, LPARAM lParam, UINT_PTR /*id*/, DWORD_PTR refData) {
	auto* self = reinterpret_cast<TabBandWindow*>(refData);
	switch (msg) {
	case WM_ERASEBKGND: {
		if (!self) return 1;
		HDC hdc = reinterpret_cast<HDC>(wParam);
		if (!hdc) return 1;
		RECT rc{}; GetClientRect(hwnd, &rc);
		HBRUSH br = CreateSolidBrush(self->m_themePalette.rebarBackground);
		if (br) { FillRect(hdc, &rc, br); DeleteObject(br); }
		return 1; // handled; prevents bright erase
	}
	case WM_PRINTCLIENT: {
		// Some children ask the rebar to paint its bg via WM_PRINTCLIENT
		if (!self) break;
		HDC hdc = reinterpret_cast<HDC>(wParam);
		if (hdc) {
			RECT rc{}; GetClientRect(hwnd, &rc);
			HBRUSH br = CreateSolidBrush(self->m_themePalette.rebarBackground);
			if (br) { FillRect(hdc, &rc, br); DeleteObject(br); }
		}
		break; // let children continue drawing
	}
	case WM_NCDESTROY:
		RemoveWindowSubclass(hwnd, RebarSubclassProc, 0);
		if (self) self->m_rebarSubclassed = false;
		break;
	}
	return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void TabBandWindow::InstallRebarDarkSubclass() {
	if (!m_parentRebar || !IsWindow(m_parentRebar) || m_rebarSubclassed) return;

	// Keep Explorer's theme resources; we only overpaint the bg.
	SetWindowTheme(m_parentRebar, L"Explorer", nullptr);  // was nullptr, which destabilized the band site
	ApplyImmersiveDarkMode(m_parentRebar, m_darkMode);

	if (SetWindowSubclass(m_parentRebar, RebarSubclassProc, 0,
		reinterpret_cast<DWORD_PTR>(this))) {
		m_rebarSubclassed = true;
		// Gentle repaint: no frame, no children, and no erase
		RedrawWindow(m_parentRebar, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
	}
}

void TabBandWindow::AdjustBandHeightToRow() {
	if (!m_parentRebar || !IsWindow(m_parentRebar)) return;
	if (m_rebarBandIndex < 0) m_rebarBandIndex = FindRebarBandIndex();
	if (m_rebarBandIndex < 0) return;

	// Determine a row height similar to RebuildLayout
	int rowHeight = 0;
	int maxRowIndex = -1;
	for (const auto& it : m_items) {
		const int h = std::max(0, static_cast<int>(it.bounds.bottom - it.bounds.top));
		if (h > rowHeight) {
			rowHeight = h;
		}
		if (it.row > maxRowIndex) {
			maxRowIndex = it.row;
		}
	}
        if (rowHeight <= 0) rowHeight = 24;
        rowHeight = std::max(rowHeight, kButtonHeight - kButtonMargin);

        int rowsFromItems = (maxRowIndex >= 0) ? (maxRowIndex + 1) : 0;
        int rows = std::max(rowsFromItems, m_lastRowCount);
        rows = std::max(rows, 1);
        rows = std::min(rows, kMaxTabRows);
        if (rows == m_lastAppliedRowCount) {
            return;
        }
        m_lastAppliedRowCount = rows;
        int desired = rows * rowHeight + (rows - 1) * kRowGap;
        desired = std::max(desired, kButtonHeight + kButtonMargin * 2);

        REBARBANDINFOW bi{ sizeof(bi) };
	bi.fMask = RBBIM_CHILDSIZE;
	bi.cyChild = desired;
	bi.cyMinChild = desired;
	bi.cyIntegral = 1;
	SendMessageW(m_parentRebar, RB_SETBANDINFO, m_rebarBandIndex, reinterpret_cast<LPARAM>(&bi));

	// Expand the band to its full height without forcing an erase (flicker-free)
	SendMessageW(m_parentRebar, RB_MAXIMIZEBAND, m_rebarBandIndex, 0);
	RedrawWindow(m_parentRebar, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
}


void TabBandWindow::RefreshTheme() {
	if (m_refreshingTheme) return;
	m_refreshingTheme = true;
	struct Guard { TabBandWindow* w; ~Guard() { if (w) w->m_refreshingTheme = false; } } g{ this };

	CloseThemeHandles();
	m_toolbarGripWidth = kToolbarGripWidth;
	if (!m_hwnd) { /* existing reset block unchanged */ return; }

	//SetWindowTheme(m_hwnd, L"Explorer", nullptr);
	const bool darkMode = IsSystemDarkMode();
	if (!m_windowDarkModeInitialized || darkMode != m_windowDarkModeValue) {
		ApplyImmersiveDarkMode(m_hwnd, darkMode);
		m_windowDarkModeInitialized = true;
		m_windowDarkModeValue = darkMode;
	}
	m_darkMode = darkMode;

	// NEW: also flip the parent rebar immediately
	EnsureRebarIntegration();
    AdjustBandHeightToRow();
	if (m_parentRebar) {
        InstallRebarDarkSubclass();   // NEW: we own the bar bg now
        UpdateRebarColors();  // NEW
	}

	UpdateAccentColor();
	ResetThemePalette();
	m_tabTheme = OpenThemeData(m_hwnd, L"Tab");
	m_rebarTheme = OpenThemeData(m_hwnd, L"Rebar");
	m_windowTheme = OpenThemeData(m_hwnd, L"Window");
	UpdateThemePalette();
	UpdateToolbarMetrics();
	UpdateNewTabButtonTheme();
        RebuildLayout();

}



void TabBandWindow::UpdateAccentColor() {
    DWORD color = 0;
    BOOL opaque = FALSE;
    if (SUCCEEDED(DwmGetColorizationColor(&color, &opaque))) {
        m_accentColor = RGB((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
    } else {
        m_accentColor = GetSysColor(COLOR_HOTLIGHT);
    }
}

void TabBandWindow::ResetThemePalette() {
    m_themePalette.tabTextValid = false;
    m_themePalette.tabSelectedTextValid = false;
    m_themePalette.groupTextValid = false;

    const COLORREF windowColor = GetSysColor(COLOR_WINDOW);
    const COLORREF buttonColor = GetSysColor(COLOR_BTNFACE);
    const COLORREF windowBase = AdjustForDarkTone(windowColor, 0.55, m_darkMode);
    const COLORREF buttonBase = AdjustForDarkTone(buttonColor, 0.4, m_darkMode);

    m_themePalette.rebarBackground = m_darkMode ? BlendColors(buttonBase, windowBase, 0.55) : windowBase;

    m_themePalette.borderTop = m_darkMode ? BlendColors(m_themePalette.rebarBackground, RGB(0, 0, 0), 0.6)
                                          : GetSysColor(COLOR_3DSHADOW);
    m_themePalette.borderBottom = m_darkMode ? BlendColors(m_themePalette.rebarBackground, RGB(255, 255, 255), 0.2)
                                             : GetSysColor(COLOR_3DLIGHT);

    if (m_darkMode) {
        const double rebarLuminance = ComputeLuminance(m_themePalette.rebarBackground);
        const bool shouldBrighten = rebarLuminance <= 0.35;
        const double blend = shouldBrighten ? 0.28 : 0.22;
        const COLORREF contrastTarget = shouldBrighten ? RGB(255, 255, 255) : RGB(0, 0, 0);
        m_themePalette.tabBase = BlendColors(m_themePalette.rebarBackground, contrastTarget, blend);
    } else {
        m_themePalette.tabBase = windowBase;
    }
    m_themePalette.tabSelectedBase = BlendColors(m_themePalette.tabBase, m_accentColor, m_darkMode ? 0.5 : 0.4);
    m_themePalette.tabText = GetSysColor(COLOR_WINDOWTEXT);
    m_themePalette.tabSelectedText = GetSysColor(COLOR_HIGHLIGHTTEXT);

    const double groupBlend = m_darkMode ? 0.6 : 0.25;
    m_themePalette.groupBase = BlendColors(buttonBase, windowBase, groupBlend);
    m_themePalette.groupText = GetSysColor(COLOR_WINDOWTEXT);
}

void TabBandWindow::UpdateThemePalette() {
    if (m_darkMode) {
        ApplyOptionColorOverrides();
        return;
    }

    if (m_rebarTheme) {
        COLORREF color = 0;
        if (SUCCEEDED(GetThemeColor(m_rebarTheme, RP_BAND, 0, TMT_FILLCOLORHINT, &color))) {
            m_themePalette.rebarBackground = color;
        }
        if (SUCCEEDED(GetThemeColor(m_rebarTheme, RP_BAND, 0, TMT_BORDERCOLORHINT, &color))) {
            m_themePalette.borderTop = color;
        }
        if (SUCCEEDED(GetThemeColor(m_rebarTheme, RP_BAND, 0, TMT_EDGEHIGHLIGHTCOLOR, &color))) {
            m_themePalette.borderBottom = color;
        }
    }

    if (m_tabTheme) {
        COLORREF color = 0;
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_BODY, 0, TMT_FILLCOLORHINT, &color))) {
            m_themePalette.tabBase = color;
            m_themePalette.groupBase = color;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, TIS_SELECTED, TMT_FILLCOLORHINT, &color))) {
            m_themePalette.tabSelectedBase = BlendColors(color, m_accentColor, 0.25);
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, TIS_SELECTED, TMT_TEXTCOLOR, &color))) {
            m_themePalette.tabSelectedText = color;
            m_themePalette.tabSelectedTextValid = true;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, TIS_NORMAL, TMT_TEXTCOLOR, &color))) {
            m_themePalette.tabText = color;
            m_themePalette.tabTextValid = true;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_BODY, 0, TMT_TEXTCOLOR, &color))) {
            m_themePalette.groupText = color;
            m_themePalette.groupTextValid = true;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_BODY, 0, TMT_BORDERCOLORHINT, &color))) {
            m_themePalette.borderBottom = color;
        }
    }

    ApplyOptionColorOverrides();

    if (m_darkMode) {
        m_themePalette.borderTop = BlendColors(m_themePalette.borderTop, RGB(0, 0, 0), 0.3);
        m_themePalette.borderBottom = BlendColors(m_themePalette.borderBottom, RGB(255, 255, 255), 0.15);
    }
}

void TabBandWindow::ApplyOptionColorOverrides() {
    auto& store = OptionsStore::Instance();
    store.Load();
    const ShellTabsOptions options = store.Get();

    auto pickTextColor = [](COLORREF background) -> COLORREF {
        return ComputeLuminance(background) > 0.55 ? RGB(0, 0, 0) : RGB(255, 255, 255);
    };

    m_progressStartColor = m_accentColor;
    m_progressEndColor = BlendColors(m_accentColor, RGB(255, 255, 255), m_darkMode ? 0.1 : 0.3);
    if (options.useCustomProgressBarGradientColors) {
        m_progressStartColor = options.progressBarGradientStartColor;
        m_progressEndColor = options.progressBarGradientEndColor;
    }

    if (options.useCustomTabUnselectedColor) {
        m_themePalette.tabBase = options.customTabUnselectedColor;
        const COLORREF textColor = pickTextColor(m_themePalette.tabBase);
        m_themePalette.tabText = textColor;
        m_themePalette.tabTextValid = true;
    }

    if (options.useCustomTabSelectedColor) {
        m_themePalette.tabSelectedBase = options.customTabSelectedColor;
        const COLORREF textColor = pickTextColor(m_themePalette.tabSelectedBase);
        m_themePalette.tabSelectedText = textColor;
        m_themePalette.tabSelectedTextValid = true;
    }
}

bool TabBandWindow::IsRebarWindow(HWND hwnd) {
    if (!hwnd) {
        return false;
    }
    wchar_t className[64] = {};
    if (!RealGetWindowClassW(hwnd, className, ARRAYSIZE(className))) {
        if (!GetClassNameW(hwnd, className, ARRAYSIZE(className))) {
            return false;
        }
    }
    return _wcsicmp(className, REBARCLASSNAMEW) == 0 || _wcsicmp(className, L"ReBarWindow32") == 0;
}

int TabBandWindow::FindRebarBandIndex() const {
    if (!m_parentRebar || !IsWindow(m_parentRebar)) {
        return -1;
    }
    const LRESULT count = SendMessageW(m_parentRebar, RB_GETBANDCOUNT, 0, 0);
    if (count <= 0) {
        return -1;
    }
    for (int index = 0; index < count; ++index) {
        REBARBANDINFOW info{sizeof(info)};
        info.fMask = RBBIM_CHILD;
        if (SendMessageW(m_parentRebar, RB_GETBANDINFO, index, reinterpret_cast<LPARAM>(&info))) {
            if (info.hwndChild == m_hwnd) {
                return index;
            }
        }
    }
    return -1;
}

void TabBandWindow::RefreshRebarMetrics() {
	if (!m_parentRebar || !IsWindow(m_parentRebar)) return;
	if (m_rebarBandIndex < 0) m_rebarBandIndex = FindRebarBandIndex();
	if (m_rebarBandIndex < 0) return;

	// 1) Ensure the band style shows ONE rebar grip, and drop the etched edge in dark mode.
	REBARBANDINFOW info{ sizeof(info) };
	info.fMask = RBBIM_STYLE | RBBIM_CHILD;
	if (!SendMessageW(m_parentRebar, RB_GETBANDINFO, m_rebarBandIndex,
		reinterpret_cast<LPARAM>(&info))) {
		return;
	}

	DWORD st = info.fStyle;
	st &= ~RBBS_NOGRIPPER;      // allow a grip
	st |= RBBS_GRIPPERALWAYS;  // and make it visible
	if (m_darkMode) st &= ~RBBS_CHILDEDGE;
	else            st |= RBBS_CHILDEDGE;

	if (st != info.fStyle) {
		REBARBANDINFOW s{ sizeof(s) };
		s.fMask = RBBIM_STYLE;
		s.fStyle = st;
		SendMessageW(m_parentRebar, RB_SETBANDINFO, m_rebarBandIndex,
			reinterpret_cast<LPARAM>(&s));
	}

	// 2) Compute the actual grip/left-border width so tab layout starts AFTER the dots.
	RECT borders{ 0, 0, 0, 0 };
	LONG rbGrip = 0;
	if (SendMessageW(m_parentRebar, RB_GETBANDBORDERS, m_rebarBandIndex,
		reinterpret_cast<LPARAM>(&borders))) {
		rbGrip = borders.left;  // includes grip + left padding
	}

	LONG themeGrip = 0;
	if (m_rebarTheme) {
		HDC hdc = GetDC(m_parentRebar);
		if (hdc) {
			SIZE part{};
			if (SUCCEEDED(GetThemePartSize(m_rebarTheme, hdc,
				RP_GRIPPERVERT, 0, nullptr, TS_TRUE, &part))) {
				themeGrip = part.cx;
			}
			ReleaseDC(m_parentRebar, hdc);
		}
	}

	const LONG want = std::max<LONG>(rbGrip, themeGrip);// tiny safety margin
	m_toolbarGripWidth = static_cast<int>(want);

	// 3) Colors: set bar-wide bk color and per-band bk so NO bright area remains.
	const COLORREF barBk = m_themePalette.rebarBackground; // explicit, even in light
	SendMessageW(m_parentRebar, RB_SETBKCOLOR, 0, static_cast<LPARAM>(barBk));

	REBARBANDINFOW colorInfo{ sizeof(colorInfo) };
	colorInfo.fMask = RBBIM_COLORS;
	colorInfo.clrFore = CLR_DEFAULT;
	colorInfo.clrBack = barBk;
	SendMessageW(m_parentRebar, RB_SETBANDINFO, m_rebarBandIndex,
		reinterpret_cast<LPARAM>(&colorInfo));

	// Tone down etched highlights so the bar doesn't glow in dark mode.
	COLORSCHEME cs{};
	cs.dwSize = sizeof(cs);
	cs.clrBtnHighlight = m_darkMode ? barBk : CLR_DEFAULT;
	cs.clrBtnShadow = m_darkMode ? barBk : CLR_DEFAULT;
	SendMessageW(m_parentRebar, RB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&cs));

	// Repaint without forcing an erase (prevents flicker).
	RedrawWindow(m_parentRebar, nullptr, nullptr,
		RDW_INVALIDATE | RDW_FRAME );
}


void TabBandWindow::EnsureRebarIntegration() {
        if (!m_hwnd) return;
        if (!m_parentRebar || !IsWindow(m_parentRebar)) {
                HWND parent = GetParent(m_hwnd);
                while (parent && !IsRebarWindow(parent)) parent = GetParent(parent);
                m_parentRebar = parent;
                m_rebarBandIndex = -1;
                if (m_parentRebar) {
            InstallRebarDarkSubclass();   // NEW: we own the bar bg now
            UpdateRebarColors();  // NEW
                }
        }

        if (!m_parentRebar) return;

        HWND frame = GetAncestor(m_parentRebar, GA_ROOT);
        UpdateAvailableDockMaskFromFrame(frame);

        const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(m_parentRebar, GWL_STYLE));
        const TabBandDockMode detectedMode = DockModeFromRebarStyle(style);
        if (detectedMode != TabBandDockMode::kAutomatic && detectedMode != m_currentDockMode) {
            m_currentDockMode = detectedMode;
            if (m_owner) {
                m_owner->OnDockingModeChanged(detectedMode);
            }
        }

        const int index = FindRebarBandIndex();
        if (index >= 0) {
                m_rebarBandIndex = index;
        AdjustBandHeightToRow();
                RefreshRebarMetrics();
        }
}


void TabBandWindow::UpdateToolbarMetrics() {
    m_toolbarGripWidth = kToolbarGripWidth;
    EnsureRebarIntegration();
    if (m_parentRebar && m_rebarBandIndex >= 0) {
        RECT borders{0, 0, 0, 0};
        if (SendMessageW(m_parentRebar, RB_GETBANDBORDERS, m_rebarBandIndex, reinterpret_cast<LPARAM>(&borders))) {
            const LONG candidate = std::max<LONG>(borders.left, 8L);
            if (candidate > 0) {
                m_toolbarGripWidth = candidate;
                return;
            }
        }
    }

    if (!m_hwnd || !m_rebarTheme) {
        return;
    }

    HDC dc = GetDC(m_hwnd);
    if (!dc) {
        return;
    }

    int part = RP_GRIPPER;
    SIZE gripSize{0, 0};
    HRESULT hr = GetThemePartSize(m_rebarTheme, dc, part, 0, nullptr, TS_TRUE, &gripSize);
    if (FAILED(hr) || gripSize.cx <= 0) {
        part = RP_GRIPPERVERT;
        gripSize = {0, 0};
        hr = GetThemePartSize(m_rebarTheme, dc, part, 0, nullptr, TS_TRUE, &gripSize);
    }

    if (SUCCEEDED(hr) && gripSize.cx > 0) {
        int width = gripSize.cx;
        MARGINS margins{0, 0, 0, 0};
        if (SUCCEEDED(GetThemeMargins(m_rebarTheme, dc, part, 0, TMT_CONTENTMARGINS, nullptr, &margins))) {
            width += margins.cxLeftWidth + margins.cxRightWidth;
        }
        if (width > 0) {
            m_toolbarGripWidth = std::max(width, 8);
        }
    }

    ReleaseDC(m_hwnd, dc);
}

void TabBandWindow::CloseThemeHandles() {
    if (m_tabTheme) {
        CloseThemeData(m_tabTheme);
        m_tabTheme = nullptr;
    }
    if (m_rebarTheme) {
        CloseThemeData(m_rebarTheme);
        m_rebarTheme = nullptr;
    }
    if (m_windowTheme) {
        CloseThemeData(m_windowTheme);
        m_windowTheme = nullptr;
    }
}

void TabBandWindow::UpdateNewTabButtonTheme() {
    if (!m_newTabButton) {
        m_buttonDarkModeInitialized = false;
        m_buttonDarkModeValue = false;
        return;
    }
    //SetWindowTheme(m_newTabButton, L"Explorer", nullptr);
    if (!m_buttonDarkModeInitialized || m_buttonDarkModeValue != m_darkMode) {
        ApplyImmersiveDarkMode(m_newTabButton, m_darkMode);
        m_buttonDarkModeInitialized = true;
        m_buttonDarkModeValue = m_darkMode;
    }
    SendMessageW(m_newTabButton, WM_SETFONT, reinterpret_cast<WPARAM>(GetDefaultFont()), FALSE);
    InvalidateRect(m_newTabButton, nullptr, FALSE);
}

bool TabBandWindow::IsSystemDarkMode() const {
    DWORD value = 1;
    DWORD size = sizeof(value);
    const LSTATUS status = RegGetValueW(HKEY_CURRENT_USER, kThemePreferenceKey, kThemePreferenceValue, RRF_RT_DWORD,
                                        nullptr, &value, &size);
    if (status != ERROR_SUCCESS) {
        return false;
    }
    return value == 0;
}

void TabBandWindow::DrawGroupHeader(HDC dc, const VisualItem& item) const {
    RECT rect = item.bounds;
    RECT indicator = rect;
    indicator.right = std::min(indicator.left + kIslandIndicatorWidth, indicator.right);
    indicator.top = rect.top;
    indicator.bottom = rect.bottom;
    if (indicator.right > indicator.left) {
        COLORREF indicatorColor = item.data.hasCustomOutline ? item.data.outlineColor : m_accentColor;
        if (item.data.selected) {
            indicatorColor = BlendColors(indicatorColor, RGB(0, 0, 0), 0.2);
        }
        HBRUSH brush = CreateSolidBrush(indicatorColor);
        if (brush) {
            FillRect(dc, &indicator, brush);
            DeleteObject(brush);
        }
    }

    if (!item.collapsedPlaceholder) {
        return;
    }
}

RECT TabBandWindow::ComputeCloseButtonRect(const VisualItem& item) const {
    RECT rect{0, 0, 0, 0};
    if (item.data.type != TabViewItemType::kTab) {
        return rect;
    }
    const int height = item.bounds.bottom - item.bounds.top;
    if (height <= kCloseButtonVerticalPadding * 2) {
        return rect;
    }
    const int badgeWidth = std::max(0, item.badgeWidth);
    const int availableWidth = item.bounds.right - item.bounds.left;
    const int minimumWidth = kCloseButtonSize + kCloseButtonEdgePadding + kCloseButtonSpacing + badgeWidth + kPaddingX + 8;
    if (availableWidth < minimumWidth) {
        return rect;
    }
    int size = std::min(kCloseButtonSize, height - kCloseButtonVerticalPadding * 2);
    if (m_windowTheme && m_hwnd) {
        HDC dc = GetDC(m_hwnd);
        if (dc) {
            SIZE themeSize{0, 0};
            if (SUCCEEDED(GetThemePartSize(m_windowTheme, dc, WP_SMALLCLOSEBUTTON, 0, nullptr, TS_TRUE, &themeSize))) {
                const int candidate = std::max(themeSize.cx, themeSize.cy);
                if (candidate > 0) {
                    size = std::min(candidate, height - kCloseButtonVerticalPadding * 2);
                }
            }
            ReleaseDC(m_hwnd, dc);
        }
    }
    if (size <= 0) {
        return rect;
    }
    const int right = item.bounds.right - kCloseButtonEdgePadding;
    const int left = right - size;
    const int top = item.bounds.top + (height - size) / 2;
    rect = {left, top, right, top + size};
    return rect;
}

void TabBandWindow::DrawTab(HDC dc, const VisualItem& item) const {
    RECT rect = item.bounds;
    const bool selected = item.data.selected;
    const TabViewItem* indicatorSource = item.hasGroupHeader ? &item.groupHeader : nullptr;
    const bool hasAccent = item.data.hasCustomOutline ||
                           (indicatorSource && indicatorSource->hasCustomOutline);
    COLORREF accentColor = hasAccent ? ResolveIndicatorColor(indicatorSource, item.data) : m_accentColor;

    const int islandIndicator = item.indicatorHandle ? kIslandIndicatorWidth : 0;
    RECT tabRect = rect;
    tabRect.left += islandIndicator;

    int state = selected ? TIS_SELECTED : TIS_NORMAL;
    COLORREF computedBackground = ResolveTabBackground(item.data);
    COLORREF textColor = ResolveTabTextColor(selected, computedBackground);
    bool usedTheme = false;
    if (m_tabTheme && !m_darkMode) {
        if (SUCCEEDED(DrawThemeBackground(m_tabTheme, dc, TABP_TABITEM, state, &tabRect, nullptr))) {
            usedTheme = true;
            COLORREF themeText = 0;
            if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, state, TMT_TEXTCOLOR, &themeText))) {
                textColor = themeText;
            } else {
                textColor = ResolveTabTextColor(selected, computedBackground);
            }
        }
    }

	if (!usedTheme) {
		COLORREF backgroundColor = computedBackground;
		textColor = ResolveTabTextColor(selected, backgroundColor);
		COLORREF baseBorder = m_darkMode
			? BlendColors(backgroundColor, RGB(255, 255, 255), selected ? 0.1 : 0.05)
			: BlendColors(backgroundColor, RGB(0, 0, 0), selected ? 0.15 : 0.1);
		COLORREF borderColor = hasAccent
			? BlendColors(accentColor, RGB(0, 0, 0), selected ? 0.25 : 0.15)
			: baseBorder;

		RECT shapeRect = tabRect;
		// Keep the fill inside the island outline for ALL states.
		// The outline is drawn at rect.bottom - 1, so don’t paint the bottom row.
		const LONG bottomLimit = rect.bottom - 1;
		if (shapeRect.bottom > bottomLimit) {
			shapeRect.bottom = bottomLimit;
		}

		const int radius = kTabCornerRadius;
		POINT points[] = {
			{shapeRect.left,        shapeRect.bottom},
			{shapeRect.left,        shapeRect.top + radius},
			{shapeRect.left + radius, shapeRect.top},
			{shapeRect.right - radius,shapeRect.top},
			{shapeRect.right,       shapeRect.top + radius},
			{shapeRect.right,       shapeRect.bottom}
		};

		HRGN region = CreatePolygonRgn(points, ARRAYSIZE(points), WINDING);
		if (region) {
			HBRUSH brush = CreateSolidBrush(backgroundColor);
			if (brush) {
				FillRgn(dc, region, brush);
				DeleteObject(brush);
			}
			HPEN pen = CreatePen(PS_SOLID, 1, borderColor);
			if (pen) {
				HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
				HBRUSH oldBrush = static_cast<HBRUSH>(SelectObject(dc, GetStockObject(HOLLOW_BRUSH)));
				Polygon(dc, points, ARRAYSIZE(points));
				SelectObject(dc, oldBrush);
				SelectObject(dc, oldPen);
				DeleteObject(pen);
			}
			DeleteObject(region);
		}

		// Bottom separator stays at rect.bottom - 1, which now aligns perfectly.
		COLORREF bottomLineColor = selected ? backgroundColor
			: (m_darkMode ? BlendColors(backgroundColor, RGB(0, 0, 0), 0.25)
				: GetSysColor(COLOR_3DLIGHT));
		HPEN bottomPen = CreatePen(PS_SOLID, 1, bottomLineColor);
		if (bottomPen) {
			HPEN oldPen = static_cast<HPEN>(SelectObject(dc, bottomPen));
			MoveToEx(dc, tabRect.left + 1, rect.bottom - 1, nullptr);
			LineTo(dc, rect.right - 1, rect.bottom - 1);
			SelectObject(dc, oldPen);
			DeleteObject(bottomPen);
		}

		computedBackground = backgroundColor;
	}

    if (item.indicatorHandle) {
        RECT indicatorRect = tabRect;
        indicatorRect.left = rect.left;
        indicatorRect.right = indicatorRect.left + kIslandIndicatorWidth;
        indicatorRect.top = rect.top;
        indicatorRect.bottom = rect.bottom;
        COLORREF indicatorColor = hasAccent ? accentColor
                                            : (m_darkMode ? RGB(120, 120, 180) : GetSysColor(COLOR_HOTLIGHT));
        if (selected) {
            indicatorColor = DarkenColor(indicatorColor, 0.2);
        }
        HBRUSH indicatorBrush = CreateSolidBrush(indicatorColor);
        if (indicatorBrush) {
            FillRect(dc, &indicatorRect, indicatorBrush);
            DeleteObject(indicatorBrush);
        }
    }

    RECT closeRect = ComputeCloseButtonRect(item);

    int trailingBoundary = rect.right - kPaddingX;
    if (closeRect.right > closeRect.left) {
        const int closeLeft = static_cast<int>(closeRect.left);
        trailingBoundary = std::min(trailingBoundary, closeLeft - kCloseButtonSpacing);
    }

    int textRight = trailingBoundary;

    int textLeft = rect.left + islandIndicator + kPaddingX;
    if (item.icon) {
        const int availableHeight = rect.bottom - rect.top;
        const int iconHeight = std::min(item.iconHeight, availableHeight - 4);
        const int iconWidth = item.iconWidth;
        const int iconY = rect.top + (availableHeight - iconHeight) / 2;
        DrawIconEx(dc, textLeft, iconY, item.icon, iconWidth, iconHeight, 0, nullptr, DI_NORMAL);
        textLeft += iconWidth + kIconGap;
    }

    const bool hasProgress = item.data.progress.visible;
    if (hasProgress) {
        DrawTabProgress(dc, item, textLeft, textRight, tabRect, computedBackground);
    }

    RECT textRect = rect;
    textRect.left = textLeft;
    textRect.top += 3;
    textRect.bottom = hasProgress ? tabRect.bottom - 6 : tabRect.bottom - 3;
    if (textRect.bottom <= textRect.top) {
        textRect.bottom = textRect.top + 1;
    }

    SetTextColor(dc, textColor);

    textRect.right = std::max(textLeft + 1, textRight);
    DrawTextW(dc, item.data.name.c_str(), static_cast<int>(item.data.name.size()), &textRect,
              DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);

    if (closeRect.right > closeRect.left) {
        const bool closeHot = (m_hotCloseIndex != kInvalidIndex && m_hotCloseIndex == item.index);
        bool closePressed = false;
        if (closeHot && m_drag.closeClick && m_drag.closeItemIndex < m_items.size()) {
            closePressed = (m_items[m_drag.closeItemIndex].index == item.index);
        }

        int closeState = CBS_NORMAL;
        if (closePressed) {
            closeState = CBS_PUSHED;
        } else if (closeHot) {
            closeState = CBS_HOT;
        }

        bool themedClose = false;
        if (m_windowTheme) {
            if (SUCCEEDED(DrawThemeBackground(m_windowTheme, dc, WP_SMALLCLOSEBUTTON, closeState, &closeRect, nullptr))) {
                themedClose = true;
            } else if (SUCCEEDED(DrawThemeBackground(m_windowTheme, dc, WP_CLOSEBUTTON, closeState, &closeRect, nullptr))) {
                themedClose = true;
            }
        }

        if (!themedClose) {
            COLORREF closeBackground = closeHot ? RGB(232, 17, 35)
                                                : (m_darkMode ? BlendColors(computedBackground, RGB(255, 255, 255), 0.15)
                                                              : BlendColors(computedBackground, RGB(0, 0, 0), 0.12));
            if (closePressed) {
                closeBackground = BlendColors(closeBackground, RGB(0, 0, 0), 0.2);
            }

            HBRUSH closeBrush = CreateSolidBrush(closeBackground);
            if (closeBrush) {
                FillRect(dc, &closeRect, closeBrush);
                DeleteObject(closeBrush);
            }

            COLORREF borderColor = closeHot ? BlendColors(closeBackground, RGB(0, 0, 0), 0.2)
                                            : BlendColors(closeBackground, RGB(0, 0, 0), m_darkMode ? 0.6 : 0.4);
            HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
            if (borderPen) {
                HPEN oldPen = static_cast<HPEN>(SelectObject(dc, borderPen));
                MoveToEx(dc, closeRect.left, closeRect.top, nullptr);
                LineTo(dc, closeRect.right, closeRect.top);
                LineTo(dc, closeRect.right, closeRect.bottom);
                LineTo(dc, closeRect.left, closeRect.bottom);
                LineTo(dc, closeRect.left, closeRect.top);
                SelectObject(dc, oldPen);
                DeleteObject(borderPen);
            }

            RECT glyphRect = closeRect;
            COLORREF glyphColor = closeHot ? RGB(255, 255, 255) : ResolveTextColor(closeBackground);
            COLORREF previousColor = SetTextColor(dc, glyphColor);
            DrawTextW(dc, L"x", 1, &glyphRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            if (previousColor != CLR_INVALID) {
                SetTextColor(dc, previousColor);
            } else {
                SetTextColor(dc, textColor);
            }
        }
    }
}


void TabBandWindow::DrawTabProgress(HDC dc, const VisualItem& item, int left, int right,
                                    const RECT& tabRect, COLORREF background) const {
    if (!dc || left >= right) {
        return;
    }

    RECT outer{left, std::max(tabRect.top + 4, tabRect.bottom - 6), right, tabRect.bottom - 2};
    if (outer.bottom <= outer.top || outer.right <= outer.left) {
        return;
    }

    const COLORREF trackColor = m_darkMode ? BlendColors(background, RGB(255, 255, 255), 0.2)
                                           : BlendColors(background, RGB(0, 0, 0), 0.15);
    if (HBRUSH trackBrush = CreateSolidBrush(trackColor)) {
        FillRect(dc, &outer, trackBrush);
        DeleteObject(trackBrush);
    }

    RECT inner = outer;
    InflateRect(&inner, -1, -1);
    if (inner.bottom <= inner.top || inner.right <= inner.left) {
        return;
    }

    if (item.data.progress.indeterminate) {
        const int width = inner.right - inner.left;
        if (width <= 0) {
            return;
        }
        const int segment = std::max(width / 4, 12);
        const ULONGLONG tick = GetTickCount64();
        const int cycle = width + segment;
        int offset = static_cast<int>((tick / 30) % cycle) - segment;
        RECT segmentRect{inner.left + offset, inner.top, inner.left + offset + segment, inner.bottom};
        if (segmentRect.left < inner.left) {
            segmentRect.left = inner.left;
        }
        if (segmentRect.right > inner.right) {
            segmentRect.right = inner.right;
        }
        if (segmentRect.right > segmentRect.left) {
            if (HBRUSH brush = CreateSolidBrush(m_progressEndColor)) {
                FillRect(dc, &segmentRect, brush);
                DeleteObject(brush);
            }
        }
    } else {
        const int width = inner.right - inner.left;
        if (width <= 0) {
            return;
        }
        int fill = static_cast<int>(std::round(item.data.progress.fraction * width));
        fill = std::clamp(fill, 0, width);
        if (fill > 0) {
            RECT fillRect = inner;
            fillRect.right = fillRect.left + fill;
            TRIVERTEX vertex[2];
            vertex[0].x = fillRect.left;
            vertex[0].y = fillRect.top;
            vertex[0].Red = static_cast<COLOR16>(GetRValue(m_progressStartColor) << 8);
            vertex[0].Green = static_cast<COLOR16>(GetGValue(m_progressStartColor) << 8);
            vertex[0].Blue = static_cast<COLOR16>(GetBValue(m_progressStartColor) << 8);
            vertex[0].Alpha = 0xFFFF;
            vertex[1].x = fillRect.right;
            vertex[1].y = fillRect.bottom;
            vertex[1].Red = static_cast<COLOR16>(GetRValue(m_progressEndColor) << 8);
            vertex[1].Green = static_cast<COLOR16>(GetGValue(m_progressEndColor) << 8);
            vertex[1].Blue = static_cast<COLOR16>(GetBValue(m_progressEndColor) << 8);
            vertex[1].Alpha = 0xFFFF;
            GRADIENT_RECT gradient{0, 1};
            GradientFill(dc, vertex, 2, &gradient, 1, GRADIENT_FILL_RECT_H);
        }
    }

    const COLORREF borderColor = BlendColors(trackColor, RGB(0, 0, 0), m_darkMode ? 0.5 : 0.35);
    if (HPEN pen = CreatePen(PS_SOLID, 1, borderColor)) {
        HGDIOBJ oldPen = SelectObject(dc, pen);
        MoveToEx(dc, outer.left, outer.top, nullptr);
        LineTo(dc, outer.right, outer.top);
        LineTo(dc, outer.right, outer.bottom);
        LineTo(dc, outer.left, outer.bottom);
        LineTo(dc, outer.left, outer.top);
        SelectObject(dc, oldPen);
        DeleteObject(pen);
    }
}

void TabBandWindow::DrawDropIndicator(HDC dc) const {
    const DropTarget* indicator = nullptr;
    if (m_drag.dragging && m_drag.target.active && !m_drag.target.outside && m_drag.target.indicatorX >= 0) {
        indicator = &m_drag.target;
    } else if (m_externalDrop.active && m_externalDrop.target.active && !m_externalDrop.target.outside &&
               m_externalDrop.target.indicatorX >= 0) {
        indicator = &m_externalDrop.target;
    }

    if (!indicator) {
        return;
    }

    HPEN pen = CreatePen(PS_SOLID, 2, m_accentColor);
    HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
    const int x = indicator->indicatorX;
    MoveToEx(dc, x, m_clientRect.top + 2, nullptr);
    LineTo(dc, x, m_clientRect.bottom - 2);
    SelectObject(dc, oldPen);
    DeleteObject(pen);
}

void TabBandWindow::DrawDragVisual(HDC dc) const {
    if (!m_drag.dragging || !m_drag.origin.hit || !m_drag.hasCurrent) {
        return;
    }

    if (m_drag.overlayVisible) {
        return;
    }

    const VisualItem* originItem = FindVisualForHit(m_drag.origin);
    if (!originItem) {
        return;
    }

    SIZE size{};
    HBITMAP bitmap = CreateDragVisualBitmap(*originItem, &size);
    if (!bitmap || size.cx <= 0 || size.cy <= 0) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return;
    }

    HDC memDC = CreateCompatibleDC(dc);
    if (!memDC) {
        DeleteObject(bitmap);
        return;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, bitmap);

    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 160;
    blend.AlphaFormat = 0;

    const int left = m_drag.current.x - size.cx / 2;
    const int top = m_drag.current.y - size.cy / 2;
    AlphaBlend(dc, left, top, size.cx, size.cy, memDC, 0, 0, size.cx, size.cy, blend);

    SelectObject(memDC, oldBitmap);
    DeleteDC(memDC);
    DeleteObject(bitmap);
}

HBITMAP TabBandWindow::CreateDragVisualBitmap(const VisualItem& item, SIZE* size) const {
    const int width = item.bounds.right - item.bounds.left;
    const int height = item.bounds.bottom - item.bounds.top;
    if (width <= 0 || height <= 0) {
        if (size) {
            size->cx = 0;
            size->cy = 0;
        }
        return nullptr;
    }

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap) {
        if (size) {
            size->cx = 0;
            size->cy = 0;
        }
        return nullptr;
    }

    HDC memDC = CreateCompatibleDC(nullptr);
    if (!memDC) {
        DeleteObject(bitmap);
        if (size) {
            size->cx = 0;
            size->cy = 0;
        }
        return nullptr;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, bitmap);
    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(memDC, font));
    SetBkMode(memDC, TRANSPARENT);

    VisualItem copy = item;
    copy.bounds = {0, 0, width, height};
    if (copy.data.type == TabViewItemType::kGroupHeader) {
        DrawGroupHeader(memDC, copy);
    } else {
        DrawTab(memDC, copy);
    }

    if (oldFont) {
        SelectObject(memDC, oldFont);
    }
    SelectObject(memDC, oldBitmap);
    DeleteDC(memDC);

    if (size) {
        size->cx = width;
        size->cy = height;
    }
    return bitmap;
}

void TabBandWindow::UpdateDragOverlay(const POINT& clientPt, const POINT& screenPt) {
    if (!m_drag.dragging) {
        HideDragOverlay(false);
        return;
    }

    if (PtInRect(&m_clientRect, clientPt)) {
        HideDragOverlay(false);
        return;
    }

    const VisualItem* originItem = FindVisualForHit(m_drag.origin);
    if (!originItem) {
        HideDragOverlay(false);
        return;
    }

    SIZE size{};
    HBITMAP bitmap = CreateDragVisualBitmap(*originItem, &size);
    if (!bitmap || size.cx <= 0 || size.cy <= 0) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        HideDragOverlay(false);
        return;
    }

    if (!m_drag.overlay) {
        m_drag.overlay = CreateDragOverlayWindow();
    }
    if (!m_drag.overlay) {
        DeleteObject(bitmap);
        return;
    }

    HDC screenDC = GetDC(nullptr);
    if (!screenDC) {
        DeleteObject(bitmap);
        return;
    }

    HDC memDC = CreateCompatibleDC(screenDC);
    if (!memDC) {
        ReleaseDC(nullptr, screenDC);
        DeleteObject(bitmap);
        return;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, bitmap);
    POINT position{screenPt.x - size.cx / 2, screenPt.y - size.cy / 2};
    POINT src{0, 0};
    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 160;
    blend.AlphaFormat = 0;
    UpdateLayeredWindow(m_drag.overlay, screenDC, &position, &size, memDC, &src, 0, &blend, ULW_ALPHA);

    SelectObject(memDC, oldBitmap);
    DeleteDC(memDC);
    ReleaseDC(nullptr, screenDC);
    DeleteObject(bitmap);

    ShowWindow(m_drag.overlay, SW_SHOWNOACTIVATE);
    m_drag.overlayVisible = true;
}

void TabBandWindow::HideDragOverlay(bool destroy) {
    if (m_drag.overlayVisible && m_drag.overlay) {
        ShowWindow(m_drag.overlay, SW_HIDE);
    }
    m_drag.overlayVisible = false;
    if (destroy && m_drag.overlay) {
        DestroyWindow(m_drag.overlay);
        m_drag.overlay = nullptr;
    }
}

void TabBandWindow::ClearExplorerContext() {
    m_explorerContext = {};
}

HICON TabBandWindow::LoadItemIcon(const TabViewItem& item, UINT iconFlags) const {
    if (item.type != TabViewItemType::kTab) {
        return nullptr;
    }

    SHFILEINFOW info{};
    const UINT flags = SHGFI_ICON | SHGFI_ADDOVERLAYS | iconFlags;
    if (item.pidl) {
        if (SHGetFileInfoW(reinterpret_cast<PCWSTR>(item.pidl), 0, &info, sizeof(info), flags | SHGFI_PIDL)) {
            return info.hIcon;
        }
    }
    if (!item.path.empty()) {
        if (SHGetFileInfoW(item.path.c_str(), 0, &info, sizeof(info), flags)) {
            return info.hIcon;
        }
    }
    return nullptr;
}

bool TabBandWindow::HandleExplorerMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT* result) {
    if (m_explorerContext.menu3) {
        return SUCCEEDED(m_explorerContext.menu3->HandleMenuMsg2(message, wParam, lParam, result));
    }
    if (m_explorerContext.menu2) {
        HRESULT hr = m_explorerContext.menu2->HandleMenuMsg(message, wParam, lParam);
        if (SUCCEEDED(hr)) {
            if (result) {
                *result = 0;
            }
            return true;
        }
    }
    return false;
}

void TabBandWindow::EnsureMouseTracking(const POINT& pt) {
    if (!m_hwnd) {
        return;
    }
    TRACKMOUSEEVENT tme{};
    tme.cbSize = sizeof(tme);
    tme.dwFlags = TME_LEAVE | TME_HOVER;
    tme.hwndTrack = m_hwnd;
    tme.dwHoverTime = kPreviewHoverTime;
    if (TrackMouseEvent(&tme)) {
        m_mouseTracking = true;
    }
    UpdateHoverPreview(pt);
}

void TabBandWindow::UpdateHoverPreview(const POINT& pt) {
    if (!m_previewVisible || !m_hwnd || !m_previewOverlay.IsVisible()) {
        return;
    }
    if (!PtInRect(&m_clientRect, pt) || m_previewItemIndex >= m_items.size()) {
        HidePreviewWindow(false);
        return;
    }
    HitInfo hit = HitTest(pt);
    if (!hit.hit || hit.itemIndex != m_previewItemIndex) {
        HidePreviewWindow(false);
        return;
    }
    POINT screen = pt;
    ClientToScreen(m_hwnd, &screen);
    PositionPreviewWindow(m_items[m_previewItemIndex], screen);
}

void TabBandWindow::HandleMouseHover(const POINT& pt) {
    if (!m_hwnd) {
        return;
    }
    if (!PtInRect(&m_clientRect, pt)) {
        HidePreviewWindow(false);
        return;
    }
    HitInfo hit = HitTest(pt);
    if (!hit.hit || hit.type != TabViewItemType::kTab || hit.itemIndex >= m_items.size()) {
        HidePreviewWindow(false);
        return;
    }
    POINT screen = pt;
    ClientToScreen(m_hwnd, &screen);
    ShowPreviewForItem(hit.itemIndex, screen);
}

void TabBandWindow::ShowPreviewForItem(size_t index, const POINT& screenPt) {
    CancelPreviewRequest();
    if (index >= m_items.size()) {
        HidePreviewWindow(false);
        return;
    }
    const auto& visual = m_items[index];
    if (!visual.data.pidl) {
        HidePreviewWindow(false);
        return;
    }
    if (m_owner && visual.data.location.IsValid()) {
        m_owner->EnsureTabPreview(visual.data.location);
    }
    auto preview = PreviewCache::Instance().GetPreview(visual.data.pidl, kPreviewImageSize);
    bool overlayShown = false;
    if (preview.has_value() && preview->bitmap) {
        overlayShown = m_previewOverlay.Show(m_hwnd, preview->bitmap, preview->size, screenPt);
    } else {
        std::wstring placeholderText = !visual.data.name.empty() ? visual.data.name : visual.data.tooltip;
        if (placeholderText.empty()) {
            placeholderText = !visual.data.path.empty() ? visual.data.path : L"Generating preview…";
        }
        HICON icon = LoadItemIcon(visual.data, SHGFI_LARGEICON);
        overlayShown = m_previewOverlay.Show(m_hwnd, nullptr, kPreviewImageSize, screenPt, placeholderText.c_str(), icon);
        if (icon) {
            DestroyIcon(icon);
        }
        if (overlayShown) {
            m_previewRequestId = PreviewCache::Instance().RequestPreviewAsync(visual.data.pidl, kPreviewImageSize, m_hwnd,
                                                                             WM_SHELLTABS_PREVIEW_READY);
        }
    }
    if (!overlayShown) {
        HidePreviewWindow(false);
        return;
    }
    m_previewItemIndex = index;
    m_previewVisible = true;
    m_previewAnchorPoint = screenPt;
    PositionPreviewWindow(visual, screenPt);
}

void TabBandWindow::HidePreviewWindow(bool destroy) {
    CancelPreviewRequest();
    m_previewOverlay.Hide(destroy);
    m_previewVisible = false;
    m_previewItemIndex = std::numeric_limits<size_t>::max();
    m_previewAnchorPoint = POINT{};
}

void TabBandWindow::PositionPreviewWindow(const VisualItem& item, const POINT& screenPt) {
    if (!m_previewOverlay.IsVisible()) {
        return;
    }
    RECT rect = item.bounds;
    MapWindowPoints(m_hwnd, nullptr, reinterpret_cast<POINT*>(&rect), 2);
    m_previewOverlay.PositionRelativeToRect(rect, screenPt);
}

void TabBandWindow::HandlePreviewReady(uint64_t requestId) {
    if (requestId == 0 || requestId != m_previewRequestId) {
        return;
    }
    m_previewRequestId = 0;
    if (!m_previewVisible || m_previewItemIndex >= m_items.size()) {
        return;
    }
    const auto& visual = m_items[m_previewItemIndex];
    auto preview = PreviewCache::Instance().GetPreview(visual.data.pidl, kPreviewImageSize);
    if (!preview.has_value() || !preview->bitmap) {
        return;
    }
    if (!m_previewOverlay.Show(m_hwnd, preview->bitmap, preview->size, m_previewAnchorPoint)) {
        HidePreviewWindow(false);
        return;
    }
    PositionPreviewWindow(visual, m_previewAnchorPoint);
}

void TabBandWindow::CancelPreviewRequest() {
    if (m_previewRequestId != 0) {
        PreviewCache::Instance().CancelRequest(m_previewRequestId);
        m_previewRequestId = 0;
    }
}

void TabBandWindow::RefreshProgressState() {
    auto snapshot = TabManager::Get().BuildView();
    bool layoutMismatch = snapshot.size() != m_tabData.size();
    if (!layoutMismatch) {
        for (size_t i = 0; i < snapshot.size(); ++i) {
            if (m_tabData[i].type != snapshot[i].type ||
                m_tabData[i].location.groupIndex != snapshot[i].location.groupIndex ||
                m_tabData[i].location.tabIndex != snapshot[i].location.tabIndex) {
                layoutMismatch = true;
                break;
            }
        }
    }
    if (layoutMismatch) {
        SetTabs(snapshot);
        return;
    }

    bool changed = false;
    for (size_t i = 0; i < snapshot.size(); ++i) {
        if (m_tabData[i].lastActivatedTick != snapshot[i].lastActivatedTick ||
            m_tabData[i].activationOrdinal != snapshot[i].activationOrdinal) {
            m_tabData[i].lastActivatedTick = snapshot[i].lastActivatedTick;
            m_tabData[i].activationOrdinal = snapshot[i].activationOrdinal;
            if (i < m_items.size()) {
                m_items[i].data.lastActivatedTick = snapshot[i].lastActivatedTick;
                m_items[i].data.activationOrdinal = snapshot[i].activationOrdinal;
            }
            changed = true;
        }
        if (m_tabData[i].progress != snapshot[i].progress) {
            m_tabData[i].progress = snapshot[i].progress;
            if (i < m_items.size()) {
                m_items[i].data.progress = snapshot[i].progress;
            }
            changed = true;
        }
    }
    UpdateProgressAnimationState();
    if (changed && m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }
}

void TabBandWindow::UpdateProgressAnimationState() {
    if (!m_hwnd) {
        return;
    }
    const bool active = AnyProgressActive();
    if (active) {
        if (!m_progressTimerActive) {
            if (SetTimer(m_hwnd, kProgressTimerId, 120, nullptr)) {
                m_progressTimerActive = true;
            }
        }
    } else if (m_progressTimerActive) {
        KillTimer(m_hwnd, kProgressTimerId);
        m_progressTimerActive = false;
    }
}

bool TabBandWindow::AnyProgressActive() const {
    for (const auto& item : m_tabData) {
        if (item.progress.visible) {
            return true;
        }
    }
    return false;
}

void TabBandWindow::HandleProgressTimer() {
    if (!m_hwnd) {
        return;
    }
    const ULONGLONG now = GetTickCount64();
    if (TabManager::Get().ExpireFolderOperations(now, kProgressStaleTimeoutMs)) {
        RefreshProgressState();
        return;
    }
    if (!AnyProgressActive()) {
        UpdateProgressAnimationState();
        return;
    }
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

void TabBandWindow::RegisterShellNotifications() {
    if (!m_hwnd || m_shellNotifyId != 0) {
        return;
    }
    m_shellNotifyMessage = RegisterWindowMessageW(L"ShellTabs.ShellChange");
    if (m_shellNotifyMessage == 0) {
        return;
    }
    SHChangeNotifyEntry entry{};
    entry.pidl = nullptr;
    entry.fRecursive = TRUE;
    m_shellNotifyId = SHChangeNotifyRegister(
        m_hwnd, SHCNRF_ShellLevel | SHCNRF_InterruptLevel | SHCNRF_NewDelivery, SHCNE_ALLEVENTS,
        m_shellNotifyMessage, 1, &entry);
    if (m_shellNotifyId == 0) {
        m_shellNotifyMessage = 0;
    }
}

void TabBandWindow::UnregisterShellNotifications() {
    if (m_shellNotifyId != 0) {
        SHChangeNotifyDeregister(m_shellNotifyId);
        m_shellNotifyId = 0;
    }
    m_shellNotifyMessage = 0;
}

void TabBandWindow::OnShellNotify(WPARAM wParam, LPARAM lParam) {
    struct ShellChangeNotification {
        PCIDLIST_ABSOLUTE from;
        PCIDLIST_ABSOLUTE to;
    };

    const auto* notification = reinterpret_cast<const ShellChangeNotification*>(lParam);
    if (!notification) {
        return;
    }
    const LONG eventId = static_cast<LONG>(wParam) & 0xFFFF;
    auto touch = [](PCIDLIST_ABSOLUTE pidl) {
        if (!pidl) {
            return;
        }
        if (auto parent = CloneParent(pidl)) {
            TabManager::Get().TouchFolderOperation(parent.get());
        } else {
            TabManager::Get().TouchFolderOperation(pidl);
        }
    };
    auto clear = [](PCIDLIST_ABSOLUTE pidl) {
        if (!pidl) {
            return;
        }
        if (auto parent = CloneParent(pidl)) {
            TabManager::Get().ClearFolderOperation(parent.get());
        } else {
            TabManager::Get().ClearFolderOperation(pidl);
        }
    };

    switch (eventId) {
        case SHCNE_CREATE:
        case SHCNE_DELETE:
        case SHCNE_MKDIR:
        case SHCNE_RMDIR:
        case SHCNE_RENAMEITEM:
        case SHCNE_RENAMEFOLDER:
        case SHCNE_UPDATEITEM:
            touch(notification->from);
            touch(notification->to);
            break;
        case SHCNE_UPDATEDIR:
            clear(notification->from);
            clear(notification->to);
            break;
        default:
            break;
    }
}

void TabBandWindow::UpdateCloseButtonHover(const POINT& pt) {
    size_t newIndex = kInvalidIndex;
    if (PtInRect(&m_clientRect, pt)) {
        HitInfo hit = HitTest(pt);
        if (hit.hit && hit.type == TabViewItemType::kTab && hit.itemIndex < m_items.size()) {
            const auto& item = m_items[hit.itemIndex];
            RECT closeRect = ComputeCloseButtonRect(item);
            if (closeRect.right > closeRect.left && PtInRect(&closeRect, pt)) {
                newIndex = item.index;
            }
        }
    }

    if (newIndex == m_hotCloseIndex) {
        return;
    }

    size_t previous = m_hotCloseIndex;
    m_hotCloseIndex = newIndex;

    if (!m_hwnd) {
        return;
    }

    if (previous != kInvalidIndex && previous < m_items.size()) {
        RECT invalidate = ComputeCloseButtonRect(m_items[previous]);
        InvalidateRect(m_hwnd, &invalidate, FALSE);
    }
    if (newIndex != kInvalidIndex && newIndex < m_items.size()) {
        RECT invalidate = ComputeCloseButtonRect(m_items[newIndex]);
        InvalidateRect(m_hwnd, &invalidate, FALSE);
    }
}

void TabBandWindow::ClearCloseButtonHover() {
    if (m_hotCloseIndex == kInvalidIndex) {
        return;
    }
    size_t previous = m_hotCloseIndex;
    m_hotCloseIndex = kInvalidIndex;
    if (!m_hwnd) {
        return;
    }
    if (previous < m_items.size()) {
        RECT invalidate = ComputeCloseButtonRect(m_items[previous]);
        InvalidateRect(m_hwnd, &invalidate, FALSE);
    }
}

void TabBandWindow::HandleCommand(WPARAM wParam, LPARAM) {
	if (!m_owner) {
		return;
	}

	const UINT id = LOWORD(wParam);

	if (id == IDC_NEW_TAB) {
		m_owner->OnNewTabRequested();
		return;
	}

        if (id == IDM_CREATE_SAVED_GROUP) {
                const int insertAfter = ResolveInsertGroupIndex();
                m_owner->OnCreateSavedGroup(insertAfter);
                ClearExplorerContext();
                return;
        }

        if (id >= IDM_LOAD_SAVED_GROUP_BASE && id <= IDM_LOAD_SAVED_GROUP_LAST) {
		for (const auto& entry : m_savedGroupCommands) {
			if (entry.first == id) {
				const int insertAfter = ResolveInsertGroupIndex();
				m_owner->OnLoadSavedGroup(entry.second, insertAfter);
				break;
			}
		}
                ClearExplorerContext();
                return;
        }

        if (id == IDM_NEW_THISPC_TAB) {
                if (m_owner) {
                        m_owner->OnNewTabRequested(-1);
                }
                ClearExplorerContext();
                return;
        }

        if (id == IDM_MANAGE_GROUPS) {
                if (m_owner) {
                        std::wstring focusId;
                        if (m_contextHit.location.groupIndex >= 0) {
                                focusId = m_owner->GetSavedGroupId(m_contextHit.location.groupIndex);
                        }
                        if (!focusId.empty()) {
                                m_owner->OnShowOptionsDialog(2, focusId);
                        } else {
                                m_owner->OnShowOptionsDialog(2);
                        }
                }
                ClearExplorerContext();
                return;
        }

        if (id == IDM_OPTIONS) {
                if (m_owner) {
                        m_owner->OnShowOptionsDialog(0);
                }
                ClearExplorerContext();
                return;
        }

        if (!m_contextHit.hit) {
                ClearExplorerContext();
                return;
        }

	// Handle the bulk of tab/island commands with a switch.
	switch (id) {
	case IDM_CLOSE_TAB:
		if (m_contextHit.location.IsValid()) {
			m_owner->OnCloseTabRequested(m_contextHit.location);
		}
		break;

	case IDM_HIDE_TAB:
		if (m_contextHit.location.IsValid()) {
			m_owner->OnHideTabRequested(m_contextHit.location);
		}
		break;

	case IDM_DETACH_TAB:
		if (m_contextHit.location.IsValid()) {
			m_owner->OnDetachTabRequested(m_contextHit.location);
		}
		break;

        case IDM_CLONE_TAB:
                if (m_contextHit.location.IsValid()) {
                        m_owner->OnCloneTabRequested(m_contextHit.location);
                }
                break;

        case IDM_CLOSE_OTHER_TABS:
                if (m_contextHit.location.IsValid()) {
                        m_owner->OnCloseOtherTabsRequested(m_contextHit.location);
                }
                break;

        case IDM_CLOSE_TABS_TO_RIGHT:
                if (m_contextHit.location.IsValid()) {
                        m_owner->OnCloseTabsToRightRequested(m_contextHit.location);
                }
                break;

        case IDM_CLOSE_TABS_TO_LEFT:
                if (m_contextHit.location.IsValid()) {
                        m_owner->OnCloseTabsToLeftRequested(m_contextHit.location);
                }
                break;

        case IDM_REOPEN_CLOSED_TAB:
                m_owner->OnReopenClosedTabRequested();
                break;

	case IDM_OPEN_TERMINAL:
		if (m_contextHit.location.IsValid()) {
			m_owner->OnOpenTerminal(m_contextHit.location);
		}
		break;

	case IDM_OPEN_VSCODE:
		if (m_contextHit.location.IsValid()) {
			m_owner->OnOpenVSCode(m_contextHit.location);
		}
		break;

        case IDM_COPY_PATH:
                if (m_contextHit.location.IsValid()) {
                        m_owner->OnCopyPath(m_contextHit.location);
                }
                break;

        case IDM_EDIT_GROUP:
                if (m_contextHit.location.groupIndex >= 0) {
                        m_owner->OnEditGroupProperties(m_contextHit.location.groupIndex);
                }
                break;

        case IDM_TOGGLE_ISLAND_HEADER: {
                if (m_contextHit.location.groupIndex >= 0) {
                        const bool visible = m_owner->IsGroupHeaderVisible(m_contextHit.location.groupIndex);
                        m_owner->OnSetGroupHeaderVisible(m_contextHit.location.groupIndex, !visible);
		}
		break;
	}

        case IDM_TOGGLE_ISLAND:
                m_owner->OnToggleGroupCollapsed(m_contextHit.location.groupIndex);
                break;

        case IDM_CLOSE_ISLAND:
                if (m_contextHit.location.groupIndex >= 0) {
                        m_owner->OnCloseIslandRequested(m_contextHit.location.groupIndex);
                }
                break;

        case IDM_UNHIDE_ALL:
                m_owner->OnUnhideAllInGroup(m_contextHit.location.groupIndex);
                break;

	case IDM_NEW_ISLAND:
		m_owner->OnCreateIslandAfter(m_contextHit.location.groupIndex);
		break;

        case IDM_DETACH_ISLAND:
                m_owner->OnDetachGroupRequested(m_contextHit.location.groupIndex);
                break;

        default:
                // Not handled in the switch; fall through to the hidden/explorer menu paths below.
		break;
	}

	// Handle "unhide specific tab" and Explorer context menu delegation.
	if (id >= IDM_HIDDEN_TAB_BASE) {
		for (const auto& entry : m_hiddenTabCommands) {
			if (entry.first == id) {
				m_owner->OnUnhideTabRequested(entry.second);
				break;
			}
		}
	}
	else if (m_explorerContext.menu && id >= m_explorerContext.idFirst &&
		id <= m_explorerContext.idLast) {
		m_owner->InvokeExplorerContextCommand(m_explorerContext.location,
			m_explorerContext.menu.Get(), id,
			m_explorerContext.idFirst, m_lastContextPoint);
	}

	ClearExplorerContext();
}


bool TabBandWindow::HandleMouseDown(const POINT& pt) {
    UpdateCloseButtonHover(pt);
    HitInfo hit = HitTest(pt);
    if (!hit.hit) {
        return false;
    }

    SetFocus(m_hwnd);
    HideDragOverlay(true);
    m_drag = {};
    if (hit.closeButton && hit.type == TabViewItemType::kTab) {
        m_drag.closeClick = true;
        m_drag.closeItemIndex = hit.itemIndex;
        m_drag.closeLocation = hit.location;
        if (hit.itemIndex < m_items.size()) {
            m_hotCloseIndex = m_items[hit.itemIndex].index;
        }
        if (GetCapture() != m_hwnd) {
            SetCapture(m_hwnd);
        }
        return true;
    }
    m_drag.tracking = true;
    m_drag.origin = hit;
    if (hit.itemIndex < m_items.size()) {
        const auto& item = m_items[hit.itemIndex];
        m_drag.originSelected = item.data.selected;
    } else {
        m_drag.originSelected = false;
    }
    m_drag.start = pt;
    m_drag.current = pt;
    m_drag.hasCurrent = true;
    return true;
}

bool TabBandWindow::HandleMouseUp(const POINT& pt) {
	// 1) Empty-island "+" click → open "This PC" and consume
	int groupIndex = -1;
        if (FindEmptyIslandPlusAt(pt, &groupIndex) && m_owner) {
                m_owner->OnNewTabRequested(groupIndex);
                return true; // handled; UI refresh hides the '+'
        }

	// 2) Usual UI paths
	UpdateCloseButtonHover(pt);
	bool handled = false;

	// Close button release
	if (m_drag.closeClick) {
		handled = true;
		bool inside = false;
		const TabLocation closeLocation = m_drag.closeLocation;

		if (m_drag.closeItemIndex < m_items.size()) {
			const auto& item = m_items[m_drag.closeItemIndex];
			const RECT closeRect = ComputeCloseButtonRect(item);
			if (closeRect.right > closeRect.left && PtInRect(&closeRect, pt)) {
				inside = true;
			}
		}

		CancelDrag();

		if (inside && m_owner && closeLocation.IsValid()) {
			m_owner->OnCloseTabRequested(closeLocation);
		}
		return handled; // early return is fine here
	}

	// Drop/drag end
	if (m_drag.dragging) {
		handled = true;
		m_drag.current = pt;
		m_drag.hasCurrent = true;

		POINT screen = pt;
		ClientToScreen(m_hwnd, &screen);
		UpdateExternalDrag(screen);

		TabBandWindow* targetWindow = FindWindowFromPoint(screen);
		if (!targetWindow || targetWindow == this) {
			UpdateDropTarget(pt);
		}
		else {
			m_drag.target = {};
			m_drag.target.active = true;
			m_drag.target.outside = true;
		}

		CompleteDrop();  // this typically finalizes the move
	}
	// Simple tracking release → click selection
	else if (m_drag.tracking) {
		handled = true;
		const HitInfo hit = HitTest(pt);
		if (hit.hit) {
			RequestSelection(hit);
		}
	}

	CancelDrag();
	return handled;  // <-- IMPORTANT: return what actually happened
}


bool TabBandWindow::HandleMouseMove(const POINT& pt) {
    if (!m_drag.tracking) {
        return false;
    }

    if (m_drag.closeClick) {
        return true;
    }

    bool handled = false;
    m_drag.current = pt;
    m_drag.hasCurrent = true;

    if (!m_drag.dragging) {
        if (std::abs(pt.x - m_drag.start.x) > kDragThreshold || std::abs(pt.y - m_drag.start.y) > kDragThreshold) {
            handled = true;
            m_drag.dragging = true;
            SetCapture(m_hwnd);
            auto& state = GetSharedDragState();
            std::scoped_lock lock(state.mutex);
            state.source = this;
            state.origin = m_drag.origin;
            state.screen = POINT{};
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
            state.payload.reset();
        }
    }

    if (m_drag.dragging) {
        handled = true;
        POINT screen = pt;
        ClientToScreen(m_hwnd, &screen);
        UpdateExternalDrag(screen);
        TabBandWindow* targetWindow = FindWindowFromPoint(screen);
        if (!targetWindow || targetWindow == this) {
            UpdateDropTarget(pt);
        } else {
            m_drag.target = {};
            m_drag.target.active = true;
            m_drag.target.outside = true;
            RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
        }
        UpdateDragOverlay(pt, screen);
    }

    return handled;
}

bool TabBandWindow::HandleDoubleClick(const POINT& pt) {
    if (!m_owner) {
        return false;
    }

    HitInfo hit = HitTest(pt);
    if (!hit.hit) {
        return false;
    }

    if (hit.closeButton) {
        return false;
    }

    if (hit.type == TabViewItemType::kGroupHeader) {
        m_owner->OnToggleGroupCollapsed(hit.location.groupIndex);
        return true;
    }
    if (hit.location.IsValid()) {
        m_owner->OnDetachTabRequested(hit.location);
        return true;
    }
    return false;
}

void TabBandWindow::HandleFileDrop(HDROP drop, bool ownsHandle) {
    if (!drop || !m_owner) {
        if (drop && ownsHandle) {
            DragFinish(drop);
        }
        return;
    }

    struct DropHandleCloser {
        HDROP handle;
        bool owns;
        ~DropHandleCloser() {
            if (handle && owns) {
                DragFinish(handle);
            }
        }
    } closer{drop, ownsHandle};

    POINT pt{};
    BOOL inside = DragQueryPoint(drop, &pt);
    if (!inside) {
        return;
    }

    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
    if (count == 0) {
        return;
    }

    std::vector<std::wstring> paths;
    paths.reserve(count);
    wchar_t buffer[MAX_PATH];
    for (UINT i = 0; i < count; ++i) {
        const UINT length = DragQueryFileW(drop, i, buffer, ARRAYSIZE(buffer));
        if (length == 0) {
            continue;
        }
        paths.emplace_back(buffer, buffer + length);
    }

    HitInfo hit = HitTest(pt);
    const bool dropOnTab = hit.hit && hit.type == TabViewItemType::kTab && hit.location.IsValid();
    bool handled = false;
    if (dropOnTab && !paths.empty()) {
        const bool move = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        m_owner->OnFilesDropped(hit.location, paths, move);
        handled = true;
    }

    if (!handled && !dropOnTab && !HasAnyTabs() && m_owner) {
        for (const auto& path : paths) {
            const DWORD attributes = GetFileAttributesW(path.c_str());
            if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
                continue;
            }
            m_owner->OnOpenFolderInNewTab(path);
        }
    }
}

bool TabBandWindow::HasFileDropData(IDataObject* dataObject) const {
    if (!dataObject) {
        return false;
    }

    FORMATETC format{CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    if (SUCCEEDED(dataObject->QueryGetData(&format))) {
        return true;
    }

    Microsoft::WRL::ComPtr<IShellItemArray> items;
    if (SUCCEEDED(dataObject->QueryInterface(IID_PPV_ARGS(&items))) && items) {
        DWORD count = 0;
        if (SUCCEEDED(items->GetCount(&count))) {
            return count > 0;
        }
        return true;
    }

    return false;
}

DWORD TabBandWindow::ComputeFileDropEffect(DWORD keyState, bool hasFileData) const {
    if (!hasFileData) {
        return DROPEFFECT_NONE;
    }

    if (keyState & MK_SHIFT) {
        return DROPEFFECT_MOVE;
    }
    if (keyState & MK_CONTROL) {
        return DROPEFFECT_COPY;
    }
    if (keyState & MK_ALT) {
        return DROPEFFECT_LINK;
    }
    return DROPEFFECT_COPY;
}

bool TabBandWindow::IsSameHit(const HitInfo& a, const HitInfo& b) const {
    if (a.hit != b.hit) {
        return false;
    }
    if (!a.hit) {
        return true;
    }
    if (a.type != b.type) {
        return false;
    }
    return a.location.groupIndex == b.location.groupIndex && a.location.tabIndex == b.location.tabIndex;
}

bool TabBandWindow::IsSelectedTabHit(const HitInfo& hit) const {
    if (!hit.hit || hit.type != TabViewItemType::kTab || !hit.location.IsValid()) {
        return false;
    }
    if (hit.itemIndex >= m_items.size()) {
        return false;
    }
    return m_items[hit.itemIndex].data.selected;
}

void TabBandWindow::StartDropHoverTimer() {
    if (!m_hwnd) {
        return;
    }
    CancelDropHoverTimer();
    if (SetTimer(m_hwnd, kDropHoverTimerId, 1500, nullptr)) {
        m_dropHoverTimerActive = true;
    }
}

void TabBandWindow::CancelDropHoverTimer() {
    if (!m_dropHoverTimerActive) {
        return;
    }
    if (m_hwnd) {
        KillTimer(m_hwnd, kDropHoverTimerId);
    }
    m_dropHoverTimerActive = false;
}

void TabBandWindow::UpdateDropHoverState(const HitInfo& hit, bool hasFileData) {
    const bool changed = !IsSameHit(hit, m_dropHoverHit) || hasFileData != m_dropHoverHasFileData;
    m_dropHoverHit = hit;
    m_dropHoverHasFileData = hasFileData;

    const bool eligible = hasFileData && hit.hit && hit.type == TabViewItemType::kTab && hit.location.IsValid() &&
                          !IsSelectedTabHit(hit);
    if (!eligible) {
        CancelDropHoverTimer();
        return;
    }

    if (changed || !m_dropHoverTimerActive) {
        StartDropHoverTimer();
    }
}

void TabBandWindow::ClearDropHoverState() {
    CancelDropHoverTimer();
    m_dropHoverHit = {};
    m_dropHoverHasFileData = false;
}

void TabBandWindow::OnDropHoverTimer() {
    CancelDropHoverTimer();
    if (!m_dropHoverHasFileData || !m_owner) {
        return;
    }
    if (!m_dropHoverHit.hit || m_dropHoverHit.type != TabViewItemType::kTab || !m_dropHoverHit.location.IsValid()) {
        return;
    }
    if (IsSelectedTabHit(m_dropHoverHit)) {
        return;
    }
    m_owner->OnTabSelected(m_dropHoverHit.location);
}

HRESULT TabBandWindow::OnNativeDragEnter(IDataObject* dataObject, DWORD keyState, const POINTL& point, DWORD* effect) {
    bool hasFileData = HasFileDropData(dataObject);
    if (effect) {
        *effect = ComputeFileDropEffect(keyState, hasFileData);
    }
    if (!m_hwnd) {
        return S_OK;
    }
    POINT client{static_cast<LONG>(point.x), static_cast<LONG>(point.y)};
    ScreenToClient(m_hwnd, &client);
    HitInfo hit = HitTest(client);
    UpdateDropHoverState(hit, hasFileData);
    return hasFileData ? S_OK : S_FALSE;
}

HRESULT TabBandWindow::OnNativeDragOver(DWORD keyState, const POINTL& point, DWORD* effect) {
    if (effect) {
        *effect = ComputeFileDropEffect(keyState, m_dropHoverHasFileData);
    }
    if (!m_hwnd) {
        return S_OK;
    }
    POINT client{static_cast<LONG>(point.x), static_cast<LONG>(point.y)};
    ScreenToClient(m_hwnd, &client);
    HitInfo hit = HitTest(client);
    UpdateDropHoverState(hit, m_dropHoverHasFileData);
    return m_dropHoverHasFileData ? S_OK : S_FALSE;
}

HRESULT TabBandWindow::OnNativeDragLeave() {
    ClearDropHoverState();
    return S_OK;
}

HRESULT TabBandWindow::OnNativeDrop(IDataObject* dataObject, DWORD keyState, const POINTL& point, DWORD* effect) {
    bool hasFileData = HasFileDropData(dataObject);
    if (effect) {
        *effect = ComputeFileDropEffect(keyState, hasFileData);
    }
    if (m_hwnd) {
        POINT client{static_cast<LONG>(point.x), static_cast<LONG>(point.y)};
        ScreenToClient(m_hwnd, &client);
        HitInfo hit = HitTest(client);
        UpdateDropHoverState(hit, hasFileData);
    }

    if (hasFileData && dataObject) {
        FORMATETC format{CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
        STGMEDIUM medium{};
        if (SUCCEEDED(dataObject->GetData(&format, &medium))) {
            HandleFileDrop(static_cast<HDROP>(medium.hGlobal), false);
            ReleaseStgMedium(&medium);
        }
    }

    ClearDropHoverState();
    return hasFileData ? S_OK : S_FALSE;
}

void TabBandWindow::CancelDrag() {
    // The drag bookkeeping can be reset while a transfer is still in flight -
    // for example when we rebuild the layout as part of detaching a tab to
    // another window. In that case the window may still own the mouse capture
    // even though m_drag.dragging has already been cleared. Always release the
    // capture if we own it so we do not leave Explorer with a stale preview
    // “stuck” on the band until some other window forces a repaint.
    if (GetCapture() == m_hwnd) {
        ReleaseCapture();
    }
    HideDragOverlay(true);
    {
        auto& state = GetSharedDragState();
        TabBandWindow* hovered = nullptr;
        std::scoped_lock lock(state.mutex);
        if (state.source == this) {
            hovered = state.hover;
            state.source = nullptr;
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
            state.payload.reset();
        } else if (state.hover == this) {
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
        }
        if (hovered && hovered != this) {
            DispatchExternalMessage(hovered->GetHwnd(), WM_SHELLTABS_EXTERNAL_DRAG_LEAVE);
        }
    }
    m_externalDrop = {};
    m_drag = {};
    m_mouseTracking = false;
    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
}

TabBandWindow::DropTarget TabBandWindow::ComputeDropTarget(const POINT& pt, const HitInfo& origin) const {
    DropTarget target{};
    target.active = true;

    if (pt.x < m_clientRect.left || pt.x > m_clientRect.right || pt.y < m_clientRect.top || pt.y > m_clientRect.bottom) {
        target.outside = true;
        return target;
    }

    HitInfo hit = HitTest(pt);
    if (!hit.hit) {
        if (origin.type == TabViewItemType::kTab && m_owner) {
            target.group = false;
            target.newGroup = true;
            target.floating = true;
            target.groupIndex = m_owner->GetGroupCount();
            target.tabIndex = 0;
            target.indicatorX = m_clientRect.right - 10;
        } else if (!m_items.empty()) {
            const VisualItem* lastHeader = FindLastGroupHeader();
            if (lastHeader) {
                if (origin.type == TabViewItemType::kGroupHeader) {
                    target.group = true;
                    target.groupIndex = lastHeader->data.location.groupIndex + 1;
                    target.indicatorX = lastHeader->bounds.right;
                } else {
                    target.group = false;
                    target.groupIndex = lastHeader->data.location.groupIndex;
                    target.tabIndex = static_cast<int>(lastHeader->data.totalTabs);
                    target.indicatorX = lastHeader->bounds.right;
                }
            } else {
                const auto& tail = m_items.back();
                target.group = false;
                target.groupIndex = tail.data.location.groupIndex;
                target.tabIndex = tail.data.location.tabIndex + 1;
                target.indicatorX = tail.bounds.right;
            }
        }
        return target;
    }

    const VisualItem& visual = m_items[hit.itemIndex];
    const int midX = (visual.bounds.left + visual.bounds.right) / 2;
    const bool leftSide = pt.x < midX;

    if (origin.type == TabViewItemType::kGroupHeader) {
        target.group = true;
        target.groupIndex = visual.data.location.groupIndex + (leftSide ? 0 : 1);
        target.indicatorX = leftSide ? visual.bounds.left : visual.bounds.right;
    } else {
        target.group = false;
        target.groupIndex = visual.data.location.groupIndex;
        if (visual.data.type == TabViewItemType::kTab) {
            target.tabIndex = visual.data.location.tabIndex + (leftSide ? 0 : 1);
            target.indicatorX = leftSide ? visual.bounds.left : visual.bounds.right;
        } else {
            target.tabIndex = leftSide ? 0 : static_cast<int>(visual.data.totalTabs);
            target.indicatorX = leftSide ? visual.bounds.left : visual.bounds.right;
        }
    }

    return target;
}

void TabBandWindow::UpdateDropTarget(const POINT& pt) {
    m_drag.target = ComputeDropTarget(pt, m_drag.origin);
    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::UpdateExternalDrag(const POINT& screenPt) {
    auto& state = GetSharedDragState();
    TabBandWindow* targetWindow = FindWindowFromPoint(screenPt);
    TabBandWindow* previousHover = nullptr;

    {
        std::scoped_lock lock(state.mutex);
        state.source = this;
        state.screen = screenPt;
        state.origin = m_drag.origin;
        previousHover = state.hover;
        state.targetValid = false;
        if (targetWindow == this) {
            state.hover = nullptr;
        }
    }

    if (previousHover && previousHover != targetWindow && previousHover != this) {
        DispatchExternalMessage(previousHover->GetHwnd(), WM_SHELLTABS_EXTERNAL_DRAG_LEAVE);
    }

    if (!targetWindow || targetWindow == this) {
        return;
    }

    {
        std::scoped_lock lock(state.mutex);
        if (state.source == this) {
            state.hover = targetWindow;
            state.targetValid = false;
        }
    }

    DispatchExternalMessage(targetWindow->GetHwnd(), WM_SHELLTABS_EXTERNAL_DRAG);
}

bool TabBandWindow::TryCompleteExternalDrop() {
    auto& state = GetSharedDragState();
    TabBandWindow* targetWindow = nullptr;
    DropTarget target{};

    {
        std::scoped_lock lock(state.mutex);
        if (state.source != this || !state.hover || state.hover == this || !state.targetValid) {
            return false;
        }
        targetWindow = state.hover;
        target = state.target;
    }

    if (!targetWindow || !targetWindow->m_owner || !m_owner || target.outside) {
        return false;
    }

    std::unique_ptr<TransferPayload> payload = std::make_unique<TransferPayload>();
    payload->target = targetWindow->m_owner;
    payload->targetGroupIndex = target.groupIndex;
    payload->targetTabIndex = target.tabIndex;
    payload->createGroup = target.newGroup;
    payload->headerVisible = !target.floating;
    payload->select = m_drag.originSelected;
    payload->source = m_owner;
    bool closeSourceWindow = false;

    if (m_drag.origin.type == TabViewItemType::kGroupHeader) {
        auto detachedGroup = m_owner->DetachGroupForTransfer(m_drag.origin.location.groupIndex, nullptr);
        if (!detachedGroup) {
            return false;
        }
        payload->type = TransferPayload::Type::Group;
        payload->group = std::move(*detachedGroup);
    } else if (m_drag.origin.location.IsValid()) {
        auto detachedTab = m_owner->DetachTabForTransfer(m_drag.origin.location, nullptr, false, &closeSourceWindow);
        if (!detachedTab) {
            return false;
        }
        payload->type = TransferPayload::Type::Tab;
        payload->tab = std::move(*detachedTab);
    } else {
        return false;
    }

    {
        std::scoped_lock lock(state.mutex);
        state.payload = std::move(payload);
        state.source = nullptr;
        state.hover = nullptr;
        state.targetValid = false;
        state.target = {};
    }

    DispatchExternalMessage(targetWindow->GetHwnd(), WM_SHELLTABS_EXTERNAL_DROP);
    if (closeSourceWindow && m_owner) {
        m_owner->CloseFrameWindowAsync();
    }
    return true;
}

void TabBandWindow::HandleExternalDragUpdate() {
    auto& state = GetSharedDragState();
    POINT screen{};
    TabBandWindow* sourceWindow = nullptr;
    TabBandWindow::HitInfo origin;

    {
        std::scoped_lock lock(state.mutex);
        if (state.hover != this) {
            return;
        }
        screen = state.screen;
        sourceWindow = state.source;
        origin = state.origin;
    }

    if (!sourceWindow) {
        HandleExternalDragLeave();
        return;
    }

    POINT client = screen;
    ScreenToClient(m_hwnd, &client);
    DropTarget target = ComputeDropTarget(client, origin);

    {
        std::scoped_lock lock(state.mutex);
        if (state.hover == this) {
            state.target = target;
            state.targetValid = !target.outside;
        }
    }

    if (!target.outside) {
        m_externalDrop.active = true;
        m_externalDrop.target = target;
        m_externalDrop.source = sourceWindow;
    } else {
        m_externalDrop = {};
    }

    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::HandleExternalDragLeave() {
    {
        auto& state = GetSharedDragState();
        std::scoped_lock lock(state.mutex);
        if (state.hover == this) {
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
        }
    }
    m_externalDrop = {};
    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::HandleExternalDropExecute() {
    std::unique_ptr<TransferPayload> payload;
    {
        auto& state = GetSharedDragState();
        std::scoped_lock lock(state.mutex);
        if (!state.payload || !m_owner || state.payload->target != m_owner) {
            return;
        }
        payload = std::move(state.payload);
    }

    if (!payload || !m_owner) {
        return;
    }

    if (payload->type == TransferPayload::Type::Tab) {
        m_owner->InsertTransferredTab(std::move(payload->tab), payload->targetGroupIndex, payload->targetTabIndex,
                                      payload->createGroup, payload->headerVisible, payload->select);
    } else if (payload->type == TransferPayload::Type::Group) {
        m_owner->InsertTransferredGroup(std::move(payload->group), payload->targetGroupIndex, payload->select);
    }

    m_externalDrop = {};
    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::CompleteDrop() {
    if (!m_owner || !m_drag.dragging) {
        return;
    }

    const auto origin = m_drag.origin;
    const auto target = m_drag.target;

    if (!target.active) {
        return;
    }

    if (TryCompleteExternalDrop()) {
        return;
    }

    if (target.outside) {
        if (origin.type == TabViewItemType::kGroupHeader) {
            m_owner->OnDetachGroupRequested(origin.location.groupIndex);
        } else if (origin.location.IsValid()) {
            m_owner->OnDetachTabRequested(origin.location);
        }
        return;
    }

    if (target.newGroup && origin.location.IsValid()) {
        m_owner->OnMoveTabToNewGroup(origin.location, target.groupIndex, !target.floating);
        return;
    }

    if (origin.type == TabViewItemType::kGroupHeader) {
        int fromGroup = origin.location.groupIndex;
        int toGroup = target.groupIndex;
        const int groupCount = GroupCount();
        if (toGroup < 0) {
            toGroup = 0;
        }
        if (toGroup > groupCount) {
            toGroup = groupCount;
        }
        if (fromGroup != toGroup && fromGroup + 1 != toGroup) {
            m_owner->OnMoveGroupRequested(fromGroup, toGroup);
        }
    } else if (origin.location.IsValid()) {
        TabLocation to{target.groupIndex, target.tabIndex};
        if (to.groupIndex < 0) {
            to.groupIndex = origin.location.groupIndex;
        }
        if (to.tabIndex < 0) {
            to.tabIndex = origin.location.tabIndex;
        }
        if (origin.location.groupIndex == to.groupIndex) {
            if (origin.location.tabIndex == to.tabIndex || origin.location.tabIndex + 1 == to.tabIndex) {
                return;
            }
        }
        if (to.tabIndex < 0) {
            to.tabIndex = 0;
        }
        if (to.groupIndex < 0) {
            to.groupIndex = 0;
        }
        if (!(origin.location.groupIndex == to.groupIndex && origin.location.tabIndex == to.tabIndex)) {
            m_owner->OnMoveTabRequested(origin.location, to);
        }
    }
}

void TabBandWindow::RequestSelection(const HitInfo& hit) {
    if (!m_owner) {
        return;
    }

    if (hit.type == TabViewItemType::kTab && hit.location.IsValid()) {
        m_owner->OnTabSelected(hit.location);
    } else if (hit.type == TabViewItemType::kGroupHeader) {
        m_owner->OnToggleGroupCollapsed(hit.location.groupIndex);
    }
}

TabBandWindow::HitInfo TabBandWindow::HitTest(const POINT& pt) const {
    HitInfo info;
    if (pt.x < m_clientRect.left || pt.x > m_clientRect.right || pt.y < m_clientRect.top || pt.y > m_clientRect.bottom) {
        return info;
    }

    for (size_t i = 0; i < m_items.size(); ++i) {
        const auto& item = m_items[i];
        if (PtInRect(&item.bounds, pt)) {
            info.hit = true;
            info.itemIndex = i;
            info.type = item.data.type;
            info.location = item.data.location;
            const int midX = (item.bounds.left + item.bounds.right) / 2;
            info.before = pt.x < midX;
            info.after = !info.before;
            RECT closeRect = ComputeCloseButtonRect(item);
            if (closeRect.right > closeRect.left && PtInRect(&closeRect, pt)) {
                info.closeButton = true;
            }
            return info;
        }
    }

    return info;
}

void TabBandWindow::ShowContextMenu(const POINT& screenPt) {
    if (!m_owner) {
        return;
    }

    POINT clientPt = screenPt;
    if (screenPt.x == -1 && screenPt.y == -1) {
        clientPt.x = m_clientRect.left + 10;
        clientPt.y = m_clientRect.top + 10;
        ClientToScreen(m_hwnd, &clientPt);
    }

    POINT pt = clientPt;
    ScreenToClient(m_hwnd, &pt);
    HitInfo hit = HitTest(pt);
    const VisualItem* hitVisual = FindVisualForHit(hit);
    ClearExplorerContext();
    m_savedGroupCommands.clear();
    m_contextHit = hit;
    m_lastContextPoint = clientPt;

    HMENU menu = CreatePopupMenu();
    if (!menu) {
        return;
    }

    m_hiddenTabCommands.clear();

    bool hasItemCommands = false;

    if (hit.hit) {
        if (hit.type == TabViewItemType::kTab) {
            AppendMenuW(menu, MF_STRING, IDM_CLOSE_TAB, L"Close Tab");
            AppendMenuW(menu, MF_STRING, IDM_HIDE_TAB, L"Hide Tab");
            AppendMenuW(menu, MF_STRING, IDM_DETACH_TAB, L"Move to New Window");
            AppendMenuW(menu, MF_STRING, IDM_CLONE_TAB, L"Clone Tab");

            const bool canCloseOthers = m_owner->CanCloseOtherTabs(hit.location);
            const bool canCloseRight = m_owner->CanCloseTabsToRight(hit.location);
            const bool canCloseLeft = m_owner->CanCloseTabsToLeft(hit.location);
            const bool canReopen = m_owner->CanReopenClosedTabs();

            AppendMenuW(menu, (canCloseOthers ? MF_STRING : MF_STRING | MF_GRAYED),
                        IDM_CLOSE_OTHER_TABS, L"Close Other Tabs");
            AppendMenuW(menu, (canCloseRight ? MF_STRING : MF_STRING | MF_GRAYED),
                        IDM_CLOSE_TABS_TO_RIGHT, L"Close Tabs to the Right");
            AppendMenuW(menu, (canCloseLeft ? MF_STRING : MF_STRING | MF_GRAYED),
                        IDM_CLOSE_TABS_TO_LEFT, L"Close Tabs to the Left");
            AppendMenuW(menu, (canReopen ? MF_STRING : MF_STRING | MF_GRAYED), IDM_REOPEN_CLOSED_TAB,
                        L"Reopen Closed Tab");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

            const bool headerVisible = m_owner->IsGroupHeaderVisible(hit.location.groupIndex);
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND_HEADER,
                        headerVisible ? L"Hide Island Indicator" : L"Show Island Indicator");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

            AppendMenuW(menu, MF_STRING, IDM_OPEN_TERMINAL, L"Open Terminal Here");
            AppendMenuW(menu, MF_STRING, IDM_OPEN_VSCODE, L"Open in VS Code");
            AppendMenuW(menu, MF_STRING, IDM_COPY_PATH, L"Copy Path");
            AppendMenuW(menu, MF_STRING, IDM_EDIT_GROUP, L"Edit Island...");

            HMENU explorerMenu = CreatePopupMenu();
            bool explorerInserted = false;
            if (explorerMenu) {
                Microsoft::WRL::ComPtr<IContextMenu> cmenu;
                Microsoft::WRL::ComPtr<IContextMenu2> cmenu2;
                Microsoft::WRL::ComPtr<IContextMenu3> cmenu3;
                UINT usedLast = 0;
                if (m_owner->BuildExplorerContextMenu(hit.location, explorerMenu, IDM_EXPLORER_CONTEXT_BASE,
                                                      IDM_EXPLORER_CONTEXT_LAST, &cmenu, &cmenu2, &cmenu3,
                                                      &usedLast)) {
                    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                    AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(explorerMenu), L"Explorer Context");
                    m_explorerContext.menu = std::move(cmenu);
                    m_explorerContext.menu2 = std::move(cmenu2);
                    m_explorerContext.menu3 = std::move(cmenu3);
                    m_explorerContext.idFirst = IDM_EXPLORER_CONTEXT_BASE;
                    m_explorerContext.idLast = usedLast;
                    m_explorerContext.location = hit.location;
                    explorerInserted = true;
                } else {
                    DestroyMenu(explorerMenu);
                }
            }
            if (!explorerInserted) {
                AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, L"Explorer Context");
            }

            hasItemCommands = true;
        } else if (hit.type == TabViewItemType::kGroupHeader && hit.itemIndex >= 0) {
            const auto& item = m_items[hit.itemIndex];
            const bool collapsed = item.data.collapsed;
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND, collapsed ? L"Show Island" : L"Hide Island");
            AppendMenuW(menu, MF_STRING, IDM_CLOSE_ISLAND, L"Close Island");
            const bool headerVisible = m_owner->IsGroupHeaderVisible(item.data.location.groupIndex);
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND_HEADER,
                        headerVisible ? L"Hide Island Indicator" : L"Show Island Indicator");

            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, IDM_EDIT_GROUP, L"Edit Island...");

            if (item.data.hiddenTabs > 0) {
                HMENU hiddenMenu = CreatePopupMenu();
                PopulateHiddenTabsMenu(hiddenMenu, item.data.location.groupIndex);
                AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(hiddenMenu), L"Hidden Tabs");
                AppendMenuW(menu, MF_STRING, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            } else {
                AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            }

            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, IDM_NEW_ISLAND, L"New Island After");
            AppendMenuW(menu, MF_STRING, IDM_DETACH_ISLAND, L"Move Island to New Window");
            hasItemCommands = true;
        } else if (hit.type == TabViewItemType::kGroupHeader && hit.location.groupIndex >= 0) {
            const bool headerVisible = m_owner->IsGroupHeaderVisible(hit.location.groupIndex);
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND_HEADER,
                        headerVisible ? L"Hide Island Indicator" : L"Show Island Indicator");

            bool collapsed = false;
            size_t hiddenCount = 0;
            if (hitVisual) {
                if (hitVisual->data.type == TabViewItemType::kGroupHeader) {
                    collapsed = hitVisual->data.collapsed;
                    hiddenCount = hitVisual->data.hiddenTabs;
                } else if (hitVisual->hasGroupHeader) {
                    collapsed = hitVisual->groupHeader.collapsed;
                    hiddenCount = hitVisual->groupHeader.hiddenTabs;
                }
            }

            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND,
                        collapsed ? L"Show Island" : L"Hide Island");
            AppendMenuW(menu, MF_STRING, IDM_CLOSE_ISLAND, L"Close Island");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, IDM_NEW_ISLAND, L"New Island After");
            AppendMenuW(menu, MF_STRING, IDM_DETACH_ISLAND, L"Move Island to New Window");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

            AppendMenuW(menu, MF_STRING, IDM_EDIT_GROUP, L"Edit Island...");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);


            if (hiddenCount > 0) {
                AppendMenuW(menu, MF_STRING, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            } else {
                AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            }
            hasItemCommands = true;
        }
    }

    bool appendedBeforeOptions = hasItemCommands;
    if (!hit.hit) {
        AppendMenuW(menu, MF_STRING, IDM_NEW_THISPC_TAB, L"New Tab");
        appendedBeforeOptions = true;
        hasItemCommands = true;
    }

    if (appendedBeforeOptions) {
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, IDM_MANAGE_GROUPS, L"Manage Groups...");
    AppendMenuW(menu, MF_STRING, IDM_OPTIONS, L"Options...");

    PopulateSavedGroupsMenu(menu, true);

    TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON, screenPt.x, screenPt.y, 0, m_hwnd, nullptr);
    DestroyMenu(menu);
}

void TabBandWindow::PopulateHiddenTabsMenu(HMENU menu, int groupIndex) {
    if (!menu) {
        return;
    }
    m_hiddenTabCommands.clear();

    if (!m_owner) {
        AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_HIDDEN_TAB_BASE, L"No hidden tabs");
        return;
    }

    const auto hiddenTabs = m_owner->GetHiddenTabs(groupIndex);
    if (hiddenTabs.empty()) {
        AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_HIDDEN_TAB_BASE, L"No hidden tabs");
        return;
    }

    UINT command = IDM_HIDDEN_TAB_BASE;
    for (const auto& entry : hiddenTabs) {
        AppendMenuW(menu, MF_STRING, command, entry.second.c_str());
        m_hiddenTabCommands.emplace_back(command, entry.first);
        ++command;
    }
}

void TabBandWindow::PopulateSavedGroupsMenu(HMENU parent, bool addSeparator) {
    if (!parent || !m_owner) {
        return;
    }

    HMENU groupsMenu = CreatePopupMenu();
    if (!groupsMenu) {
        return;
    }

    const auto names = m_owner->GetSavedGroupNames();
    if (names.empty()) {
        AppendMenuW(groupsMenu, MF_STRING | MF_GRAYED, 0, L"No Saved Groups");
    } else {
        UINT command = IDM_LOAD_SAVED_GROUP_BASE;
        for (const auto& name : names) {
            if (command > IDM_LOAD_SAVED_GROUP_LAST) {
                break;
            }
            AppendMenuW(groupsMenu, MF_STRING, command, name.c_str());
            m_savedGroupCommands.emplace_back(command, name);
            ++command;
        }
    }

    if (addSeparator) {
        AppendMenuW(parent, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(parent, MF_POPUP, reinterpret_cast<UINT_PTR>(groupsMenu), L"Groups");
    AppendMenuW(parent, MF_STRING, IDM_CREATE_SAVED_GROUP, L"Create Saved Group...");

}

bool TabBandWindow::HasAnyTabs() const {
    for (const auto& item : m_tabData) {
        if (item.type == TabViewItemType::kTab) {
            return true;
        }
    }
    return false;
}

int TabBandWindow::ResolveInsertGroupIndex() const {
    if (!m_owner) {
        return -1;
    }
    if (m_contextHit.hit && m_contextHit.location.groupIndex >= 0) {
        return m_contextHit.location.groupIndex;
    }
    return m_owner->GetGroupCount() - 1;
}

int TabBandWindow::GroupCount() const {
    int count = 0;
    int lastGroup = std::numeric_limits<int>::min();
    for (const auto& item : m_tabData) {
        if (item.location.groupIndex < 0) {
            continue;
        }
        if (item.location.groupIndex != lastGroup) {
            ++count;
            lastGroup = item.location.groupIndex;
        }
    }
    return count;
}

const TabBandWindow::VisualItem* TabBandWindow::FindLastGroupHeader() const {
    for (auto it = m_items.rbegin(); it != m_items.rend(); ++it) {
        if (it->data.type == TabViewItemType::kGroupHeader) {
            return &(*it);
        }
    }
    return nullptr;
}

const TabBandWindow::VisualItem* TabBandWindow::FindVisualForHit(const HitInfo& hit) const {
    if (!hit.hit) {
        return nullptr;
    }

    for (const auto& item : m_items) {
        if (item.data.type != hit.type) {
            continue;
        }
        if (hit.type == TabViewItemType::kGroupHeader) {
            if (item.data.location.groupIndex == hit.location.groupIndex) {
                return &item;
            }
        } else if (hit.location.IsValid() && item.data.location.IsValid()) {
            if (item.data.location.groupIndex == hit.location.groupIndex &&
                item.data.location.tabIndex == hit.location.tabIndex) {
                return &item;
            }
        }
    }

    if (hit.type == TabViewItemType::kGroupHeader) {
        for (const auto& item : m_items) {
            if (item.data.type == TabViewItemType::kTab && item.indicatorHandle &&
                item.data.location.groupIndex == hit.location.groupIndex) {
                return &item;
            }
        }
    }
    return nullptr;
}
LRESULT CALLBACK TabBandWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    TabBandWindow* self = reinterpret_cast<TabBandWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (message == WM_NCCREATE) {
        auto create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
        self = static_cast<TabBandWindow*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        if (self) {
            self->m_hwnd = hwnd;
        }
    }

    auto fallback = [&]() -> LRESULT { return DefWindowProcW(hwnd, message, wParam, lParam); };

    if (!self) {
        return fallback();
    }

    auto dispatch = [&]() -> LRESULT {
        const UINT optionsChangedMessage = GetOptionsChangedMessage();
        if (optionsChangedMessage != 0 && message == optionsChangedMessage) {
            self->RefreshTheme();
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        const UINT progressMessage = GetProgressUpdateMessage();
        if (progressMessage != 0 && message == progressMessage) {
            self->RefreshProgressState();
            return 0;
        }
        if (self->m_shellNotifyMessage != 0 && message == self->m_shellNotifyMessage) {
            self->OnShellNotify(wParam, lParam);
            return 0;
        }
        switch (message) {
            case WM_CREATE: {
                self->m_newTabButton = CreateWindowExW(0, L"BUTTON", L"+",
                                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON | BS_CENTER |
                                                           BS_VCENTER | BS_FLAT,
                                                       0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDC_NEW_TAB),
                                                       GetModuleHandleInstance(), nullptr);
                if (self->m_newTabButton) {
                    //SetWindowTheme(self->m_newTabButton, L"Explorer", nullptr);
                    ApplyImmersiveDarkMode(self->m_newTabButton, self->m_darkMode);
                    SendMessageW(self->m_newTabButton, WM_SETFONT,
                                 reinterpret_cast<WPARAM>(GetDefaultFont()), FALSE);
                }
                self->RefreshTheme();
                DragAcceptFiles(hwnd, TRUE);
                return 0;
            }
            case WM_SIZE: {
                const int width = LOWORD(lParam);
                const int height = HIWORD(lParam);
                self->EnsureRebarIntegration();
                self->Layout(width, height);
                return 0;
            }

            case WM_WINDOWPOSCHANGED: {
                self->EnsureRebarIntegration();
                return fallback();
            }
            case WM_DRAWITEM: {
                LRESULT handled = 0;
                if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                    return handled;
                }
                return fallback();
            }
            case WM_INITMENUPOPUP:
            case WM_MEASUREITEM: {
                LRESULT handled = 0;
                if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                    return handled;
                }
                return fallback();
            }
            case WM_MENUCHAR: {
                LRESULT handled = 0;
                if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                    return handled;
                }
                return fallback();
            }
            case WM_COMMAND: {
                self->HandleCommand(wParam, lParam);
                return 0;
            }
            case WM_LBUTTONDOWN: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                if (self->HandleMouseDown(pt)) {
                    return 0;
                }
                return fallback();
            }
            case WM_LBUTTONUP: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                if (self->HandleMouseUp(pt)) {
                    return 0;
                }
                return fallback();
            }
            case WM_MOUSEMOVE: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                self->EnsureMouseTracking(pt);
                self->UpdateCloseButtonHover(pt);
                if (self->HandleMouseMove(pt)) {
                    return 0;
                }
                return fallback();
            }
            case WM_MOUSEHOVER: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                self->HandleMouseHover(pt);
                return 0;
            }
            case WM_NCHITTEST: {
                POINT screen{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                POINT client = screen;
                ScreenToClient(hwnd, &client);
                if (client.x >= 0 && client.y >= 0 && client.x < self->m_toolbarGripWidth) {
                    return HTTRANSPARENT;
                }
                return HTCLIENT;
            }
            case WM_MOUSELEAVE: {
                self->m_mouseTracking = false;
                self->HidePreviewWindow(false);
                self->ClearCloseButtonHover();
                return 0;
            }
            case WM_RBUTTONUP: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                ClientToScreen(hwnd, &pt);
                self->ShowContextMenu(pt);
                return 0;
            }
            case WM_LBUTTONDBLCLK: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                if (self->HandleDoubleClick(pt)) {
                    return 0;
                }
                return fallback();
            }
            case WM_DROPFILES: {
                self->HandleFileDrop(reinterpret_cast<HDROP>(wParam), true);
                return 0;
            }
            case WM_TIMER: {
                if (wParam == TabBandWindow::kDropHoverTimerId) {
                    self->OnDropHoverTimer();
                    return 0;
                }
                if (wParam == TabBandWindow::kProgressTimerId) {
                    self->HandleProgressTimer();
                    return 0;
                }
                return fallback();
            }
            //case WM_THEMECHANGED:
            case WM_SETTINGCHANGE:
            case WM_SYSCOLORCHANGE: {
                self->RefreshTheme();
                InvalidateRect(hwnd, nullptr, TRUE);
                return 0;
            }
            case WM_CONTEXTMENU: {
                POINT screenPt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                if (screenPt.x == -1 && screenPt.y == -1) {
                    screenPt.x = 0;
                    screenPt.y = 0;
                    ClientToScreen(hwnd, &screenPt);
                }
                self->ShowContextMenu(screenPt);
                return 0;
            }
            case WM_SHELLTABS_DEFER_NAVIGATE: {
                if (self->m_owner) {
                    self->m_owner->OnDeferredNavigate();
                }
                return 0;
            }
            case WM_SHELLTABS_OPEN_FOLDER: {
                const auto* payload = reinterpret_cast<const OpenFolderMessagePayload*>(wParam);
                if (self->m_owner && payload && payload->path && payload->length > 0) {
                    std::wstring path(payload->path, payload->length);
                    self->m_owner->OnOpenFolderInNewTab(path);
                }
                return 0;
            }
            case WM_SHELLTABS_EXTERNAL_DRAG: {
                self->HandleExternalDragUpdate();
                return 0;
            }
            case WM_SHELLTABS_EXTERNAL_DRAG_LEAVE: {
                self->HandleExternalDragLeave();
                return 0;
            }
            case WM_SHELLTABS_EXTERNAL_DROP: {
                self->HandleExternalDropExecute();
                return 0;
            }
            case WM_SHELLTABS_PREVIEW_READY: {
                self->HandlePreviewReady(static_cast<uint64_t>(wParam));
                return 0;
            }
            case WM_COPYDATA: {
                auto* data = reinterpret_cast<const COPYDATASTRUCT*>(lParam);
                if (!data || data->dwData != SHELLTABS_COPYDATA_OPEN_FOLDER || data->cbData == 0 ||
                    !data->lpData) {
                    return fallback();
                }

                if (!self->m_owner) {
                    return TRUE;
                }

                const wchar_t* buffer = static_cast<const wchar_t*>(data->lpData);
                const size_t charCount = data->cbData / sizeof(wchar_t);
                if (charCount == 0) {
                    return TRUE;
                }

                std::wstring path(buffer, buffer + charCount);
                if (!path.empty() && path.back() == L'\0') {
                    path.pop_back();
                }

                if (!path.empty()) {
                    self->m_owner->OnOpenFolderInNewTab(path);
                }
                return TRUE;
            }
            case WM_PAINT: {
                PAINTSTRUCT ps;
                HDC dc = BeginPaint(hwnd, &ps);
                self->Draw(dc);
                EndPaint(hwnd, &ps);
                return 0;
            }
            case WM_ERASEBKGND: {
                return 1;
            }
            case WM_CAPTURECHANGED: {
                self->CancelDrag();
                return fallback();
            }
            case WM_DESTROY: {
                DragAcceptFiles(hwnd, FALSE);
                self->ClearExplorerContext();
                self->ClearVisualItems();
                self->CloseThemeHandles();
                self->ClearDropHoverState();
                self->HidePreviewWindow(true);
                self->UnregisterShellNotifications();
                TabManager::Get().UnregisterProgressListener(hwnd);
                if (self->m_progressTimerActive) {
                    KillTimer(hwnd, TabBandWindow::kProgressTimerId);
                    self->m_progressTimerActive = false;
                }
                if (self->m_dropTarget) {
                    RevokeDragDrop(hwnd);
                    self->m_dropTarget.Reset();
                }
                self->m_parentRebar = nullptr;
                self->m_rebarBandIndex = -1;
                UnregisterWindow(hwnd, self);
                self->m_hwnd = nullptr;
                self->m_newTabButton = nullptr;
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                return fallback();
            }
            default:
                return fallback();
        }
    };

    return GuardExplorerCall(L"TabBandWindow::WndProc", dispatch, fallback);
}

}  // namespace shelltabs
