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
#include <list>
#include <iterator>

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
#include "ExplorerThemeUtils.h"
#include "ThemeHooks.h"

#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
using Microsoft::WRL::ComPtr;

namespace shelltabs {

LRESULT CALLBACK NewTabButtonWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

namespace {

constexpr wchar_t kNewTabButtonClassName[] = L"ShellTabsNewTabButton";

bool EnsureNewTabButtonClassRegistered() {
    static bool attempted = false;
    static bool succeeded = false;
    if (!attempted) {
        attempted = true;
        WNDCLASSW wc{};
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = NewTabButtonWndProc;
        wc.hInstance = GetModuleHandleInstance();
        wc.lpszClassName = kNewTabButtonClassName;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        SetLastError(ERROR_SUCCESS);
        const ATOM atom = RegisterClassW(&wc);
        if (atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
            succeeded = true;
        }
    }
    return succeeded;
}

struct FontMetricsKey {
    LONG height = 0;
    LONG aveCharWidth = 0;
    LONG weight = 0;
    BYTE italic = 0;
    BYTE pitchAndFamily = 0;
    BYTE charSet = 0;

    bool operator==(const FontMetricsKey& other) const noexcept {
        return height == other.height && aveCharWidth == other.aveCharWidth && weight == other.weight &&
               italic == other.italic && pitchAndFamily == other.pitchAndFamily && charSet == other.charSet;
    }
};

struct TextWidthCacheKey {
    std::wstring text;
    FontMetricsKey metrics;

    bool operator==(const TextWidthCacheKey& other) const noexcept {
        return metrics == other.metrics && text == other.text;
    }
};

struct TextWidthCacheKeyHasher {
    size_t operator()(const TextWidthCacheKey& key) const noexcept {
        size_t hash = std::hash<std::wstring>{}(key.text);
        hash ^= static_cast<size_t>(key.metrics.height) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
        hash ^= static_cast<size_t>(key.metrics.aveCharWidth) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
        hash ^= static_cast<size_t>(key.metrics.weight) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
        hash ^= static_cast<size_t>(key.metrics.italic) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
        hash ^= static_cast<size_t>(key.metrics.pitchAndFamily) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
        hash ^= static_cast<size_t>(key.metrics.charSet) + 0x9e3779b97f4a7c15ULL + (hash << 6) + (hash >> 2);
        return hash;
    }
};

class TextWidthCache {
public:
    explicit TextWidthCache(size_t capacity) : m_capacity(std::max<size_t>(1, capacity)) {}

    bool TryGet(const std::wstring& text, const FontMetricsKey& metrics, int& width) {
        TextWidthCacheKey lookupKey{ text, metrics };
        auto it = m_lookup.find(lookupKey);
        if (it == m_lookup.end()) {
            return false;
        }

        m_entries.splice(m_entries.begin(), m_entries, it->second);
        width = it->second->width;
        return true;
    }

    void Put(const std::wstring& text, const FontMetricsKey& metrics, int width) {
        TextWidthCacheKey key{ text, metrics };
        auto it = m_lookup.find(key);
        if (it != m_lookup.end()) {
            it->second->width = width;
            m_entries.splice(m_entries.begin(), m_entries, it->second);
            return;
        }

        m_entries.emplace_front();
        Entry& entry = m_entries.front();
        entry.key = std::move(key);
        entry.width = width;
        m_lookup.emplace(entry.key, m_entries.begin());

        if (m_entries.size() > m_capacity) {
            auto last = std::prev(m_entries.end());
            m_lookup.erase(last->key);
            m_entries.pop_back();
        }
    }

    void Clear() {
        m_entries.clear();
        m_lookup.clear();
    }

private:
    struct Entry {
        TextWidthCacheKey key;
        int width = 0;
    };

    size_t m_capacity;
    std::list<Entry> m_entries;
    std::unordered_map<TextWidthCacheKey, std::list<Entry>::iterator, TextWidthCacheKeyHasher> m_lookup;
};

TextWidthCache& GetTextWidthCache() {
    static TextWidthCache cache(128);
    return cache;
}

void ClearTextWidthCache() {
    GetTextWidthCache().Clear();
}

std::atomic<uint32_t> g_availableDockMask{0};
std::mutex g_availableDockMaskMutex;
std::unordered_map<HWND, uint32_t> g_availableDockMaskByFrame;

class SelectObjectGuard {
public:
    SelectObjectGuard(HDC dc, HGDIOBJ object) noexcept : m_dc(dc) {
        if (m_dc && object) {
            HGDIOBJ previous = SelectObject(m_dc, object);
            if (previous && previous != HGDI_ERROR) {
                m_previous = previous;
            }
        }
    }

    SelectObjectGuard(const SelectObjectGuard&) = delete;
    SelectObjectGuard& operator=(const SelectObjectGuard&) = delete;
    SelectObjectGuard(SelectObjectGuard&&) = delete;
    SelectObjectGuard& operator=(SelectObjectGuard&&) = delete;

    ~SelectObjectGuard() {
        if (m_dc && m_previous) {
            SelectObject(m_dc, m_previous);
        }
    }

private:
    HDC m_dc = nullptr;
    HGDIOBJ m_previous = nullptr;
};

void RecomputeAvailableDockMaskLocked() {
    uint32_t combined = 0;
    for (const auto& entry : g_availableDockMaskByFrame) {
        combined |= entry.second;
    }
    g_availableDockMask.store(combined, std::memory_order_release);
}

void StoreAvailableDockMaskForFrame(HWND frame, uint32_t mask) {
    if (!frame || mask == 0) {
        return;
    }

    std::scoped_lock lock(g_availableDockMaskMutex);
    g_availableDockMaskByFrame[frame] = mask;
    RecomputeAvailableDockMaskLocked();
}

void ClearAvailableDockMaskForFrame(HWND frame) {
    if (!frame) {
        return;
    }

    std::scoped_lock lock(g_availableDockMaskMutex);
    if (g_availableDockMaskByFrame.erase(frame) > 0) {
        RecomputeAvailableDockMaskLocked();
    }
}

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
           a.isSavedGroup == b.isSavedGroup && a.pinned == b.pinned && a.progress == b.progress;
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

    StoreAvailableDockMaskForFrame(frame, mask);
}
// Older Windows SDKs used by consumers of the project might not expose the
// SID_SDataObject symbol (the service identifier for the current data object).
// Define the GUID locally so the build remains compatible with those SDKs.
constexpr GUID kSidDataObject = {0x000214e8, 0x0000, 0x0000,
                                 {0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

const wchar_t kWindowClassName[] = L"ShellTabsBandWindow";
constexpr size_t kInvalidIndex = std::numeric_limits<size_t>::max();
constexpr int kButtonWidth = 24;
constexpr int kButtonHeight = 24;
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
constexpr int kDropIndicatorHalfWidth = 3;
constexpr int kDropInvalidatePadding = 2;
// Small placeholder for empty island content
constexpr int kEmptyIslandBodyMinWidth = 24; // enough space for a centered "+"
constexpr int kEmptyIslandBodyMaxWidth = 32; // clamp empty outline length
constexpr int kEmptyPlusSize = 14; // glyph size
constexpr int kPinnedGlyphWidth = 12;
constexpr int kPinnedGlyphPadding = 6;
constexpr int kPinnedTabMaxWidth = 160;

constexpr UINT WM_SHELLTABS_EXTERNAL_DRAG = WM_APP + 60;
constexpr UINT WM_SHELLTABS_EXTERNAL_DRAG_LEAVE = WM_APP + 61;
constexpr UINT WM_SHELLTABS_EXTERNAL_DROP = WM_APP + 62;
constexpr UINT WM_SHELLTABS_THEME_CHANGED = WM_APP + 80;
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

bool IsHighContrastActive() {
    HIGHCONTRASTW info{sizeof(info)};
    if (!SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(info), &info, FALSE)) {
        return false;
    }
    return (info.dwFlags & HCF_HIGHCONTRASTON) != 0;
}

bool RectHasArea(const RECT& rect) {
    return rect.right > rect.left && rect.bottom > rect.top;
}

struct HostChromeSample {
    COLORREF top = GetSysColor(COLOR_BTNFACE);
    COLORREF bottom = GetSysColor(COLOR_BTNFACE);
    bool valid = false;
};

HostChromeSample SampleHostChrome(HWND host, const RECT& windowRect) {
    HostChromeSample sample;
    if (!host) {
        return sample;
    }

    RECT localRect = windowRect;
    if (localRect.left >= localRect.right || localRect.top >= localRect.bottom) {
        return sample;
    }

    MapWindowPoints(nullptr, host, reinterpret_cast<POINT*>(&localRect), 2);

    HDC dc = GetWindowDC(host);
    if (!dc) {
        return sample;
    }

    const LONG height = localRect.bottom - localRect.top;
    const LONG width = localRect.right - localRect.left;
    if (height <= 0 || width <= 0) {
        ReleaseDC(host, dc);
        return sample;
    }

    const LONG sampleHeight = std::max<LONG>(4, std::min<LONG>(height / 3, 24));

    RECT topRect = localRect;
    topRect.bottom = std::min<LONG>(topRect.top + sampleHeight, localRect.bottom);

    RECT bottomRect = localRect;
    bottomRect.top = std::max<LONG>(bottomRect.bottom - sampleHeight, localRect.top);

    const auto topSample = SampleAverageColor(dc, topRect);
    const auto bottomSample = SampleAverageColor(dc, bottomRect);
    if (topSample) {
        sample.top = *topSample;
    }
    if (bottomSample) {
        sample.bottom = *bottomSample;
    }
    if (!bottomSample && topSample) {
        sample.bottom = sample.top;
    } else if (!topSample && bottomSample) {
        sample.top = sample.bottom;
    }
    sample.valid = topSample.has_value() || bottomSample.has_value();

    ReleaseDC(host, dc);
    return sample;
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

GlowColorSet TabBandWindow::BuildRebarGlowColors(const ThemePalette& palette) const {
    GlowColorSet colors{};
    const COLORREF top = palette.rebarGradientValid ? palette.rebarGradientTop : palette.rebarBackground;
    const COLORREF bottom = palette.rebarGradientValid ? palette.rebarGradientBottom : palette.rebarBackground;
    colors.start = top;
    colors.end = bottom;
    colors.gradient = palette.rebarGradientValid && top != bottom;
    colors.valid = true;
    if (!colors.gradient) {
        colors.end = colors.start;
    }
    return colors;
}

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
    InvalidateRebarIntegration();
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
        InvalidateRebarIntegration();
        EnsureRebarIntegration();
        if (!m_dropTarget) {
            m_dropTarget.Attach(new BandDropTarget(this));
        }
        if (m_dropTarget) {
            EnsureDropTargetRegistered();
        }
        RegisterShellNotifications();
        if (auto* manager = ResolveManager()) {
            manager->RegisterProgressListener(m_hwnd);
        }
        UpdateProgressAnimationState();
        m_themeNotifier.Initialize(m_hwnd, [hwnd = m_hwnd]() {
            if (IsWindow(hwnd)) {
                PostMessageW(hwnd, WM_SHELLTABS_THEME_CHANGED, 0, 0);
            }
        });
    }

    return m_hwnd;
}

void TabBandWindow::EnsureDropTargetRegistered() {
    if (!m_hwnd || !m_dropTarget) {
        return;
    }
    if (m_dropTargetRegistered) {
        return;
    }

    const HRESULT hr = RegisterDragDrop(m_hwnd, m_dropTarget.Get());
    if (SUCCEEDED(hr) || hr == DRAGDROP_E_ALREADYREGISTERED) {
        m_dropTargetRegistered = true;
        m_dropTargetRegistrationPending = false;
        return;
    }

    if (hr == CO_E_NOTINITIALIZED) {
        ScheduleDropTargetRegistrationRetry();
    }
}

void TabBandWindow::ScheduleDropTargetRegistrationRetry() {
    if (!m_hwnd || m_dropTargetRegistrationPending) {
        return;
    }

    if (PostMessageW(m_hwnd, WM_SHELLTABS_REGISTER_DRAGDROP, 0, 0)) {
        m_dropTargetRegistrationPending = true;
    }
}

void TabBandWindow::Destroy() {
    CancelDrag();
    ClearExplorerContext();
    ClearVisualItems();
    CloseThemeHandles();
    ClearGdiCache();
    ClearDropHoverState();
    HidePreviewWindow(true);
    if (m_hwnd) {
        if (auto* manager = ResolveManager()) {
            manager->UnregisterProgressListener(m_hwnd);
        }
    }
    UnregisterShellNotifications();
    if (m_hwnd && m_progressTimerActive) {
        KillTimer(m_hwnd, kProgressTimerId);
        m_progressTimerActive = false;
    }
    if (m_hwnd && m_dropTargetRegistered) {
        RevokeDragDrop(m_hwnd);
        m_dropTargetRegistered = false;
    }
    m_dropTarget.Reset();
    m_dropTargetRegistrationPending = false;
    m_themeNotifier.Shutdown();
    m_darkMode = false;
    m_refreshingTheme = false;
    m_windowDarkModeInitialized = false;
    m_windowDarkModeValue = false;
    m_newTabButtonHot = false;
    m_newTabButtonPressed = false;
    m_newTabButtonKeyboardPressed = false;
    m_newTabButtonTrackingMouse = false;
    m_newTabButtonPointerPressed = false;
    m_newTabButtonCommandPending = false;
    ResetThemePalette();
    ReleaseBackBuffer();

    if (m_newTabButton) {
        DestroyWindow(m_newTabButton);
        m_newTabButton = nullptr;
    }
    if (m_hwnd) {
        UnregisterWindow(m_hwnd, this);
        DestroyWindow(m_hwnd);
        m_hwnd = nullptr;
    }
    if (m_parentFrame) {
        ClearAvailableDockMaskForFrame(m_parentFrame);
        m_parentFrame = nullptr;
    }
    m_parentRebar = nullptr;
    m_rebarBandIndex = -1;
    InvalidateRebarIntegration();
    m_tabData.clear();
    m_activeProgressCount = 0;
    m_tabLocationIndex.clear();
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

void TabBandWindow::RebuildTabLocationIndex() {
    m_tabLocationIndex.clear();
    if (m_tabData.empty()) {
        return;
    }

    m_tabLocationIndex.reserve(m_tabData.size());
    for (size_t i = 0; i < m_tabData.size(); ++i) {
        const auto& item = m_tabData[i];
        if (item.type != TabViewItemType::kTab) {
            continue;
        }
        if (!item.location.IsValid()) {
            continue;
        }
        m_tabLocationIndex[item.location] = i;
    }
}

void TabBandWindow::SetTabs(const std::vector<TabViewItem>& items) {
    m_tabData = items;
    if (auto* manager = ResolveManager(); manager) {
        m_tabLayoutVersion = manager->GetLayoutVersion();
    } else {
        m_tabLayoutVersion = 0;
    }
    RecomputeActiveProgressCount();
    RebuildTabLocationIndex();
    m_contextHit = {};
    ClearExplorerContext();

    if (!m_hwnd) {
        DestroyVisualItemResources(m_items);
        m_items.clear();
        m_progressRects.clear();
        m_activeProgressIndices.clear();
        m_emptyIslandPlusButtons.clear();
        SetRectEmpty(&m_newTabBounds);
        m_nextRedrawIncremental = false;
        m_lastAppliedRowCount = 0;
        InvalidateGroupOutlineCache();
        return;
    }

    std::vector<VisualItem> oldItems;
    oldItems.swap(m_items);

    VisualItemReuseContext reuseContext;
    VisualItemReuseContext* reuseContextPtr = nullptr;
    if (!oldItems.empty()) {
        reuseContext.source = &oldItems;
        reuseContext.reserved.assign(oldItems.size(), false);
        reuseContext.indexByKey.reserve(oldItems.size());
        for (size_t i = 0; i < oldItems.size(); ++i) {
            const uint64_t key = oldItems[i].stableId != 0
                                     ? oldItems[i].stableId
                                     : ComputeTabViewStableId(oldItems[i].data);
            reuseContext.indexByKey[key].push_back(i);
        }
        reuseContextPtr = &reuseContext;
    }

    HideDragOverlay(true);
    HidePreviewWindow(false);
    DropTarget clearedDropTarget{};
    ApplyInternalDropTarget(m_drag.target, clearedDropTarget);
    m_drag = {};
    m_emptyIslandPlusButtons.clear();

    LayoutResult layout = BuildLayoutItems(items, reuseContextPtr);

    LayoutDiffStats diff = ComputeLayoutDiff(oldItems, layout.items);
    if (reuseContextPtr) {
        ApplyPreservedVisualItems(oldItems, layout.items, diff);
    }
    const int normalizedRowCount = layout.rowCount > 0 ? layout.rowCount : std::max(m_lastRowCount, 1);
    const bool rowCountChanged = normalizedRowCount != m_lastRowCount;

    m_items = std::move(layout.items);

    auto isValidRect = [](const RECT& rect) {
        return rect.right > rect.left && rect.bottom > rect.top;
    };

    if (layout.newTabVisible && isValidRect(layout.newTabBounds)) {
        m_newTabBounds = layout.newTabBounds;
        if (m_newTabButton) {
            const int width = m_newTabBounds.right - m_newTabBounds.left;
            const int height = m_newTabBounds.bottom - m_newTabBounds.top;
            MoveWindow(m_newTabButton, m_newTabBounds.left, m_newTabBounds.top, width, height, TRUE);
            ShowWindow(m_newTabButton, SW_SHOW);
        }
    } else {
        SetRectEmpty(&m_newTabBounds);
        if (m_newTabButton) {
            ShowWindow(m_newTabButton, SW_HIDE);
        }
    }

    RebuildProgressRectCache();
    RebuildGroupOutlineCache();
    m_lastRowCount = normalizedRowCount;

    if (!diff.removedIndices.empty()) {
        std::vector<VisualItem> removed;
        removed.reserve(diff.removedIndices.size());
        for (size_t index : diff.removedIndices) {
            if (index < oldItems.size()) {
                removed.emplace_back(std::move(oldItems[index]));
            }
        }
        if (!removed.empty()) {
            DestroyVisualItemResources(removed);
        }
    }

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
    RebuildLayout();
}

void TabBandWindow::DestroyVisualItemResources(std::vector<VisualItem>& items) {
    for (auto& item : items) {
        item.icon.Reset();
    }
}

void TabBandWindow::ClearVisualItems() {
    HideDragOverlay(true);
    HidePreviewWindow(false);

    DestroyVisualItemResources(m_items);

    m_items.clear();
    m_progressRects.clear();
    m_activeProgressIndices.clear();
    DropTarget clearedDropTarget{};
    ApplyInternalDropTarget(m_drag.target, clearedDropTarget);
    m_drag = {};
    m_contextHit = {};
    m_emptyIslandPlusButtons.clear();
    SetRectEmpty(&m_newTabBounds);
    InvalidateGroupOutlineCache();
}

void TabBandWindow::ReleaseBackBuffer() {
    if (m_backBufferDC && m_backBufferOldBitmap && m_backBufferOldBitmap != HGDI_ERROR) {
        SelectObject(m_backBufferDC, m_backBufferOldBitmap);
    }
    m_backBufferOldBitmap = nullptr;

    if (m_backBufferBitmap) {
        DeleteObject(m_backBufferBitmap);
        m_backBufferBitmap = nullptr;
    }

    if (m_backBufferDC) {
        DeleteDC(m_backBufferDC);
        m_backBufferDC = nullptr;
    }

    m_backBufferSize = {0, 0};
}

TabBandWindow::LayoutResult TabBandWindow::BuildLayoutItems(const std::vector<TabViewItem>& items,
                                                            VisualItemReuseContext* reuseContext) {
    LayoutResult result;
    if (!m_hwnd) {
        return result;
    }

    m_emptyIslandPlusButtons.clear();

    RECT bounds = m_clientRect;
    const int boundsLeft = static_cast<int>(bounds.left);
    const int boundsRight = static_cast<int>(bounds.right);
    const int boundsTop = static_cast<int>(bounds.top);
    const int boundsBottom = static_cast<int>(bounds.bottom);
    const int availableWidth = boundsRight - boundsLeft;
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

    FontMetricsKey metricsKey;
    metricsKey.height = tm.tmHeight;
    metricsKey.aveCharWidth = tm.tmAveCharWidth;
    metricsKey.weight = tm.tmWeight;
    metricsKey.italic = tm.tmItalic;
    metricsKey.pitchAndFamily = tm.tmPitchAndFamily;
    metricsKey.charSet = tm.tmCharSet;

    auto& widthCache = GetTextWidthCache();

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

    const int bandWidth = boundsRight - boundsLeft;
    const int gripWidth = std::clamp(m_toolbarGripWidth, 0, std::max(0, bandWidth));

    const int boundsHeight = boundsBottom - boundsTop;
    int buttonHeight = std::max(0, boundsHeight - kButtonMargin * 2);
    if (buttonHeight > kButtonHeight) {
        buttonHeight = kButtonHeight;
    }
    if (buttonHeight == 0 && boundsHeight > 0) {
        buttonHeight = std::min(boundsHeight, kButtonHeight);
    }

    int buttonWidth = 0;
    if (buttonHeight > 0) {
        buttonWidth = std::min(kButtonWidth, buttonHeight);
    }
    const int maxAvailableWidth = std::max(0, bandWidth - kButtonMargin);
    if (buttonWidth == 0 && maxAvailableWidth > 0) {
        buttonWidth = std::min(kButtonWidth, maxAvailableWidth);
    }
    if (buttonWidth == 0 && bandWidth > 0) {
        buttonWidth = std::min(kButtonWidth, bandWidth);
    }
    const int trailingReserve = (buttonWidth > 0) ? std::min(buttonWidth + kTabGap + kButtonMargin, bandWidth) : 0;

    int x = boundsLeft + gripWidth - 3;   // DO NOT TOUCH

    const int startY = boundsTop + 2;
    const int maxX = std::max(boundsLeft + gripWidth - 3, boundsRight - trailingReserve);

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
            x = boundsLeft + gripWidth - 3;  // DO NOT TOUCH
            return true;
        }
        return false;
    };

    auto acquireReuse = [&](VisualItem& visual) -> const VisualItem* {
        if (!reuseContext || !reuseContext->source) {
            visual.reuseSourceIndex = kInvalidIndex;
            return nullptr;
        }
        const uint64_t key = visual.stableId != 0 ? visual.stableId : ComputeTabViewStableId(visual.data);
        auto it = reuseContext->indexByKey.find(key);
        if (it == reuseContext->indexByKey.end()) {
            visual.reuseSourceIndex = kInvalidIndex;
            return nullptr;
        }

        const auto& candidates = it->second;
        auto selectCandidate = [&](auto&& predicate) -> size_t {
            for (size_t idx : candidates) {
                if (idx >= reuseContext->reserved.size()) {
                    continue;
                }
                if (reuseContext->reserved[idx]) {
                    continue;
                }
                if (predicate((*reuseContext->source)[idx])) {
                    return idx;
                }
            }
            return kInvalidIndex;
        };

        size_t selected = selectCandidate([&](const VisualItem& candidate) {
            return candidate.indicatorHandle == visual.indicatorHandle &&
                   candidate.collapsedPlaceholder == visual.collapsedPlaceholder &&
                   candidate.hasGroupHeader == visual.hasGroupHeader &&
                   candidate.firstInGroup == visual.firstInGroup &&
                   EquivalentTabViewItem(candidate.data, visual.data);
        });
        if (selected == kInvalidIndex) {
            selected = selectCandidate([&](const VisualItem& candidate) {
                return candidate.indicatorHandle == visual.indicatorHandle &&
                       candidate.collapsedPlaceholder == visual.collapsedPlaceholder &&
                       candidate.hasGroupHeader == visual.hasGroupHeader;
            });
        }
        if (selected == kInvalidIndex) {
            selected = selectCandidate([&](const VisualItem& candidate) {
                return EquivalentTabViewItem(candidate.data, visual.data);
            });
        }
        if (selected == kInvalidIndex) {
            selected = selectCandidate([](const VisualItem&) { return true; });
        }

        if (selected == kInvalidIndex) {
            visual.reuseSourceIndex = kInvalidIndex;
            return nullptr;
        }

        reuseContext->reserved[selected] = true;
        visual.reuseSourceIndex = selected;
        return &(*reuseContext->source)[selected];
    };

    int currentGroup = -1;
    TabViewItem currentHeader{};
    bool headerMetadata = false;
    bool expectFirstTab = false;
    bool pendingIndicator = false;
    TabViewItem indicatorHeader{};

    result.items.reserve(items.size() + 8);

#if defined(_DEBUG)
    size_t cacheLookups = 0;
    size_t cacheHits = 0;
#endif

    RECT newTabBounds{};
    bool newTabVisible = false;

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

            if (currentGroup >= 0 && x > boundsLeft) {
                x += kGroupGap;
            }

            int width = kIslandIndicatorWidth;
            if (x + width > maxX) {
                if (!try_wrap()) {
                    break;
                }
            }

            VisualItem visual;
            visual.data = item;
            visual.stableId = item.stableId != 0 ? item.stableId : ComputeTabViewStableId(item);
            visual.firstInGroup = true;
            visual.collapsedPlaceholder = collapsed;
            visual.indicatorHandle = true;
            acquireReuse(visual);
            visual.bounds = {x, rowTop(row), x + width, rowBottom(row)};
            visual.row = row;
            result.items.emplace_back(std::move(visual));
            x += width;

            if (item.headerVisible && !collapsed && !hasVisibleTabs) {
                const int remaining = maxX - x;
                if (remaining > 0) {
                    const int placeholderWidth = std::min(remaining, kEmptyIslandBodyMaxWidth);
                    if (placeholderWidth > 0) {
                        RECT placeholder{ x, rowTop(row), x + placeholderWidth, rowBottom(row) };

                        VisualItem emptyBody;
                        emptyBody.data = item;
                        emptyBody.stableId = item.stableId != 0 ? item.stableId : ComputeTabViewStableId(item);
                        emptyBody.hasGroupHeader = true;
                        emptyBody.groupHeader = currentHeader;
                        acquireReuse(emptyBody);
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
                            m_emptyIslandPlusButtons.push_back({ currentGroup, plus, placeholder });
                        }

                        x = placeholder.right;
                    }
                }
            }
            continue;
        }

        VisualItem visual;
        visual.data = item;
        visual.stableId = item.stableId != 0 ? item.stableId : ComputeTabViewStableId(item);

        if (currentGroup != item.location.groupIndex) {
            currentGroup = item.location.groupIndex;
            headerMetadata = false;
            expectFirstTab = true;
            if (!result.items.empty()) {
                x += kGroupGap;
            }
            pendingIndicator = false;
        } else if (!expectFirstTab) {
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

        int measuredTextWidth = 0;
        if (!item.name.empty()) {
#if defined(_DEBUG)
            ++cacheLookups;
#endif
            if (widthCache.TryGet(item.name, metricsKey, measuredTextWidth)) {
#if defined(_DEBUG)
                ++cacheHits;
#endif
            } else {
                SIZE textSize{0, 0};
                GetTextExtentPoint32W(dc, item.name.c_str(), static_cast<int>(item.name.size()), &textSize);
                measuredTextWidth = textSize.cx;
                widthCache.Put(item.name, metricsKey, measuredTextWidth);
            }
        }

        int width = measuredTextWidth + kPaddingX * 2;
        if (item.pinned) {
            width += kPinnedGlyphWidth + kPinnedGlyphPadding;
        }
        width = std::max(width, kItemMinWidth);

        visual.badgeWidth = item.pinned ? (kPinnedGlyphWidth + kPinnedGlyphPadding) : 0;

        const VisualItem* preserved = acquireReuse(visual);
        if (preserved && EquivalentTabViewItem(preserved->data, visual.data) && preserved->icon) {
            visual.icon = preserved->icon;
            visual.iconWidth = preserved->iconWidth;
            visual.iconHeight = preserved->iconHeight;
            visual.reusedIconMetrics = true;
        }

        if (!visual.icon) {
            visual.icon = LoadItemIcon(item, SHGFI_SMALLICON);
            if (visual.icon) {
                visual.iconWidth = baseIconWidth;
                visual.iconHeight = baseIconHeight;
                if (auto metrics = visual.icon.GetMetrics()) {
                    visual.iconWidth = metrics->cx;
                    visual.iconHeight = metrics->cy;
                } else {
                    ICONINFO iconInfo{};
                    HICON iconHandle = visual.icon.Get();
                    if (iconHandle && GetIconInfo(iconHandle, &iconInfo)) {
                        BITMAP bitmap{};
                        if (iconInfo.hbmColor &&
                            GetObject(iconInfo.hbmColor, sizeof(bitmap), &bitmap) == sizeof(bitmap)) {
                            visual.iconWidth = bitmap.bmWidth;
                            visual.iconHeight = bitmap.bmHeight;
                        } else if (iconInfo.hbmMask &&
                                   GetObject(iconInfo.hbmMask, sizeof(bitmap), &bitmap) == sizeof(bitmap)) {
                            visual.iconWidth = bitmap.bmWidth;
                            visual.iconHeight = bitmap.bmHeight / 2;
                        }
                        if (iconInfo.hbmColor) {
                            DeleteObject(iconInfo.hbmColor);
                        }
                        if (iconInfo.hbmMask) {
                            DeleteObject(iconInfo.hbmMask);
                        }
                    }
                }
            }
        }

        if (visual.icon) {
            if (visual.iconWidth <= 0) {
                visual.iconWidth = baseIconWidth;
            }
            if (visual.iconHeight <= 0) {
                visual.iconHeight = baseIconHeight;
            }
            width += visual.iconWidth + kIconGap;
        }

        width += kCloseButtonSize + kCloseButtonEdgePadding + kCloseButtonSpacing;
        if (item.pinned) {
            width = std::min(width, kPinnedTabMaxWidth);
        }

        bool wrapped = false;
        if (x + width > maxX) {
            if (!try_wrap()) {
                width = std::max(40, maxX - x);
                if (width <= 0) {
                    break;
                }
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

        width = std::clamp(width, 40, std::max(40, maxX - x));

        visual.bounds = {x, rowTop(row), x + width, rowBottom(row)};
        visual.row = row;
        visual.index = result.items.size();
        result.items.emplace_back(std::move(visual));
        x += width;
    }

    if (buttonWidth > 0 && buttonHeight > 0) {
        int slotRow = std::clamp(row, 0, kMaxTabRows - 1);
        const int baseLeft = boundsLeft + gripWidth - 3;
        const int fallbackLeft = std::max(baseLeft, boundsRight - buttonWidth - kButtonMargin);

        int slotLeft = fallbackLeft;
        if (!result.items.empty()) {
            slotLeft = std::max(baseLeft, x + kTabGap);
            slotLeft = std::min(slotLeft, fallbackLeft);

            const VisualItem& tail = result.items.back();
            const int tailGroup = tail.data.location.groupIndex;
            for (auto it = m_emptyIslandPlusButtons.rbegin(); it != m_emptyIslandPlusButtons.rend(); ++it) {
                if (it->groupIndex == tailGroup) {
                    slotLeft = std::max(slotLeft, static_cast<int>(it->placeholder.right) + kTabGap);
                    slotLeft = std::min(slotLeft, fallbackLeft);
                    slotRow = tail.row;
                    break;
                }
            }
        }

        const int slotTopBound = rowTop(slotRow);
        const int slotBottomBound = rowBottom(slotRow);
        const int verticalSpace = slotBottomBound - slotTopBound;
        int slotTop = slotTopBound;
        if (verticalSpace > buttonHeight) {
            slotTop += (verticalSpace - buttonHeight) / 2;
        }
        if (slotTop + buttonHeight > slotBottomBound) {
            slotTop = std::max(slotTopBound, slotBottomBound - buttonHeight);
        }

        newTabBounds = {slotLeft, slotTop, slotLeft + buttonWidth, slotTop + buttonHeight};
        newTabVisible = true;
        maxRowUsed = std::max(maxRowUsed, slotRow);
    }

    if (row > maxRowUsed) {
        maxRowUsed = row;
    }
    result.rowCount = std::clamp(maxRowUsed + 1, 1, kMaxTabRows);

    result.newTabBounds = newTabBounds;
    result.newTabVisible = newTabVisible;

    if (oldFont) {
        SelectObject(dc, oldFont);
    }
    ReleaseDC(m_hwnd, dc);

#if defined(_DEBUG)
    if (cacheLookups > 0) {
        const size_t cacheMisses = cacheLookups - cacheHits;
        const double hitRate = cacheLookups > 0 ? (static_cast<double>(cacheHits) * 100.0) / cacheLookups : 0.0;
        LogMessage(LogLevel::Info,
                   L"Tab text width cache: lookups=%Iu hits=%Iu misses=%Iu hitRate=%.2f%%",
                   static_cast<size_t>(cacheLookups),
                   static_cast<size_t>(cacheHits),
                   static_cast<size_t>(cacheMisses),
                   hitRate);
    }
#endif
    return result;
}

void TabBandWindow::RebuildLayout() {
    if (!m_hwnd) {
        DestroyVisualItemResources(m_items);
        m_items.clear();
        m_progressRects.clear();
        m_activeProgressIndices.clear();
        m_emptyIslandPlusButtons.clear();
        SetRectEmpty(&m_newTabBounds);
        m_nextRedrawIncremental = false;
        InvalidateGroupOutlineCache();
        RebuildTabLocationIndex();
        return;
    }

    std::vector<VisualItem> oldItems;
    oldItems.swap(m_items);

    HideDragOverlay(true);
    HidePreviewWindow(false);
    DropTarget clearedDropTarget{};
    ApplyInternalDropTarget(m_drag.target, clearedDropTarget);
    m_drag = {};
    m_contextHit = {};
    m_emptyIslandPlusButtons.clear();

    LayoutResult layout = BuildLayoutItems(m_tabData, nullptr);
    RebuildTabLocationIndex();
    m_items = std::move(layout.items);
    if (layout.newTabVisible && layout.newTabBounds.right > layout.newTabBounds.left &&
        layout.newTabBounds.bottom > layout.newTabBounds.top) {
        m_newTabBounds = layout.newTabBounds;
        if (m_newTabButton) {
            const int width = m_newTabBounds.right - m_newTabBounds.left;
            const int height = m_newTabBounds.bottom - m_newTabBounds.top;
            MoveWindow(m_newTabButton, m_newTabBounds.left, m_newTabBounds.top, width, height, TRUE);
            ShowWindow(m_newTabButton, SW_SHOW);
        }
    } else {
        SetRectEmpty(&m_newTabBounds);
        if (m_newTabButton) {
            ShowWindow(m_newTabButton, SW_HIDE);
        }
    }
    RebuildProgressRectCache();
    RebuildGroupOutlineCache();

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

TabBandWindow::LayoutDiffStats TabBandWindow::ComputeLayoutDiff(std::vector<VisualItem>& oldItems,
                                                                std::vector<VisualItem>& newItems) const {
    LayoutDiffStats stats;

    if (!m_hwnd) {
        stats.inserted = newItems.size();
        stats.removed = oldItems.size();
        return stats;
    }

    std::unordered_map<uint64_t, std::vector<size_t>> oldMap;
    oldMap.reserve(oldItems.size());
    for (size_t i = 0; i < oldItems.size(); ++i) {
        const uint64_t stableId = oldItems[i].stableId != 0 ? oldItems[i].stableId
                                                            : ComputeTabViewStableId(oldItems[i].data);
        oldMap[stableId].push_back(i);
    }

    std::vector<bool> consumed(oldItems.size(), false);
    stats.matchedOldIndices.assign(newItems.size(), kInvalidIndex);

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

    for (size_t newIndex = 0; newIndex < newItems.size(); ++newIndex) {
        auto& item = newItems[newIndex];
        const uint64_t key = item.stableId != 0 ? item.stableId : ComputeTabViewStableId(item.data);
        auto it = oldMap.find(key);
        if (it == oldMap.end() || it->second.empty()) {
            ++stats.inserted;
            enqueueRect(item.bounds);
            continue;
        }

        auto& candidates = it->second;
        auto selectCandidate = [&](auto&& predicate) -> size_t {
            for (size_t idx = 0; idx < candidates.size(); ++idx) {
                const size_t candidateIndex = candidates[idx];
                if (predicate(oldItems[candidateIndex])) {
                    const size_t result = candidateIndex;
                    candidates[idx] = candidates.back();
                    candidates.pop_back();
                    return result;
                }
            }
            return kInvalidIndex;
        };

        size_t oldIndex = kInvalidIndex;
        if (item.reuseSourceIndex != kInvalidIndex && item.reuseSourceIndex < oldItems.size() &&
            !consumed[item.reuseSourceIndex]) {
            for (size_t idx = 0; idx < candidates.size(); ++idx) {
                if (candidates[idx] == item.reuseSourceIndex) {
                    oldIndex = item.reuseSourceIndex;
                    candidates[idx] = candidates.back();
                    candidates.pop_back();
                    break;
                }
            }
        }

        if (oldIndex == kInvalidIndex) {
            oldIndex = selectCandidate([&](const VisualItem& oldItem) {
                return oldItem.indicatorHandle == item.indicatorHandle &&
                       oldItem.collapsedPlaceholder == item.collapsedPlaceholder &&
                       oldItem.hasGroupHeader == item.hasGroupHeader &&
                       oldItem.firstInGroup == item.firstInGroup &&
                       EquivalentTabViewItem(oldItem.data, item.data);
            });
        }

        if (oldIndex == kInvalidIndex) {
            oldIndex = selectCandidate([&](const VisualItem& oldItem) {
                return oldItem.indicatorHandle == item.indicatorHandle &&
                       oldItem.collapsedPlaceholder == item.collapsedPlaceholder &&
                       oldItem.hasGroupHeader == item.hasGroupHeader;
            });
        }

        if (oldIndex == kInvalidIndex) {
            oldIndex = selectCandidate([&](const VisualItem& oldItem) {
                return EquivalentTabViewItem(oldItem.data, item.data);
            });
        }

        if (oldIndex == kInvalidIndex) {
            oldIndex = candidates.back();
            candidates.pop_back();
        }

        consumed[oldIndex] = true;
        VisualItem& oldItem = oldItems[oldIndex];
        VisualItem& newItem = item;
        stats.matchedOldIndices[newIndex] = oldIndex;
        newItem.reuseSourceIndex = oldIndex;

        if (oldItem.icon) {
            if (newItem.icon) {
                newItem.icon.Reset();
            }
            newItem.icon = std::move(oldItem.icon);
            newItem.iconWidth = oldItem.iconWidth;
            newItem.iconHeight = oldItem.iconHeight;
        }

        const bool moved = !EqualRect(&oldItem.bounds, &newItem.bounds);
        const bool metadataChanged = oldItem.firstInGroup != newItem.firstInGroup ||
                                     oldItem.badgeWidth != newItem.badgeWidth ||
                                     oldItem.hasGroupHeader != newItem.hasGroupHeader ||
                                     oldItem.collapsedPlaceholder != newItem.collapsedPlaceholder ||
                                     oldItem.indicatorHandle != newItem.indicatorHandle ||
                                     (newItem.hasGroupHeader &&
                                      !EquivalentTabViewItem(oldItem.groupHeader, newItem.groupHeader));

        const bool contentChanged = !EquivalentTabViewItem(oldItem.data, newItem.data) || metadataChanged;

        if (moved) {
            ++stats.moved;
        }
        if (contentChanged) {
            ++stats.updated;
        }
        if (moved || contentChanged) {
            RECT unionRect{};
            RECT oldRect = NormalizeRect(oldItem.bounds);
            RECT newRect = NormalizeRect(newItem.bounds);
            UnionRect(&unionRect, &oldRect, &newRect);
            enqueueRect(unionRect);
        }
    }

    for (size_t i = 0; i < consumed.size(); ++i) {
        if (!consumed[i]) {
            ++stats.removed;
            enqueueRect(oldItems[i].bounds);
            stats.removedIndices.push_back(i);
        }
    }

    return stats;
}

void TabBandWindow::ApplyPreservedVisualItems(const std::vector<VisualItem>& preserved,
                                              std::vector<VisualItem>& current,
                                              const LayoutDiffStats& diff) const {
    if (preserved.empty() || current.empty()) {
        return;
    }
    if (diff.matchedOldIndices.size() != current.size()) {
        return;
    }

    for (size_t index = 0; index < current.size(); ++index) {
        const size_t oldIndex = diff.matchedOldIndices[index];
        if (oldIndex == kInvalidIndex || oldIndex >= preserved.size()) {
            continue;
        }

        const VisualItem& oldItem = preserved[oldIndex];
        VisualItem& newItem = current[index];

        if (!newItem.icon && oldItem.icon) {
            newItem.icon = oldItem.icon;
            newItem.iconWidth = oldItem.iconWidth;
            newItem.iconHeight = oldItem.iconHeight;
        } else if (oldItem.icon && newItem.icon && (newItem.iconWidth <= 0 || newItem.iconHeight <= 0 ||
                                                    newItem.reusedIconMetrics)) {
            newItem.iconWidth = oldItem.iconWidth;
            newItem.iconHeight = oldItem.iconHeight;
        }

        if (EqualRect(&oldItem.bounds, &newItem.bounds)) {
            newItem.bounds = oldItem.bounds;
            newItem.row = oldItem.row;
        }
    }
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

        RebarColorScheme desired{};
        desired.background = CLR_DEFAULT;  // transparent to the rebar we paint
        desired.foreground = CLR_DEFAULT;

        if (m_lastRebarColors && *m_lastRebarColors == desired) {
                return;
        }

        const int count = static_cast<int>(SendMessageW(m_parentRebar, RB_GETBANDCOUNT, 0, 0));
        if (count <= 0) {
                return;
        }

        bool applied = false;
        for (int i = 0; i < count; ++i) {
                REBARBANDINFOW bi{ sizeof(bi) };
                bi.fMask = RBBIM_COLORS;
                bi.clrBack = desired.background;
                bi.clrFore = desired.foreground;
                if (SendMessageW(m_parentRebar, RB_SETBANDINFO, i, reinterpret_cast<LPARAM>(&bi))) {
                        applied = true;
                }
        }

        if (applied) {
                m_lastRebarColors = desired;
                m_rebarNeedsRepaint = true;
        }
}

void TabBandWindow::FlushRebarRepaint() {
        if (!m_rebarNeedsRepaint) {
                return;
        }
        HWND rebar = m_parentRebar;
        if (rebar && IsWindow(rebar)) {
                RedrawWindow(rebar, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
        }
        m_rebarNeedsRepaint = false;
}


void TabBandWindow::InvalidateRebarIntegration() {
        m_rebarIntegrationDirty = true;
        m_lastIntegratedRebar = nullptr;
        m_lastIntegratedFrame = nullptr;
        m_lastRebarColors.reset();
        m_rebarNeedsRepaint = false;
}

bool TabBandWindow::NeedsRebarIntegration() const {
        if (!m_hwnd) {
                return false;
        }
        if (m_rebarIntegrationDirty) {
                return true;
        }
        if (!m_parentRebar || !IsWindow(m_parentRebar)) {
                return true;
        }
        if (m_parentRebar != m_lastIntegratedRebar) {
                return true;
        }
        if (m_lastIntegratedFrame && !IsWindow(m_lastIntegratedFrame)) {
                return true;
        }
        HWND frame = GetAncestor(m_parentRebar, GA_ROOT);
        return frame != m_lastIntegratedFrame;
}

bool TabBandWindow::DrawRebarThemePart(HDC dc, const RECT& bounds, int partId, int stateId,
                                       bool suppressFallback, const GlowColorSet* overrideColors) const {
        if (!dc || !m_rebarTheme) {
                return false;
        }

        RECT partBounds = bounds;
        if (AreThemeHooksActive() && (overrideColors || suppressFallback)) {
                GlowColorSet colors = overrideColors ? *overrideColors : GlowColorSet{};
                ThemePaintOverrideGuard guard(m_hwnd, ExplorerSurfaceKind::Rebar, colors, suppressFallback);
                const HRESULT hr = DrawThemeBackground(m_rebarTheme, dc, partId, stateId, &partBounds, nullptr);
                if (SUCCEEDED(hr)) {
                        return true;
                }
                if (suppressFallback) {
                        return false;
                }
        }

        const HRESULT hr = DrawThemeBackground(m_rebarTheme, dc, partId, stateId, &partBounds, nullptr);
        return SUCCEEDED(hr);
}

void TabBandWindow::DrawBackground(HDC dc, const RECT& bounds) const {
        if (!dc) return;
        if (bounds.right <= bounds.left || bounds.bottom <= bounds.top) return;

        if (NeedsRebarIntegration()) {
                const_cast<TabBandWindow*>(this)->EnsureRebarIntegration();
        }

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
                const GlowColorSet colors = BuildRebarGlowColors(m_themePalette);
                if (DrawRebarThemePart(dc, bounds, RP_BACKGROUND, 0, true, &colors) ||
                    DrawRebarThemePart(dc, bounds, RP_BAND, 0, true, &colors)) {
                        backgroundDrawn = true;
                }
        }

	// Fallback fill (your code) now actually runs in dark mode:
        if (!backgroundDrawn && m_themePalette.rebarGradientValid) {
                TRIVERTEX vertices[2] = {
                        {bounds.left, bounds.top, static_cast<COLOR16>(GetRValue(m_themePalette.rebarGradientTop) << 8),
                         static_cast<COLOR16>(GetGValue(m_themePalette.rebarGradientTop) << 8),
                         static_cast<COLOR16>(GetBValue(m_themePalette.rebarGradientTop) << 8), 0},
                        {bounds.right, bounds.bottom,
                         static_cast<COLOR16>(GetRValue(m_themePalette.rebarGradientBottom) << 8),
                         static_cast<COLOR16>(GetGValue(m_themePalette.rebarGradientBottom) << 8),
                         static_cast<COLOR16>(GetBValue(m_themePalette.rebarGradientBottom) << 8), 0},
                };
                GRADIENT_RECT rect = {0, 1};
                if (GradientFill(dc, vertices, 2, &rect, 1, GRADIENT_FILL_RECT_V)) {
                        backgroundDrawn = true;
                }
        }

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
                        if (!DrawRebarThemePart(dc, gripRect, RP_GRIPPER, 0, false, nullptr)) {
                                DrawRebarThemePart(dc, gripRect, RP_GRIPPERVERT, 0, false, nullptr);
                        }
                }
        }
}


void TabBandWindow::Draw(HDC dc) {
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
    } guard(this, incremental);

    const bool sizeChanged = m_backBufferSize.cx != width || m_backBufferSize.cy != height;
    if (sizeChanged) {
        ReleaseBackBuffer();
    }

    if (!m_backBufferDC) {
        m_backBufferDC = CreateCompatibleDC(dc);
        if (!m_backBufferDC) {
            PaintSurface(dc, windowRect);
            return;
        }
    }

    if (!m_backBufferBitmap) {
        HBITMAP newBitmap = CreateCompatibleBitmap(dc, width, height);
        if (!newBitmap) {
            ReleaseBackBuffer();
            PaintSurface(dc, windowRect);
            return;
        }

        HGDIOBJ old = SelectObject(m_backBufferDC, newBitmap);
        if (!old || old == HGDI_ERROR) {
            DeleteObject(newBitmap);
            ReleaseBackBuffer();
            PaintSurface(dc, windowRect);
            return;
        }

        m_backBufferOldBitmap = old;
        m_backBufferBitmap = newBitmap;
        m_backBufferSize.cx = width;
        m_backBufferSize.cy = height;
    }

    RECT localRect{0, 0, width, height};
    PaintSurface(m_backBufferDC, localRect);
    BitBlt(dc, windowRect.left, windowRect.top, width, height, m_backBufferDC, 0, 0, SRCCOPY);
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
                const RECT& plus = b.plus;
                const int w = static_cast<int>(plus.right - plus.left);
                const int h = static_cast<int>(plus.bottom - plus.top);
                const int cx = plus.left + w / 2;
                const int cy = plus.top + h / 2;

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
    if (m_highContrast) {
        return item.selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_WINDOW);
    }
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
    if (m_highContrast) {
        return item.selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_BTNFACE);
    }
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
    if (m_highContrast) {
        return selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_WINDOWTEXT);
    }
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
    if (m_highContrast) {
        return item.selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_WINDOWTEXT);
    }
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
                if (PtInRect(&b.plus, pt)) {
			if (outGroupIndex) *outGroupIndex = b.groupIndex;
			return true;
		}
	}
	return false;
}



std::vector<TabBandWindow::GroupOutline> TabBandWindow::ComputeGroupOutlines() const {
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

const std::vector<TabBandWindow::GroupOutline>& TabBandWindow::BuildGroupOutlines() const {
    if (!m_groupOutlineCache.valid) {
        RebuildGroupOutlineCache();
    }
    return m_groupOutlineCache.outlines;
}

void TabBandWindow::InvalidateGroupOutlineCache() {
    m_groupOutlineCache.outlines.clear();
    m_groupOutlineCache.valid = false;
}

void TabBandWindow::RebuildGroupOutlineCache() const {
    m_groupOutlineCache.outlines = ComputeGroupOutlines();
    m_groupOutlineCache.valid = true;
}

bool TabBandWindow::DropPreviewAffectsIndicators(const DropTarget& target) const {
    if (!target.active || target.outside) {
        return false;
    }
    if (target.group) {
        return true;
    }
    if (target.newGroup) {
        return true;
    }
    return false;
}

void TabBandWindow::OnDropPreviewTargetChanged(const DropTarget& previous, const DropTarget& current) {
    const bool previousAffects = DropPreviewAffectsIndicators(previous);
    const bool currentAffects = DropPreviewAffectsIndicators(current);
    if (!previousAffects && !currentAffects) {
        return;
    }
    if (previousAffects != currentAffects || previous.group != current.group ||
        previous.groupIndex != current.groupIndex || previous.newGroup != current.newGroup ||
        previous.floating != current.floating) {
        RebuildGroupOutlineCache();
    }
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
                if (self) {
                        self->m_rebarSubclassed = false;
                        self->m_parentRebar = nullptr;
                        self->m_rebarBandIndex = -1;
                        self->InvalidateRebarIntegration();
                }
                break;
        }
	return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void TabBandWindow::InstallRebarDarkSubclass() {
	if (!m_parentRebar || !IsWindow(m_parentRebar) || m_rebarSubclassed) return;

	// Keep Explorer's theme resources; we only overpaint the bg.
        SetWindowTheme(m_parentRebar, L"Explorer", nullptr);  // was nullptr, which destabilized the band site
        ApplyImmersiveDarkMode(m_parentRebar, m_darkMode && !m_highContrast);

        if (SetWindowSubclass(m_parentRebar, RebarSubclassProc, 0,
                reinterpret_cast<DWORD_PTR>(this))) {
                m_rebarSubclassed = true;
                m_rebarNeedsRepaint = true;
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
    ClearGdiCache();
    m_toolbarGripWidth = kToolbarGripWidth;
    if (!m_hwnd) { /* existing reset block unchanged */ return; }

    m_themeNotifier.RefreshColorsFromSystem();
    m_themeColors = m_themeNotifier.GetThemeColors();
    m_highContrast = IsHighContrastActive();

    // Ensure the band window itself opts into Explorer's visual styles so the
    // subsequent theme handles pull the correct resources for both light and
    // dark modes.
    SetWindowTheme(m_hwnd, L"Explorer", nullptr);
    const bool darkMode = IsSystemDarkMode();
    const bool immersiveDark = !m_highContrast && darkMode;
    if (!m_windowDarkModeInitialized || immersiveDark != m_windowDarkModeValue) {
        ApplyImmersiveDarkMode(m_hwnd, immersiveDark);
        m_windowDarkModeInitialized = true;
        m_windowDarkModeValue = immersiveDark;
    }
    m_darkMode = immersiveDark;

    // Ensure the parent rebar picks up the refreshed theme immediately.
    InvalidateRebarIntegration();
    EnsureRebarIntegration();
    AdjustBandHeightToRow();
    if (m_parentRebar) {
        InstallRebarDarkSubclass();   // NEW: we own the bar bg now
    }

    UpdateAccentColor();
    ResetThemePalette();

        const auto openTheme = [&](const wchar_t* classList, const wchar_t* operation) -> HTHEME {
            SetLastError(ERROR_SUCCESS);
            HTHEME handle = OpenThemeData(m_hwnd, classList);
            if (!handle) {
                const DWORD error = GetLastError();
                if (error != ERROR_SUCCESS) {
                    LogLastError(operation, error);
                } else {
                    LogMessage(LogLevel::Error,
                               L"%s failed: OpenThemeData returned nullptr without extended error.",
                               operation);
                }
            }
            return handle;
        };

        HTHEME tabTheme = openTheme(L"Tab", L"OpenThemeData(Tab)");
        HTHEME rebarTheme = openTheme(L"Rebar", L"OpenThemeData(Rebar)");
        HTHEME windowTheme = openTheme(L"Window", L"OpenThemeData(Window)");

        if (!tabTheme || !rebarTheme || !windowTheme) {
            if (tabTheme) {
                CloseThemeData(tabTheme);
                tabTheme = nullptr;
            }
            if (rebarTheme) {
                CloseThemeData(rebarTheme);
                rebarTheme = nullptr;
            }
            if (windowTheme) {
                CloseThemeData(windowTheme);
                windowTheme = nullptr;
            }
        }

    m_tabTheme = tabTheme;
    m_rebarTheme = rebarTheme;
    m_windowTheme = windowTheme;
    UpdateThemePalette();
    if (m_parentRebar) {
        UpdateRebarColors();
    }
    FlushRebarRepaint();
    UpdateToolbarMetrics();
    UpdateNewTabButtonTheme();
    RebuildLayout();

}

void TabBandWindow::OnSavedGroupsChanged() {
    if (m_owner) {
        m_owner->OnSavedGroupsChanged();
    }
    if (m_hwnd && IsWindow(m_hwnd)) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
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
    m_themePalette.rebarGradientValid = false;

    ClearTextWidthCache();

    if (m_highContrast) {
        const COLORREF windowColor = GetSysColor(COLOR_WINDOW);
        const COLORREF buttonColor = GetSysColor(COLOR_BTNFACE);
        const COLORREF highlight = GetSysColor(COLOR_HIGHLIGHT);
        const COLORREF frame = GetSysColor(COLOR_WINDOWFRAME);

        m_themePalette.rebarBackground = buttonColor;
        m_themePalette.rebarGradientTop = buttonColor;
        m_themePalette.rebarGradientBottom = buttonColor;
        m_themePalette.borderTop = frame;
        m_themePalette.borderBottom = frame;
        m_themePalette.tabBase = windowColor;
        m_themePalette.tabSelectedBase = highlight;
        m_themePalette.tabText = GetSysColor(COLOR_WINDOWTEXT);
        m_themePalette.tabSelectedText = GetSysColor(COLOR_HIGHLIGHTTEXT);
        m_themePalette.groupBase = buttonColor;
        m_themePalette.groupText = GetSysColor(COLOR_BTNTEXT);
        m_themePalette.tabTextValid = true;
        m_themePalette.tabSelectedTextValid = true;
        m_themePalette.groupTextValid = true;
        return;
    }

    const COLORREF windowColor = m_themeColors.valid ? m_themeColors.background : GetSysColor(COLOR_WINDOW);
    const COLORREF buttonColor = GetSysColor(COLOR_BTNFACE);
    const COLORREF foregroundColor = m_themeColors.valid ? m_themeColors.foreground : GetSysColor(COLOR_WINDOWTEXT);

    RECT windowRect{};
    if (m_hwnd) {
        GetWindowRect(m_hwnd, &windowRect);
    }
    HWND host = m_parentRebar ? m_parentRebar : GetParent(m_hwnd);
    HostChromeSample chrome = SampleHostChrome(host, windowRect);

    const COLORREF baseBackground = chrome.valid ? BlendColors(chrome.top, chrome.bottom, 0.5)
                                                 : AdjustForDarkTone(windowColor, 0.55, m_darkMode);
    m_themePalette.rebarBackground = baseBackground;
    if (chrome.valid) {
        m_themePalette.rebarGradientTop = chrome.top;
        m_themePalette.rebarGradientBottom = chrome.bottom;
        m_themePalette.rebarGradientValid = true;
    } else {
        m_themePalette.rebarGradientTop = baseBackground;
        m_themePalette.rebarGradientBottom = baseBackground;
    }

    if (m_darkMode) {
        m_themePalette.borderTop = BlendColors(baseBackground, RGB(0, 0, 0), 0.6);
        m_themePalette.borderBottom = BlendColors(baseBackground, RGB(255, 255, 255), 0.18);
    } else {
        m_themePalette.borderTop = BlendColors(baseBackground, RGB(255, 255, 255), 0.18);
        m_themePalette.borderBottom = BlendColors(baseBackground, RGB(0, 0, 0), 0.22);
    }

    const COLORREF tabBase = chrome.valid ? BlendColors(baseBackground, windowColor, m_darkMode ? 0.25 : 0.12)
                                          : AdjustForDarkTone(windowColor, 0.4, m_darkMode);
    m_themePalette.tabBase = tabBase;
    m_themePalette.tabSelectedBase = BlendColors(tabBase, m_accentColor, m_darkMode ? 0.5 : 0.35);
    m_themePalette.tabText = foregroundColor;
    m_themePalette.tabSelectedText = GetSysColor(COLOR_HIGHLIGHTTEXT);
    m_themePalette.tabTextValid = true;
    m_themePalette.tabSelectedTextValid = true;

    const COLORREF groupBase = chrome.valid ? BlendColors(baseBackground, buttonColor, 0.5)
                                            : BlendColors(buttonColor, windowColor, m_darkMode ? 0.55 : 0.25);
    m_themePalette.groupBase = groupBase;
    m_themePalette.groupText = foregroundColor;
    m_themePalette.groupTextValid = true;
}

void TabBandWindow::UpdateThemePalette() {
    if (m_highContrast) {
        return;
    }
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
    if (m_highContrast) {
        return;
    }
    auto& store = OptionsStore::Instance();
    static bool loggedOptionsLoadFailure = false;
    std::wstring errorContext;
    if (!store.Load(&errorContext)) {
        if (!loggedOptionsLoadFailure) {
            if (!errorContext.empty()) {
                LogMessage(LogLevel::Warning,
                           L"TabBandWindow::ApplyOptionColorOverrides failed to load options: %ls",
                           errorContext.c_str());
            } else {
                LogMessage(LogLevel::Warning,
                           L"TabBandWindow::ApplyOptionColorOverrides failed to load options");
            }
            loggedOptionsLoadFailure = true;
        }
    } else if (loggedOptionsLoadFailure) {
        loggedOptionsLoadFailure = false;
    }
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
        if (SendMessageW(m_parentRebar, RB_SETBANDINFO, m_rebarBandIndex,
                reinterpret_cast<LPARAM>(&colorInfo))) {
                m_lastRebarColors = RebarColorScheme{barBk, CLR_DEFAULT};
        }

	// Tone down etched highlights so the bar doesn't glow in dark mode.
	COLORSCHEME cs{};
	cs.dwSize = sizeof(cs);
        cs.clrBtnHighlight = (m_darkMode && !m_highContrast) ? barBk : CLR_DEFAULT;
        cs.clrBtnShadow = (m_darkMode && !m_highContrast) ? barBk : CLR_DEFAULT;
	SendMessageW(m_parentRebar, RB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&cs));

	// Repaint without forcing an erase (prevents flicker).
	RedrawWindow(m_parentRebar, nullptr, nullptr,
		RDW_INVALIDATE | RDW_FRAME );
}


void TabBandWindow::EnsureRebarIntegration() {
        if (!m_hwnd) return;

        const HWND previousRebar = m_parentRebar;
        const bool previouslyValidRebar = previousRebar && IsWindow(previousRebar);

        HWND parent = GetParent(m_hwnd);
        while (parent && !IsRebarWindow(parent)) {
                parent = GetParent(parent);
        }
        if (parent != m_parentRebar) {
                m_parentRebar = parent;
                m_rebarBandIndex = -1;
                m_rebarSubclassed = false;
                m_rebarIntegrationDirty = true;
                m_lastRebarColors.reset();
                m_rebarNeedsRepaint = false;
        }

        const bool hasValidRebar = m_parentRebar && IsWindow(m_parentRebar);
        const bool rebarNewlyAcquired = hasValidRebar && (!previouslyValidRebar || m_parentRebar != previousRebar);

        if (!hasValidRebar) {
                if (m_parentFrame) {
                        ClearAvailableDockMaskForFrame(m_parentFrame);
                        m_parentFrame = nullptr;
                }
                m_lastIntegratedRebar = nullptr;
                m_lastIntegratedFrame = nullptr;
                m_rebarIntegrationDirty = true;
                m_lastRebarColors.reset();
                m_rebarNeedsRepaint = false;
                return;
        }

        HWND frame = GetAncestor(m_parentRebar, GA_ROOT);
        if (frame != m_parentFrame) {
                if (m_parentFrame) {
                        ClearAvailableDockMaskForFrame(m_parentFrame);
                }
                m_parentFrame = frame;
                m_rebarIntegrationDirty = true;
        }

        if (!m_rebarIntegrationDirty &&
            m_parentRebar == m_lastIntegratedRebar &&
            frame == m_lastIntegratedFrame) {
                return;
        }

        if (rebarNewlyAcquired) {
                // The rebar appeared after an earlier miss; reapply dark mode and palette now.
                const bool immersiveDark = m_darkMode && !m_highContrast;
                ApplyImmersiveDarkMode(m_parentRebar, immersiveDark);
                InstallRebarDarkSubclass();
                ResetThemePalette();
                UpdateThemePalette();
                UpdateRebarColors();
                UpdateNewTabButtonTheme();
                FlushRebarRepaint();
                InvalidateRect(m_hwnd, nullptr, TRUE);
        }

        if (frame) {
                UpdateAvailableDockMaskFromFrame(frame);
        }

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

        m_lastIntegratedRebar = m_parentRebar;
        m_lastIntegratedFrame = frame;
        m_rebarIntegrationDirty = false;
}


void TabBandWindow::UpdateToolbarMetrics() {
    m_toolbarGripWidth = kToolbarGripWidth;
    EnsureRebarIntegration();

    bool gripWidthResolved = false;
    if (m_parentRebar && m_rebarBandIndex >= 0) {
        RECT borders{0, 0, 0, 0};
        if (SendMessageW(m_parentRebar, RB_GETBANDBORDERS, m_rebarBandIndex, reinterpret_cast<LPARAM>(&borders))) {
            const LONG candidate = std::max<LONG>(borders.left, 8L);
            if (candidate > 0) {
                m_toolbarGripWidth = candidate;
                gripWidthResolved = true;
            }
        }
    }

    if (!m_hwnd) {
        ResetCloseButtonMetrics();
        return;
    }

    const UINT currentDpi = GetDpiForWindow(m_hwnd);
    const bool shouldUpdateCloseButton =
        m_windowTheme && (!m_closeButtonSizeCached || m_cachedCloseButtonDpi != currentDpi);

    HDC dc = nullptr;
    auto ensureDc = [&]() -> HDC {
        if (!dc) {
            dc = GetDC(m_hwnd);
        }
        return dc;
    };

    if (!gripWidthResolved && m_rebarTheme) {
        if (HDC themeDc = ensureDc()) {
            int part = RP_GRIPPER;
            SIZE gripSize{0, 0};
            HRESULT hr = GetThemePartSize(m_rebarTheme, themeDc, part, 0, nullptr, TS_TRUE, &gripSize);
            if (FAILED(hr) || gripSize.cx <= 0) {
                part = RP_GRIPPERVERT;
                gripSize = {0, 0};
                hr = GetThemePartSize(m_rebarTheme, themeDc, part, 0, nullptr, TS_TRUE, &gripSize);
            }

            if (SUCCEEDED(hr) && gripSize.cx > 0) {
                int width = gripSize.cx;
                MARGINS margins{0, 0, 0, 0};
                if (SUCCEEDED(GetThemeMargins(m_rebarTheme, themeDc, part, 0, TMT_CONTENTMARGINS, nullptr, &margins))) {
                    width += margins.cxLeftWidth + margins.cxRightWidth;
                }
                if (width > 0) {
                    m_toolbarGripWidth = std::max(width, 8);
                }
                gripWidthResolved = true;
            }
        }
    }

    if (!m_windowTheme) {
        ResetCloseButtonMetrics();
    } else if (shouldUpdateCloseButton) {
        bool updated = false;
        if (HDC themeDc = ensureDc()) {
            SIZE themeSize{0, 0};
            int candidate = kCloseButtonSize;
            const HRESULT hr =
                GetThemePartSize(m_windowTheme, themeDc, WP_SMALLCLOSEBUTTON, 0, nullptr, TS_TRUE, &themeSize);
            if (SUCCEEDED(hr) && themeSize.cx > 0 && themeSize.cy > 0) {
                candidate = std::max(themeSize.cx, themeSize.cy);
            }
            m_cachedCloseButtonSize = candidate;
            m_cachedCloseButtonDpi = currentDpi;
            m_closeButtonSizeCached = true;
            updated = true;
        }
        if (!updated) {
            m_closeButtonSizeCached = false;
        }
    }

    if (dc) {
        ReleaseDC(m_hwnd, dc);
    }
}

void TabBandWindow::ResetCloseButtonMetrics() {
    m_closeButtonSizeCached = false;
    m_cachedCloseButtonSize = 0;
    m_cachedCloseButtonDpi = 0;
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
    ResetCloseButtonMetrics();
}

void TabBandWindow::HandleDpiChanged(UINT dpiX, UINT dpiY, const RECT* suggestedRect) {
    UNREFERENCED_PARAMETER(dpiX);
    UNREFERENCED_PARAMETER(dpiY);
    if (!m_hwnd) {
        return;
    }
    if (suggestedRect) {
        const int width = suggestedRect->right - suggestedRect->left;
        const int height = suggestedRect->bottom - suggestedRect->top;
        SetWindowPos(m_hwnd, nullptr, suggestedRect->left, suggestedRect->top, width, height,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }
    m_closeButtonSizeCached = false;
    UpdateToolbarMetrics();
    RebuildLayout();
}

HBRUSH TabBandWindow::GetCachedBrush(COLORREF color) const {
    auto it = m_brushCache.find(color);
    if (it != m_brushCache.end()) {
        return it->second.Get();
    }
    HBRUSH brush = CreateSolidBrush(color);
    if (!brush) {
        return nullptr;
    }
    m_brushCache.emplace(color, BrushHandle(brush));
    return brush;
}

HPEN TabBandWindow::GetCachedPen(COLORREF color, int width, int style) const {
    PenKey key{color, width, style};
    auto it = m_penCache.find(key);
    if (it != m_penCache.end()) {
        return it->second.Get();
    }
    HPEN pen = CreatePen(style, width, color);
    if (!pen) {
        return nullptr;
    }
    m_penCache.emplace(key, PenHandle(pen));
    return pen;
}

void TabBandWindow::ClearGdiCache() {
    for (auto& entry : m_brushCache) {
        entry.second.Reset();
    }
    m_brushCache.clear();
    for (auto& entry : m_penCache) {
        entry.second.Reset();
    }
    m_penCache.clear();
}

void TabBandWindow::UpdateNewTabButtonTheme() {
    if (!m_newTabButton) {
        m_newTabButtonHot = false;
        m_newTabButtonPressed = false;
        m_newTabButtonKeyboardPressed = false;
        m_newTabButtonTrackingMouse = false;
        m_newTabButtonPointerPressed = false;
        m_newTabButtonCommandPending = false;
        return;
    }
    m_newTabButtonTrackingMouse = false;
    m_newTabButtonPointerPressed = false;
    m_newTabButtonCommandPending = false;
    InvalidateRect(m_newTabButton, nullptr, TRUE);
}

void TabBandWindow::PaintNewTabButton(HWND hwnd, HDC dc) const {
    if (!hwnd || !dc) {
        return;
    }

    RECT bounds{};
    GetClientRect(hwnd, &bounds);
    const int width = bounds.right - bounds.left;
    const int height = bounds.bottom - bounds.top;
    if (width <= 0 || height <= 0) {
        return;
    }

    COLORREF hostBackground = m_highContrast ? GetSysColor(COLOR_BTNFACE) : m_themePalette.rebarBackground;
    HBRUSH hostBrush = GetCachedBrush(hostBackground);
    if (hostBrush) {
        FillRect(dc, &bounds, hostBrush);
    } else {
        FillRect(dc, &bounds, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
    }

    UINT dpi = 96;
    if (hwnd) {
        const UINT windowDpi = GetDpiForWindow(hwnd);
        if (windowDpi != 0) {
            dpi = windowDpi;
        }
    }

    const int outerMargin = std::max(2, MulDiv(4, static_cast<int>(dpi), 96));
    int squareSize = MulDiv(18, static_cast<int>(dpi), 96);
    squareSize = std::min(squareSize, width - outerMargin * 2);
    squareSize = std::min(squareSize, height - outerMargin * 2);
    if (squareSize < 6) {
        squareSize = std::min(width, height) - outerMargin * 2;
    }
    if (squareSize <= 0) {
        return;
    }

    const int squareLeft = bounds.left + (width - squareSize) / 2;
    const int squareTop = bounds.top + (height - squareSize) / 2;
    RECT square{squareLeft, squareTop, squareLeft + squareSize, squareTop + squareSize};

    COLORREF baseSquare = m_highContrast ? GetSysColor(COLOR_BTNFACE) : RGB(240, 240, 240);
    COLORREF borderColor = m_highContrast ? GetSysColor(COLOR_WINDOWFRAME) : RGB(200, 200, 200);
    COLORREF glyphColor = m_highContrast ? GetSysColor(COLOR_BTNTEXT) : RGB(64, 64, 64);

    if (m_highContrast) {
        if (m_newTabButtonPressed) {
            baseSquare = GetSysColor(COLOR_HIGHLIGHT);
            glyphColor = GetSysColor(COLOR_HIGHLIGHTTEXT);
            borderColor = GetSysColor(COLOR_HIGHLIGHT);
        } else if (m_newTabButtonHot || m_newTabButtonKeyboardPressed) {
            borderColor = GetSysColor(COLOR_HIGHLIGHT);
        }
    } else if (m_darkMode) {
        baseSquare = BlendColors(RGB(255, 255, 255), RGB(70, 70, 70), 0.35);
        borderColor = BlendColors(baseSquare, RGB(0, 0, 0), 0.4);
        glyphColor = RGB(32, 32, 32);
    }

    if (!m_highContrast) {
        if (m_newTabButtonPressed) {
            baseSquare = BlendColors(baseSquare, RGB(0, 0, 0), 0.2);
            glyphColor = BlendColors(glyphColor, RGB(0, 0, 0), 0.2);
        } else if (m_newTabButtonHot || m_newTabButtonKeyboardPressed) {
            baseSquare = BlendColors(baseSquare, RGB(255, 255, 255), 0.18);
        }
    }

    int cornerRadius = std::max(2, MulDiv(3, static_cast<int>(dpi), 96));
    cornerRadius = std::min(cornerRadius, squareSize / 2);
    HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
    HBRUSH fillBrush = CreateSolidBrush(baseSquare);
    if (borderPen && fillBrush) {
        const HPEN oldPen = static_cast<HPEN>(SelectObject(dc, borderPen));
        const HBRUSH oldBrush = static_cast<HBRUSH>(SelectObject(dc, fillBrush));
        RoundRect(dc, square.left, square.top, square.right, square.bottom, cornerRadius, cornerRadius);
        SelectObject(dc, oldBrush);
        SelectObject(dc, oldPen);
    }
    if (fillBrush) {
        DeleteObject(fillBrush);
    }
    if (borderPen) {
        DeleteObject(borderPen);
    }

    const int glyphPadding = std::max(2, MulDiv(4, static_cast<int>(dpi), 96));
    const int glyphExtent = std::max(4, squareSize - glyphPadding * 2);
    const int glyphHalf = glyphExtent / 2;
    const int centerX = square.left + squareSize / 2;
    const int centerY = square.top + squareSize / 2;
    int glyphThickness = std::max(2, MulDiv(3, static_cast<int>(dpi), 96));
    glyphThickness = std::min(glyphThickness, std::max(2, glyphExtent / 3));

    HBRUSH glyphBrush = CreateSolidBrush(glyphColor);
    if (glyphBrush) {
        RECT horizontal{
            centerX - glyphHalf,
            centerY - glyphThickness / 2,
            centerX + glyphHalf + 1,
            centerY - glyphThickness / 2 + glyphThickness
        };
        RECT vertical{
            centerX - glyphThickness / 2,
            centerY - glyphHalf,
            centerX - glyphThickness / 2 + glyphThickness,
            centerY + glyphHalf + 1
        };

        horizontal.top = std::max(horizontal.top, square.top + 1);
        horizontal.bottom = std::min(horizontal.bottom, square.bottom - 1);
        vertical.left = std::max(vertical.left, square.left + 1);
        vertical.right = std::min(vertical.right, square.right - 1);

        if (horizontal.right > horizontal.left && horizontal.bottom > horizontal.top) {
            FillRect(dc, &horizontal, glyphBrush);
        }
        if (vertical.right > vertical.left && vertical.bottom > vertical.top) {
            FillRect(dc, &vertical, glyphBrush);
        }

        DeleteObject(glyphBrush);
    }

    if (GetFocus() == hwnd) {
        RECT focusRect = square;
        const int inflate = std::max(1, MulDiv(2, static_cast<int>(dpi), 96));
        InflateRect(&focusRect, inflate, inflate);
        DrawFocusRect(dc, &focusRect);
    }
}

void TabBandWindow::HandleNewTabButtonMouseMove(HWND hwnd, POINT pt) {
    if (!hwnd) {
        return;
    }

    RECT bounds{};
    GetClientRect(hwnd, &bounds);
    const bool inside = PtInRect(&bounds, pt) != 0;

    bool stateChanged = false;
    if (inside != m_newTabButtonHot) {
        m_newTabButtonHot = inside;
        stateChanged = true;
    }

    const bool shouldAppearPressed =
        (m_newTabButtonPointerPressed && inside) || m_newTabButtonKeyboardPressed;
    if (m_newTabButtonPressed != shouldAppearPressed) {
        m_newTabButtonPressed = shouldAppearPressed;
        stateChanged = true;
    }

    if (!m_newTabButtonTrackingMouse) {
        TRACKMOUSEEVENT track{sizeof(track)};
        track.dwFlags = TME_LEAVE;
        track.hwndTrack = hwnd;
        if (TrackMouseEvent(&track)) {
            m_newTabButtonTrackingMouse = true;
        }
    }

    if (stateChanged) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void TabBandWindow::HandleNewTabButtonMouseLeave(HWND hwnd) {
    m_newTabButtonTrackingMouse = false;

    bool stateChanged = false;
    if (m_newTabButtonHot) {
        m_newTabButtonHot = false;
        stateChanged = true;
    }

    const bool shouldAppearPressed = m_newTabButtonKeyboardPressed;
    if (m_newTabButtonPressed != shouldAppearPressed) {
        m_newTabButtonPressed = shouldAppearPressed;
        stateChanged = true;
    }

    if (stateChanged) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void TabBandWindow::HandleNewTabButtonLButtonDown(HWND hwnd, POINT pt) {
    if (!hwnd) {
        return;
    }

    SetFocus(hwnd);
    SetCapture(hwnd);
    m_newTabButtonPointerPressed = true;
    m_newTabButtonKeyboardPressed = false;
    m_newTabButtonCommandPending = true;

    const bool inside = IsPointInsideNewTabButton(hwnd, pt);
    bool stateChanged = false;
    if (inside != m_newTabButtonHot) {
        m_newTabButtonHot = inside;
        stateChanged = true;
    }

    if (m_newTabButtonPressed != inside) {
        m_newTabButtonPressed = inside;
        stateChanged = true;
    }

    if (!m_newTabButtonTrackingMouse) {
        TRACKMOUSEEVENT track{sizeof(track)};
        track.dwFlags = TME_LEAVE;
        track.hwndTrack = hwnd;
        if (TrackMouseEvent(&track)) {
            m_newTabButtonTrackingMouse = true;
        }
    }

    if (stateChanged) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void TabBandWindow::HandleNewTabButtonLButtonUp(HWND hwnd, POINT pt) {
    if (GetCapture() == hwnd) {
        ReleaseCapture();
    }

    const bool inside = IsPointInsideNewTabButton(hwnd, pt);
    const bool shouldInvoke = m_newTabButtonPointerPressed && m_newTabButtonCommandPending && inside;

    m_newTabButtonPointerPressed = false;
    m_newTabButtonCommandPending = false;

    bool stateChanged = false;
    if (!inside && m_newTabButtonHot) {
        m_newTabButtonHot = false;
        stateChanged = true;
    }

    const bool shouldAppearPressed = m_newTabButtonKeyboardPressed;
    if (m_newTabButtonPressed != shouldAppearPressed) {
        m_newTabButtonPressed = shouldAppearPressed;
        stateChanged = true;
    }

    if (stateChanged) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    if (shouldInvoke) {
        TriggerNewTabButtonAction();
    }
}

void TabBandWindow::HandleNewTabButtonCaptureLost() {
    bool stateChanged = false;

    if (m_newTabButtonPointerPressed) {
        m_newTabButtonPointerPressed = false;
        stateChanged = true;
    }

    const bool shouldAppearPressed = m_newTabButtonKeyboardPressed;
    if (m_newTabButtonPressed != shouldAppearPressed) {
        m_newTabButtonPressed = shouldAppearPressed;
        stateChanged = true;
    }

    m_newTabButtonCommandPending = m_newTabButtonKeyboardPressed;

    if (stateChanged && m_newTabButton) {
        InvalidateRect(m_newTabButton, nullptr, FALSE);
    }
}

void TabBandWindow::HandleNewTabButtonFocusChanged(HWND hwnd, bool focused) {
    HWND parent = GetParent(hwnd);
    if (parent) {
        const int code = focused ? BN_SETFOCUS : BN_KILLFOCUS;
        SendMessageW(parent, WM_COMMAND, MAKEWPARAM(IDC_NEW_TAB, code), reinterpret_cast<LPARAM>(hwnd));
    }

    if (focused) {
        return;
    }

    bool stateChanged = false;
    if (m_newTabButtonPointerPressed) {
        m_newTabButtonPointerPressed = false;
        stateChanged = true;
    }
    if (m_newTabButtonKeyboardPressed) {
        m_newTabButtonKeyboardPressed = false;
        stateChanged = true;
    }
    if (m_newTabButtonPressed) {
        m_newTabButtonPressed = false;
        stateChanged = true;
    }
    m_newTabButtonCommandPending = false;

    if (stateChanged) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void TabBandWindow::HandleNewTabButtonKeyDown(HWND hwnd, UINT key, bool repeat) {
    UNREFERENCED_PARAMETER(key);
    if (repeat) {
        return;
    }

    m_newTabButtonPointerPressed = false;
    m_newTabButtonKeyboardPressed = true;
    m_newTabButtonCommandPending = true;

    if (!m_newTabButtonPressed) {
        m_newTabButtonPressed = true;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void TabBandWindow::HandleNewTabButtonKeyUp(HWND hwnd, UINT key) {
    UNREFERENCED_PARAMETER(key);
    if (!m_newTabButtonKeyboardPressed) {
        return;
    }

    m_newTabButtonKeyboardPressed = false;
    const bool shouldInvoke = m_newTabButtonCommandPending;
    m_newTabButtonCommandPending = false;

    if (m_newTabButtonPressed) {
        m_newTabButtonPressed = false;
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    if (shouldInvoke) {
        TriggerNewTabButtonAction();
    }
}

bool TabBandWindow::IsPointInsideNewTabButton(HWND hwnd, POINT pt) const {
    if (!hwnd) {
        return false;
    }
    RECT bounds{};
    GetClientRect(hwnd, &bounds);
    return PtInRect(&bounds, pt) != 0;
}

void TabBandWindow::TriggerNewTabButtonAction() {
    if (!m_hwnd || !m_newTabButton) {
        return;
    }
    // Issue the request directly so Explorer cannot swallow or duplicate our
    // WM_COMMAND dispatch when the custom "+" button is clicked.
    RequestNewTab();
}

void TabBandWindow::RequestNewTab() {
    if (!m_owner) {
        return;
    }
    m_owner->OnNewTabRequested();
}

LRESULT CALLBACK NewTabButtonWndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* owner = reinterpret_cast<TabBandWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (message == WM_NCCREATE) {
        auto* create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
        owner = static_cast<TabBandWindow*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(owner));
        return owner ? TRUE : FALSE;
    }

    if (!owner) {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    switch (message) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC dc = BeginPaint(hwnd, &ps);
            if (dc) {
                owner->PaintNewTabButton(hwnd, dc);
            }
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_ERASEBKGND:
            return 1;
        case WM_MOUSEMOVE: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            owner->HandleNewTabButtonMouseMove(hwnd, pt);
            return 0;
        }
        case WM_MOUSELEAVE:
            owner->HandleNewTabButtonMouseLeave(hwnd);
            return 0;
        case WM_LBUTTONDOWN: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            owner->HandleNewTabButtonLButtonDown(hwnd, pt);
            return 0;
        }
        case WM_LBUTTONUP: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            owner->HandleNewTabButtonLButtonUp(hwnd, pt);
            return 0;
        }
        case WM_LBUTTONDBLCLK:
            // Consume double-click to prevent creating multiple tabs
            return 0;
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
            owner->HandleNewTabButtonCaptureLost();
            return 0;
        case WM_SETFOCUS:
            owner->HandleNewTabButtonFocusChanged(hwnd, true);
            return 0;
        case WM_KILLFOCUS:
            owner->HandleNewTabButtonFocusChanged(hwnd, false);
            return 0;
        case WM_KEYDOWN:
            if (wParam == VK_SPACE || wParam == VK_RETURN) {
                const bool repeat = (HIWORD(lParam) & KF_REPEAT) != 0;
                owner->HandleNewTabButtonKeyDown(hwnd, static_cast<UINT>(wParam), repeat);
                return 0;
            }
            break;
        case WM_KEYUP:
            if (wParam == VK_SPACE || wParam == VK_RETURN) {
                owner->HandleNewTabButtonKeyUp(hwnd, static_cast<UINT>(wParam));
                return 0;
            }
            break;
        case WM_THEMECHANGED:
            owner->UpdateNewTabButtonTheme();
            return 0;
        case WM_ENABLE:
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case WM_GETDLGCODE:
            return DLGC_BUTTON | DLGC_UNDEFPUSHBUTTON;
        case WM_SETCURSOR:
            if (LOWORD(lParam) == HTCLIENT) {
                SetCursor(LoadCursor(nullptr, IDC_HAND));
                return TRUE;
            }
            break;
        default:
            break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

bool TabBandWindow::IsSystemDarkMode() const {
    if (m_highContrast) {
        return ComputeLuminance(GetSysColor(COLOR_WINDOW)) < 0.4;
    }
    if (m_themeColors.valid) {
        return ComputeLuminance(m_themeColors.background) < 0.5;
    }

    return IsAppDarkModePreferred();
}

void TabBandWindow::DrawGroupHeader(HDC dc, const VisualItem& item) const {
    RECT rect = item.bounds;
    RECT indicator = rect;
    indicator.right = std::min(indicator.left + kIslandIndicatorWidth, indicator.right);
    indicator.top = rect.top;
    indicator.bottom = rect.bottom;
    if (indicator.right > indicator.left) {
        COLORREF indicatorColor = item.data.hasCustomOutline ? item.data.outlineColor : m_accentColor;
        if (m_highContrast && !item.data.hasCustomOutline) {
            indicatorColor = GetSysColor(COLOR_WINDOWTEXT);
        }
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
    const int paddedHeight = height - kCloseButtonVerticalPadding * 2;
    if (paddedHeight <= 0) {
        return rect;
    }
    const int badgeWidth = std::max(0, item.badgeWidth);
    const int availableWidth = item.bounds.right - item.bounds.left;
    const int targetSize =
        (m_closeButtonSizeCached && m_cachedCloseButtonSize > 0) ? m_cachedCloseButtonSize : kCloseButtonSize;
    const int size = std::min(targetSize, paddedHeight);
    if (size <= 0) {
        return rect;
    }
    const int minimumWidth = size + kCloseButtonEdgePadding + kCloseButtonSpacing + badgeWidth + kPaddingX + 8;
    if (availableWidth < minimumWidth) {
        return rect;
    }
    const int right = item.bounds.right - kCloseButtonEdgePadding;
    const int left = right - size;
    const int top = item.bounds.top + (height - size) / 2;
    rect = {left, top, right, top + size};
    return rect;
}

void TabBandWindow::DrawPinnedGlyph(HDC dc, const RECT& tabRect, int x, COLORREF color) const {
    const int top = static_cast<int>(tabRect.top);
    const int bottom = static_cast<int>(tabRect.bottom);
    const int availableHeight = bottom - top;
    if (availableHeight <= 4) {
        return;
    }

    const int headRadius = std::max(2, std::min(kPinnedGlyphWidth / 2, availableHeight / 5));
    const int headCenter = top + (availableHeight / 2);
    int headTop = headCenter - headRadius;
    int headBottom = headCenter + headRadius;
    headTop = std::max(headTop, top + 1);
    headBottom = std::min(headBottom, bottom - 2);

    const int maxStem = std::max(1, bottom - headBottom - 2);
    int stemLength = std::min(std::max(headRadius, availableHeight / 3), maxStem);
    if (stemLength < 1) {
        stemLength = std::min(maxStem, 1);
    }

    const int baseHalf = std::max(1, headRadius);
    const int triangleHeight = std::max(1, headRadius / 2);
    const int tipY = headBottom - 1 + stemLength;
    const int triangleBottom = std::min(bottom - 1, tipY + triangleHeight);

    HPEN pen = CreatePen(PS_SOLID, 1, color);
    HBRUSH brush = CreateSolidBrush(color);
    if (!pen || !brush) {
        if (pen) {
            DeleteObject(pen);
        }
        if (brush) {
            DeleteObject(brush);
        }
        return;
    }

    HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
    HBRUSH oldBrush = static_cast<HBRUSH>(SelectObject(dc, brush));

    Ellipse(dc, x, headTop, x + kPinnedGlyphWidth, headBottom);

    MoveToEx(dc, x + kPinnedGlyphWidth / 2, headBottom - 1, nullptr);
    LineTo(dc, x + kPinnedGlyphWidth / 2, tipY);

    POINT triangle[3] = {
        {x + kPinnedGlyphWidth / 2, tipY},
        {x + kPinnedGlyphWidth / 2 - baseHalf, triangleBottom},
        {x + kPinnedGlyphWidth / 2 + baseHalf, triangleBottom},
    };
    Polygon(dc, triangle, ARRAYSIZE(triangle));

    SelectObject(dc, oldBrush);
    SelectObject(dc, oldPen);
    DeleteObject(brush);
    DeleteObject(pen);
}

TabBandWindow::TabPaintMetrics TabBandWindow::ComputeTabPaintMetrics(const VisualItem& item) const {
    TabPaintMetrics metrics;
    metrics.itemBounds = item.bounds;
    metrics.tabBounds = item.bounds;
    metrics.islandIndicator = item.indicatorHandle ? kIslandIndicatorWidth : 0;
    metrics.tabBounds.left += metrics.islandIndicator;
    metrics.closeButton = ComputeCloseButtonRect(item);
    metrics.iconLeft = metrics.itemBounds.left + metrics.islandIndicator + kPaddingX;
    metrics.textLeft = metrics.iconLeft;
    metrics.textRight = metrics.itemBounds.right - kPaddingX;

    if (metrics.closeButton.right > metrics.closeButton.left) {
        const int closeLeft = static_cast<int>(metrics.closeButton.left);
        metrics.textRight = std::min(metrics.textRight, closeLeft - kCloseButtonSpacing);
    }

    if (item.icon) {
        const int availableHeight = metrics.itemBounds.bottom - metrics.itemBounds.top;
        metrics.iconHeight = std::min(item.iconHeight, availableHeight - 4);
        metrics.iconWidth = item.iconWidth;
        metrics.textLeft += metrics.iconWidth + kIconGap;
    }

    if (metrics.textRight < metrics.textLeft) {
        metrics.textRight = metrics.textLeft;
    }

    return metrics;
}

void TabBandWindow::DrawTab(HDC dc, const VisualItem& item) const {
    const TabPaintMetrics metrics = ComputeTabPaintMetrics(item);
    RECT rect = metrics.itemBounds;
    RECT tabRect = metrics.tabBounds;
    const bool selected = item.data.selected;
    const TabViewItem* indicatorSource = item.hasGroupHeader ? &item.groupHeader : nullptr;
    const bool hasAccent = item.data.hasCustomOutline ||
                           (indicatorSource && indicatorSource->hasCustomOutline);
    COLORREF accentColor = hasAccent ? ResolveIndicatorColor(indicatorSource, item.data) : m_accentColor;

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
        if (m_highContrast) {
            RECT fillRect = tabRect;
            fillRect.bottom = std::min(fillRect.bottom, rect.bottom - 1);
            if (HBRUSH brush = GetCachedBrush(backgroundColor)) {
                FillRect(dc, &fillRect, brush);
            }
            if (HPEN pen = GetCachedPen(GetSysColor(COLOR_WINDOWTEXT))) {
                SelectObjectGuard penGuard(dc, pen);
                MoveToEx(dc, fillRect.left, fillRect.top, nullptr);
                LineTo(dc, fillRect.right, fillRect.top);
                LineTo(dc, fillRect.right, fillRect.bottom);
                LineTo(dc, fillRect.left, fillRect.bottom);
                LineTo(dc, fillRect.left, fillRect.top);
            }
        } else {
            COLORREF baseBorder = m_darkMode
                                      ? BlendColors(backgroundColor, RGB(255, 255, 255), selected ? 0.1 : 0.05)
                                      : BlendColors(backgroundColor, RGB(0, 0, 0), selected ? 0.15 : 0.1);
            COLORREF borderColor = hasAccent
                                       ? BlendColors(accentColor, RGB(0, 0, 0), selected ? 0.25 : 0.15)
                                       : baseBorder;

            RECT shapeRect = tabRect;
            const LONG bottomLimit = rect.bottom - 1;
            if (shapeRect.bottom > bottomLimit) {
                shapeRect.bottom = bottomLimit;
            }

            const int radius = kTabCornerRadius;
            POINT points[] = {
                {shapeRect.left, shapeRect.bottom},
                {shapeRect.left, shapeRect.top + radius},
                {shapeRect.left + radius, shapeRect.top},
                {shapeRect.right - radius, shapeRect.top},
                {shapeRect.right, shapeRect.top + radius},
                {shapeRect.right, shapeRect.bottom},
            };

            HRGN region = CreatePolygonRgn(points, ARRAYSIZE(points), WINDING);
            if (region) {
                if (HBRUSH brush = GetCachedBrush(backgroundColor)) {
                    FillRgn(dc, region, brush);
                }
                if (HPEN pen = GetCachedPen(borderColor)) {
                    SelectObjectGuard penGuard(dc, pen);
                    if (HGDIOBJ hollowBrush = GetStockObject(HOLLOW_BRUSH)) {
                        SelectObjectGuard brushGuard(dc, hollowBrush);
                        Polygon(dc, points, ARRAYSIZE(points));
                    } else {
                        Polygon(dc, points, ARRAYSIZE(points));
                    }
                }
                DeleteObject(region);
            }

            COLORREF bottomLineColor = selected ? backgroundColor
                                                : (m_darkMode ? BlendColors(backgroundColor, RGB(0, 0, 0), 0.25)
                                                              : GetSysColor(COLOR_3DLIGHT));
            if (HPEN bottomPen = GetCachedPen(bottomLineColor)) {
                SelectObjectGuard penGuard(dc, bottomPen);
                MoveToEx(dc, tabRect.left + 1, rect.bottom - 1, nullptr);
                LineTo(dc, rect.right - 1, rect.bottom - 1);
            }
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
        if (m_highContrast) {
            indicatorColor = hasAccent ? accentColor : GetSysColor(COLOR_WINDOWTEXT);
        }
        if (selected) {
            indicatorColor = DarkenColor(indicatorColor, 0.2);
        }
        if (HBRUSH indicatorBrush = GetCachedBrush(indicatorColor)) {
            FillRect(dc, &indicatorRect, indicatorBrush);
        }
    }

    RECT closeRect = metrics.closeButton;

    int textLeft = metrics.textLeft;
    const int textRight = metrics.textRight;
    if (item.data.pinned) {
        DrawPinnedGlyph(dc, tabRect, textLeft, textColor);
        textLeft += kPinnedGlyphWidth + kPinnedGlyphPadding;
    }
  
    if (item.icon) {
        const int availableHeight = rect.bottom - rect.top;
        const int iconHeight = std::min(metrics.iconHeight, availableHeight - 4);
        const int iconWidth = metrics.iconWidth;
        const int iconY = rect.top + (availableHeight - iconHeight) / 2;
        DrawIconEx(dc, metrics.iconLeft, iconY, item.icon.Get(), iconWidth, iconHeight, 0, nullptr, DI_NORMAL);
    }

    const bool hasProgress = item.data.progress.visible;
    if (hasProgress) {
        DrawTabProgress(dc, item, metrics, computedBackground);
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
            if (m_highContrast) {
                closeBackground = computedBackground;
            }
            if (closePressed) {
                closeBackground = BlendColors(closeBackground, RGB(0, 0, 0), 0.2);
            }

            if (HBRUSH closeBrush = GetCachedBrush(closeBackground)) {
                FillRect(dc, &closeRect, closeBrush);
            }

            COLORREF borderColor = closeHot ? BlendColors(closeBackground, RGB(0, 0, 0), 0.2)
                                            : BlendColors(closeBackground, RGB(0, 0, 0), m_darkMode ? 0.6 : 0.4);
            if (m_highContrast) {
                borderColor = GetSysColor(COLOR_WINDOWTEXT);
            }
            if (HPEN borderPen = GetCachedPen(borderColor)) {
                SelectObjectGuard penGuard(dc, borderPen);
                MoveToEx(dc, closeRect.left, closeRect.top, nullptr);
                LineTo(dc, closeRect.right, closeRect.top);
                LineTo(dc, closeRect.right, closeRect.bottom);
                LineTo(dc, closeRect.left, closeRect.bottom);
                LineTo(dc, closeRect.left, closeRect.top);
            }

            RECT glyphRect = closeRect;
            COLORREF glyphColor = closeHot ? RGB(255, 255, 255) : ResolveTextColor(closeBackground);
            if (m_highContrast) {
                glyphColor = GetSysColor(COLOR_WINDOWTEXT);
                if (closeHot) {
                    glyphColor = GetSysColor(COLOR_HIGHLIGHTTEXT);
                }
            }
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


void TabBandWindow::DrawTabProgress(HDC dc, const VisualItem& item, const TabPaintMetrics& metrics,
                                    COLORREF background) const {
    if (!dc) {
        return;
    }

    RECT outer{};
    if (!ComputeProgressBounds(item, metrics, &outer)) {
        return;
    }

    const COLORREF trackColor = m_darkMode ? BlendColors(background, RGB(255, 255, 255), 0.2)
                                           : BlendColors(background, RGB(0, 0, 0), 0.15);
    if (HBRUSH trackBrush = GetCachedBrush(trackColor)) {
        FillRect(dc, &outer, trackBrush);
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
            if (HBRUSH brush = GetCachedBrush(m_progressEndColor)) {
                FillRect(dc, &segmentRect, brush);
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
    if (HPEN pen = GetCachedPen(borderColor)) {
        SelectObjectGuard penGuard(dc, pen);
        MoveToEx(dc, outer.left, outer.top, nullptr);
        LineTo(dc, outer.right, outer.top);
        LineTo(dc, outer.right, outer.bottom);
        LineTo(dc, outer.left, outer.bottom);
        LineTo(dc, outer.left, outer.top);
    }
}

bool TabBandWindow::ComputeProgressBounds(const VisualItem& item, const TabPaintMetrics& metrics, RECT* out) const {
    if (!out) {
        return false;
    }
    *out = {};
    if (item.data.type != TabViewItemType::kTab || !item.data.progress.visible) {
        return false;
    }
    RECT bounds{metrics.textLeft, std::max(metrics.tabBounds.top + 4, metrics.tabBounds.bottom - 6),
                metrics.textRight, metrics.tabBounds.bottom - 2};
    if (!RectHasArea(bounds)) {
        return false;
    }
    *out = bounds;
    return true;
}

void TabBandWindow::EnsureProgressRectCache() {
    if (m_progressRects.size() != m_items.size()) {
        m_progressRects.assign(m_items.size(), RECT{});
    }
}

void TabBandWindow::RebuildProgressRectCache() {
    EnsureProgressRectCache();
    for (size_t i = 0; i < m_items.size(); ++i) {
        RECT rect{};
        const auto& item = m_items[i];
        if (item.data.type == TabViewItemType::kTab) {
            const TabPaintMetrics metrics = ComputeTabPaintMetrics(item);
            if (ComputeProgressBounds(item, metrics, &rect)) {
                m_progressRects[i] = rect;
                continue;
            }
        }
        m_progressRects[i] = RECT{};
    }
}

void TabBandWindow::RecomputeActiveProgressCount() {
    size_t count = 0;
    for (const auto& item : m_tabData) {
        if (item.progress.visible) {
            ++count;
        }
    }
    m_activeProgressCount = count;
}

void TabBandWindow::InvalidateProgressForIndices(const std::vector<size_t>& indices) {
    if (!m_hwnd || indices.empty()) {
        return;
    }
    EnsureProgressRectCache();
    for (size_t index : indices) {
        if (index >= m_items.size()) {
            continue;
        }
        RECT previous = m_progressRects[index];
        RECT current{};
        const auto& item = m_items[index];
        if (item.data.type == TabViewItemType::kTab) {
            const TabPaintMetrics metrics = ComputeTabPaintMetrics(item);
            if (!ComputeProgressBounds(item, metrics, &current)) {
                current = RECT{};
            }
        }

        RECT dirty{};
        bool hasDirty = false;
        if (RectHasArea(previous)) {
            dirty = previous;
            hasDirty = true;
        }
        if (RectHasArea(current)) {
            if (hasDirty) {
                RECT combined{};
                if (UnionRect(&combined, &dirty, &current)) {
                    dirty = combined;
                } else {
                    dirty.left = std::min(dirty.left, current.left);
                    dirty.top = std::min(dirty.top, current.top);
                    dirty.right = std::max(dirty.right, current.right);
                    dirty.bottom = std::max(dirty.bottom, current.bottom);
                }
            } else {
                dirty = current;
                hasDirty = true;
            }
        }

        if (hasDirty) {
            InvalidateRect(m_hwnd, &dirty, FALSE);
        }
        m_progressRects[index] = current;
    }
}

void TabBandWindow::InvalidateActiveProgress() {
    if (!m_hwnd) {
        return;
    }
    EnsureProgressRectCache();
    m_activeProgressIndices.clear();
    m_activeProgressIndices.reserve(m_items.size());
    for (size_t i = 0; i < m_items.size(); ++i) {
        if (m_items[i].data.type != TabViewItemType::kTab) {
            continue;
        }
        if (m_items[i].data.progress.visible) {
            m_activeProgressIndices.push_back(i);
        }
    }
    InvalidateProgressForIndices(m_activeProgressIndices);
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

    if (HPEN pen = GetCachedPen(m_accentColor, 2)) {
        SelectObjectGuard penGuard(dc, pen);
        const int x = indicator->indicatorX;
        MoveToEx(dc, x, m_clientRect.top + 2, nullptr);
        LineTo(dc, x, m_clientRect.bottom - 2);
    }
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

IconCache::Reference TabBandWindow::LoadItemIcon(const TabViewItem& item, UINT iconFlags) const {
    if (item.type != TabViewItemType::kTab) {
        return {};
    }

    const UINT resolvedFlags = (iconFlags & (SHGFI_LARGEICON | SHGFI_SMALLICON)) != 0
                                   ? (iconFlags & (SHGFI_LARGEICON | SHGFI_SMALLICON))
                                   : SHGFI_SMALLICON;
    std::wstring path = item.path;
    if (path.empty() && item.pidl) {
        path = GetParsingName(item.pidl);
    }
    const std::wstring familyKey = BuildIconCacheFamilyKey(item.pidl, item.path.empty() ? path : item.path);
    PCIDLIST_ABSOLUTE pidl = item.pidl;

    auto loader = [pidl, path, resolvedFlags]() -> HICON {
        SHFILEINFOW info{};
        const UINT flags = SHGFI_ICON | SHGFI_ADDOVERLAYS | resolvedFlags;
        if (pidl) {
            if (SHGetFileInfoW(reinterpret_cast<PCWSTR>(pidl), 0, &info, sizeof(info), flags | SHGFI_PIDL)) {
                return info.hIcon;
            }
        }
        if (!path.empty()) {
            if (SHGetFileInfoW(path.c_str(), 0, &info, sizeof(info), flags)) {
                return info.hIcon;
            }
        }
        return nullptr;
    };

    return IconCache::Instance().Acquire(familyKey, resolvedFlags, loader);
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
    if (!hit.hit || hit.type != HitType::kTab || hit.itemIndex != m_previewItemIndex) {
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
    if (!hit.hit || hit.type != HitType::kTab || hit.itemIndex >= m_items.size()) {
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
        IconCache::Reference icon = LoadItemIcon(visual.data, SHGFI_LARGEICON);
        overlayShown = m_previewOverlay.Show(m_hwnd, nullptr, kPreviewImageSize, screenPt, placeholderText.c_str(),
                                             icon.Get());
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
    RefreshProgressState({}, nullptr);
}

void TabBandWindow::RefreshProgressState(const std::vector<TabLocation>& prioritizedTabs) {
    RefreshProgressState(prioritizedTabs, nullptr);
}

void TabBandWindow::RefreshProgressState(const TabProgressUpdatePayload* payload) {
    RefreshProgressState({}, payload);
}

void TabBandWindow::RefreshProgressState(const std::vector<TabLocation>& prioritizedTabs,
                                         const TabProgressUpdatePayload* payload) {
    auto* manager = ResolveManager();
    if (!manager) {
        if (!m_tabData.empty()) {
            SetTabs({});
            m_tabLayoutVersion = 0;
            UpdateProgressAnimationState();
        }
        return;
    }

    const uint32_t layoutVersion = manager->GetLayoutVersion();

    TabProgressUpdatePayload synthesized;
    if (!payload && !prioritizedTabs.empty() && layoutVersion == m_tabLayoutVersion) {
        synthesized.layoutVersion = layoutVersion;
        synthesized.entries.reserve(prioritizedTabs.size());
        for (const auto& location : prioritizedTabs) {
            const auto* tab = manager->Get(location);
            if (!tab || tab->hidden) {
                continue;
            }
            TabProgressSnapshotEntry entry;
            entry.type = TabViewItemType::kTab;
            entry.location = location;
            entry.lastActivatedTick = tab->lastActivatedTick;
            entry.activationOrdinal = tab->activationOrdinal;
            if (tab->progress.active) {
                entry.progress.visible = true;
                entry.progress.indeterminate = tab->progress.indeterminate;
                entry.progress.fraction =
                    tab->progress.indeterminate ? 0.0 : ClampProgress(tab->progress.fraction);
            }
            synthesized.entries.emplace_back(std::move(entry));
        }
        if (!synthesized.entries.empty()) {
            payload = &synthesized;
        }
    }

    auto applyPayload = [&](const TabProgressUpdatePayload* updatePayload) -> bool {
        if (!updatePayload) {
            return false;
        }
        if (updatePayload->layoutVersion != m_tabLayoutVersion) {
            return false;
        }

        bool resyncNeeded = false;
        bool changed = false;
        std::vector<size_t> progressChanged;
        progressChanged.reserve(updatePayload->entries.size());

        for (const auto& entry : updatePayload->entries) {
            size_t index = kInvalidIndex;
            if (entry.type == TabViewItemType::kGroupHeader) {
                index = FindGroupHeaderIndex(entry.location.groupIndex);
            } else {
                index = FindTabDataIndex(entry.location);
            }

            if (index == kInvalidIndex || index >= m_tabData.size() ||
                m_tabData[index].type != entry.type) {
                resyncNeeded = true;
                break;
            }

            auto& data = m_tabData[index];
            bool entryChanged = false;
            if (data.lastActivatedTick != entry.lastActivatedTick ||
                data.activationOrdinal != entry.activationOrdinal) {
                data.lastActivatedTick = entry.lastActivatedTick;
                data.activationOrdinal = entry.activationOrdinal;
                entryChanged = true;
            }

            if (data.progress != entry.progress) {
                data.progress = entry.progress;
                if (entry.type == TabViewItemType::kTab) {
                    progressChanged.push_back(index);
                }
                entryChanged = true;
            }

            if (index < m_items.size()) {
                auto& visual = m_items[index].data;
                if (visual.lastActivatedTick != entry.lastActivatedTick ||
                    visual.activationOrdinal != entry.activationOrdinal) {
                    visual.lastActivatedTick = entry.lastActivatedTick;
                    visual.activationOrdinal = entry.activationOrdinal;
                    entryChanged = true;
                }
                if (visual.progress != entry.progress) {
                    visual.progress = entry.progress;
                }
            }

            changed |= entryChanged;
        }

        if (resyncNeeded) {
            return false;
        }

        std::vector<size_t> priorityIndices;
        priorityIndices.reserve(prioritizedTabs.size());
        for (const auto& location : prioritizedTabs) {
            const size_t index = FindTabDataIndex(location);
            if (index != kInvalidIndex) {
                priorityIndices.push_back(index);
            }
        }

        for (size_t index : priorityIndices) {
            if (std::find(progressChanged.begin(), progressChanged.end(), index) ==
                progressChanged.end()) {
                progressChanged.push_back(index);
            }
        }

        if (!progressChanged.empty()) {
            InvalidateProgressForIndices(progressChanged);
        } else if (changed && m_hwnd) {
            InvalidateRect(m_hwnd, nullptr, FALSE);
        }

        UpdateProgressAnimationState();
        return true;
    };

    if (applyPayload(payload)) {
        return;
    }

    if (layoutVersion != m_tabLayoutVersion) {
        SetTabs(manager->BuildView());
        m_tabLayoutVersion = layoutVersion;
        UpdateProgressAnimationState();
        return;
    }

    const auto snapshot = manager->CollectProgressStates();
    bool layoutMismatch = snapshot.size() != m_tabData.size();
    if (!layoutMismatch) {
        for (size_t i = 0; i < snapshot.size(); ++i) {
            if (m_tabData[i].type != snapshot[i].type ||
                m_tabData[i].location.groupIndex != snapshot[i].location.groupIndex ||
                m_tabData[i].location.tabIndex != snapshot[i].location.tabIndex) {
                layoutMismatch = true;
                break;
            }
            if (snapshot[i].type == TabViewItemType::kTab) {
                const auto* tab = manager->Get(snapshot[i].location);
                if (!tab || !ArePidlsEqual(tab->pidl.get(), m_tabData[i].pidl)) {
                    layoutMismatch = true;
                    break;
                }
            }
        }
    }
    if (layoutMismatch) {
        SetTabs(manager->BuildView());
        m_tabLayoutVersion = layoutVersion;
        UpdateProgressAnimationState();
        return;
    }

    std::vector<size_t> priorityIndices;
    priorityIndices.reserve(prioritizedTabs.size());
    for (const auto& location : prioritizedTabs) {
        const size_t index = FindTabDataIndex(location);
        if (index != kInvalidIndex) {
            priorityIndices.push_back(index);
        }
    }

    bool changed = false;
    std::vector<size_t> progressChanged;
    progressChanged.reserve(snapshot.size());
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
            const bool wasVisible = m_tabData[i].progress.visible;
            const bool nowVisible = snapshot[i].progress.visible;
            if (wasVisible != nowVisible) {
                if (nowVisible) {
                    ++m_activeProgressCount;
                } else if (m_activeProgressCount > 0) {
                    --m_activeProgressCount;
                }
            }
            m_tabData[i].progress = snapshot[i].progress;
            if (i < m_items.size()) {
                m_items[i].data.progress = snapshot[i].progress;
            }
            progressChanged.push_back(i);
            changed = true;
        }
    }

    for (size_t index : priorityIndices) {
        if (index < snapshot.size() &&
            std::find(progressChanged.begin(), progressChanged.end(), index) ==
                progressChanged.end()) {
            progressChanged.push_back(index);
        }
    }

    UpdateProgressAnimationState();

    if (!progressChanged.empty()) {
        InvalidateProgressForIndices(progressChanged);
    } else if (changed && m_hwnd) {
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
    return m_activeProgressCount > 0;
}

void TabBandWindow::HandleProgressTimer() {
    if (!m_hwnd) {
        return;
    }
    const ULONGLONG now = GetTickCount64();
    if (auto* manager = ResolveManager(); manager) {
        const auto expired = manager->ExpireFolderOperations(now, kProgressStaleTimeoutMs);
        if (!expired.empty()) {
            RefreshProgressState(expired);
            return;
        }
    }
    if (!AnyProgressActive()) {
        UpdateProgressAnimationState();
        return;
    }
    InvalidateActiveProgress();
}

TabManager* TabBandWindow::ResolveManager() const noexcept {
    return m_owner ? &m_owner->GetTabManager() : nullptr;
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
    auto* manager = ResolveManager();
    if (!manager) {
        return;
    }
    const LONG eventId = static_cast<LONG>(wParam) & 0xFFFF;
    auto touch = [manager](PCIDLIST_ABSOLUTE pidl) {
        if (!pidl) {
            return;
        }
        if (auto parent = CloneParent(pidl)) {
            manager->TouchFolderOperation(parent.get());
        } else {
            manager->TouchFolderOperation(pidl);
        }
    };
    auto clear = [manager](PCIDLIST_ABSOLUTE pidl) {
        if (!pidl) {
            return;
        }
        if (auto parent = CloneParent(pidl)) {
            manager->ClearFolderOperation(parent.get());
        } else {
            manager->ClearFolderOperation(pidl);
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
        if (hit.hit && hit.type == HitType::kTab && hit.itemIndex < m_items.size()) {
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

void TabBandWindow::HandleCommand(WPARAM wParam, LPARAM lParam) {
        if (!m_owner) {
                return;
        }

        const UINT id = LOWORD(wParam);
        const UINT code = HIWORD(wParam);

        if (id == IDC_NEW_TAB) {
                const HWND source = reinterpret_cast<HWND>(lParam);
                const bool fromNewTabButton = source == m_newTabButton;
                if (fromNewTabButton || (!source && (code == BN_CLICKED || code == 0))) {
                        RequestNewTab();
                }
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
                                m_owner->OnShowOptionsDialog(OptionsDialogPage::kGroups, focusId);
                        } else {
                                m_owner->OnShowOptionsDialog(OptionsDialogPage::kGroups);
                        }
                }
                ClearExplorerContext();
                return;
        }

        if (id == IDM_CONTEXT_MENU_CUSTOMIZATIONS) {
                if (m_owner) {
                        m_owner->OnShowOptionsDialog(OptionsDialogPage::kContextMenus);
                }
                ClearExplorerContext();
                return;
        }

        if (id == IDM_OPTIONS) {
                if (m_owner) {
                        m_owner->OnShowOptionsDialog(OptionsDialogPage::kGeneral);
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

        case IDM_TOGGLE_PIN_TAB:
                if (m_contextHit.location.IsValid()) {
                        m_owner->OnToggleTabPinned(m_contextHit.location);
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
    if (!hit.hit || hit.type == HitType::kWhitespace || hit.type == HitType::kNewTab) {
        return false;
    }

    SetFocus(m_hwnd);
    HideDragOverlay(true);
    DropTarget clearedDropTarget{};
    ApplyInternalDropTarget(m_drag.target, clearedDropTarget);
    m_drag = {};
    if (hit.closeButton && hit.type == HitType::kTab) {
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
                        DropTarget previous = m_drag.target;
                        DropTarget outside{};
                        outside.active = true;
                        outside.outside = true;
                        ApplyInternalDropTarget(previous, outside);
                }

                CompleteDrop();  // this typically finalizes the move
        }
	// Simple tracking release → click selection
	else if (m_drag.tracking) {
		handled = true;
		const HitInfo hit = HitTest(pt);
                if (hit.hit && hit.type != HitType::kWhitespace && hit.type != HitType::kNewTab) {
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
            DropTarget previous = m_drag.target;
            DropTarget outside{};
            outside.active = true;
            outside.outside = true;
            ApplyInternalDropTarget(previous, outside);
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
    if (!hit.hit || hit.type == HitType::kWhitespace || hit.type == HitType::kNewTab) {
        return false;
    }

    if (hit.closeButton) {
        return false;
    }

    if (hit.type == HitType::kGroupHeader) {
        m_owner->OnToggleGroupCollapsed(hit.location.groupIndex);
        return true;
    }
    if (hit.type == HitType::kTab && hit.location.IsValid()) {
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
    const bool dropOnTab = hit.hit && hit.type == HitType::kTab && hit.location.IsValid();
    const bool dropOnWhitespace = hit.hit && (hit.type == HitType::kWhitespace || hit.type == HitType::kNewTab);
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    bool handled = false;

    if (dropOnTab && !paths.empty()) {
        const bool move = shift;
        LogMessage(LogLevel::Info, L"HDROP delegated to tab (%d,%d); move=%d, items=%u", hit.location.groupIndex,
                   hit.location.tabIndex, move ? 1 : 0, static_cast<unsigned>(paths.size()));
        m_owner->OnFilesDropped(hit.location, paths, move);
        handled = true;
    }

    if (!handled && dropOnWhitespace && !paths.empty()) {
        std::vector<std::wstring> directoryPaths;
        directoryPaths.reserve(paths.size());
        for (const auto& path : paths) {
            const DWORD attributes = GetFileAttributesW(path.c_str());
            if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
                directoryPaths.push_back(path);
            }
        }

        TabManager* manager = ResolveManager();
        auto resolveFallbackLocation = [&]() -> std::optional<TabLocation> {
            if (!manager) {
                return std::nullopt;
            }
            TabLocation selected = manager->SelectedLocation();
            if (selected.IsValid()) {
                return selected;
            }
            const VisualItem* nearest = nullptr;
            int bestDistance = std::numeric_limits<int>::max();
            for (const auto& item : m_items) {
                if (item.data.type != TabViewItemType::kTab || !item.data.location.IsValid()) {
                    continue;
                }
                int distance = 0;
                if (pt.x < item.bounds.left) {
                    distance = item.bounds.left - pt.x;
                } else if (pt.x > item.bounds.right) {
                    distance = pt.x - item.bounds.right;
                }
                if (distance < bestDistance) {
                    bestDistance = distance;
                    nearest = &item;
                }
            }
            if (nearest) {
                return nearest->data.location;
            }
            return std::nullopt;
        };

        auto openDirectoryTabs = [&](bool foreground) -> size_t {
            if (!m_owner) {
                return 0;
            }
            size_t opened = 0;
            bool openForeground = foreground;
            for (const auto& directory : directoryPaths) {
                const bool selectTab = openForeground && opened == 0;
                m_owner->OnOpenFolderInNewTab(directory, selectTab);
                ++opened;
                if (selectTab) {
                    openForeground = false;
                }
            }
            return opened;
        };

        const bool preferTabs = !shift && !directoryPaths.empty();
        if (preferTabs) {
            size_t opened = openDirectoryTabs(ctrl);
            if (opened > 0) {
                LogMessage(LogLevel::Info, L"HDROP opened %Iu tab(s) from whitespace drop (foreground=%d)", opened,
                           ctrl ? 1 : 0);
                handled = true;
            }
        }

        if (!handled) {
            std::optional<TabLocation> fallback = resolveFallbackLocation();
            if (fallback && fallback->IsValid()) {
                const bool move = shift;
                LogMessage(LogLevel::Info, L"HDROP whitespace fallback to tab (%d,%d); move=%d, items=%u",
                           fallback->groupIndex, fallback->tabIndex, move ? 1 : 0,
                           static_cast<unsigned>(paths.size()));
                m_owner->OnFilesDropped(*fallback, paths, move);
                handled = true;
            } else if (!directoryPaths.empty()) {
                size_t opened = openDirectoryTabs(false);
                if (opened > 0) {
                    LogMessage(LogLevel::Info,
                               L"HDROP opened %Iu tab(s) from whitespace drop without fallback (shift=%d)", opened,
                               shift ? 1 : 0);
                    handled = true;
                }
            } else {
                LogMessage(LogLevel::Warning,
                           L"HDROP whitespace drop ignored (no directories and no fallback target)");
            }
        }
    }

    if (!handled && !dropOnTab && !dropOnWhitespace && !HasAnyTabs() && m_owner) {
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
    if (!hit.hit || hit.type != HitType::kTab || !hit.location.IsValid()) {
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

    const bool eligible = hasFileData && hit.hit && hit.type == HitType::kTab && hit.location.IsValid() &&
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
    if (!m_dropHoverHit.hit || m_dropHoverHit.type != HitType::kTab || !m_dropHoverHit.location.IsValid()) {
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
    DropTarget previousExternal = m_externalDrop.active ? m_externalDrop.target : DropTarget{};
    DropTarget previousDrag = m_drag.target;
    DropTarget cleared{};
    ApplyExternalDropTarget(previousExternal, cleared, nullptr);
    ApplyInternalDropTarget(previousDrag, cleared);
    m_externalDrop = {};
    m_drag = {};
    m_mouseTracking = false;
    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, FALSE);
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
    if (!hit.hit || hit.type == HitType::kWhitespace || hit.type == HitType::kNewTab) {
        auto trailingIndicatorX = [&]() -> LONG {
            if (m_newTabBounds.right > m_newTabBounds.left) {
                return m_newTabBounds.left;
            }
            return m_clientRect.right - 10;
        };
        if (origin.type == HitType::kTab && m_owner) {
            target.group = false;
            target.newGroup = true;
            target.floating = true;
            target.groupIndex = m_owner->GetGroupCount();
            target.tabIndex = 0;
            target.indicatorX = trailingIndicatorX();
        } else if (!m_items.empty()) {
            const VisualItem* lastHeader = FindLastGroupHeader();
            if (lastHeader) {
                if (origin.type == HitType::kGroupHeader) {
                    target.group = true;
                    target.groupIndex = lastHeader->data.location.groupIndex + 1;
                    target.indicatorX = std::min<LONG>(lastHeader->bounds.right, trailingIndicatorX());
                } else {
                    target.group = false;
                    target.groupIndex = lastHeader->data.location.groupIndex;
                    target.tabIndex = static_cast<int>(lastHeader->data.totalTabs);
                    target.indicatorX = std::min<LONG>(lastHeader->bounds.right, trailingIndicatorX());
                }
            } else {
                const auto& tail = m_items.back();
                target.group = false;
                target.groupIndex = tail.data.location.groupIndex;
                target.tabIndex = tail.data.location.tabIndex + 1;
                target.indicatorX = std::min<LONG>(tail.bounds.right, trailingIndicatorX());
            }
        }
        return target;
    }

    const VisualItem& visual = m_items[hit.itemIndex];
    const int midX = (visual.bounds.left + visual.bounds.right) / 2;
    const bool leftSide = pt.x < midX;

    if (origin.type == HitType::kGroupHeader) {
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

    AdjustDropTargetForPinned(origin, target);
    return target;
}

int TabBandWindow::ComputeIndicatorXForInsertion(int groupIndex, int tabIndex) const {
    const VisualItem* header = nullptr;
    const VisualItem* previous = nullptr;
    const VisualItem* next = nullptr;
    for (const auto& visual : m_items) {
        if (visual.data.type == TabViewItemType::kGroupHeader) {
            if (visual.data.location.groupIndex == groupIndex) {
                header = &visual;
            }
            continue;
        }
        if (visual.data.location.groupIndex != groupIndex) {
            continue;
        }
        if (visual.data.location.tabIndex >= tabIndex) {
            next = &visual;
            break;
        }
        previous = &visual;
    }

    if (next) {
        return next->bounds.left;
    }
    if (previous) {
        return previous->bounds.right;
    }
    if (header) {
        return header->bounds.right;
    }

    const LONG left = m_clientRect.left;
    const LONG right = m_clientRect.right;
    const LONG fallback = static_cast<LONG>(m_clientRect.left + m_toolbarGripWidth);
    const LONG low = std::min(left, right);
    const LONG high = std::max(left, right);
    const LONG clamped = std::clamp(fallback, low, high);
    return static_cast<int>(clamped);
}

void TabBandWindow::AdjustDropTargetForPinned(const HitInfo& origin, DropTarget& target) const {
    if (!target.active || target.group || target.newGroup || target.outside) {
        return;
    }
    if (!origin.location.IsValid()) {
        return;
    }
    auto* manager = ResolveManager();
    if (!manager) {
        return;
    }
    const TabInfo* moving = manager->Get(origin.location);
    if (!moving) {
        return;
    }
    const TabGroup* destination = manager->GetGroup(target.groupIndex);
    if (!destination) {
        return;
    }

    int destinationSize = static_cast<int>(destination->tabs.size());
    int adjustedIndex = std::clamp(target.tabIndex, 0, destinationSize);

    int pinnedCount = 0;
    for (const auto& tab : destination->tabs) {
        if (!tab.pinned) {
            break;
        }
        ++pinnedCount;
    }

    if (moving->pinned) {
        int effectivePinned = pinnedCount;
        if (origin.location.groupIndex == target.groupIndex && effectivePinned > 0) {
            effectivePinned = std::max(0, effectivePinned - 1);
        }
        adjustedIndex = std::clamp(adjustedIndex, 0, effectivePinned);
    } else {
        int lowerBound = std::min(pinnedCount, destinationSize);
        adjustedIndex = std::max(adjustedIndex, lowerBound);
        adjustedIndex = std::clamp(adjustedIndex, lowerBound, destinationSize);
    }

    target.tabIndex = adjustedIndex;
    const int indicator = ComputeIndicatorXForInsertion(target.groupIndex, adjustedIndex);
    if (indicator >= 0) {
        target.indicatorX = indicator;
    }
}

RECT TabBandWindow::ComputeDropIndicatorRect(const DropTarget& target) const {
    RECT rect{};
    if (!target.active || target.outside || target.indicatorX < 0) {
        return rect;
    }

    rect.left = target.indicatorX;
    rect.right = target.indicatorX + 1;
    rect.top = m_clientRect.top + 2;
    rect.bottom = m_clientRect.bottom - 2;
    if (rect.bottom <= rect.top) {
        rect.top = m_clientRect.top;
        rect.bottom = m_clientRect.bottom;
    }
    InflateRect(&rect, kDropIndicatorHalfWidth, kDropInvalidatePadding);

    RECT clipped{};
    if (ClipRectToClient(rect, m_clientRect, &clipped) && RectHasArea(clipped)) {
        return clipped;
    }
    return RECT{};
}

bool TabBandWindow::TryGetGroupBounds(int groupIndex, RECT* bounds) const {
    if (!bounds || groupIndex < 0) {
        return false;
    }

    bool found = false;
    RECT result{};
    for (const auto& item : m_items) {
        if (item.data.location.groupIndex != groupIndex) {
            continue;
        }
        if (!found) {
            result = item.bounds;
            found = true;
        } else {
            RECT combined{};
            if (UnionRect(&combined, &result, &item.bounds)) {
                result = combined;
            } else {
                result.left = std::min(result.left, item.bounds.left);
                result.top = std::min(result.top, item.bounds.top);
                result.right = std::max(result.right, item.bounds.right);
                result.bottom = std::max(result.bottom, item.bounds.bottom);
            }
        }
    }

    if (!found) {
        return false;
    }

    *bounds = result;
    return true;
}

bool TabBandWindow::TryGetTabBounds(int groupIndex, int tabIndex, RECT* bounds) const {
    if (!bounds || groupIndex < 0 || tabIndex < 0) {
        return false;
    }

    for (const auto& item : m_items) {
        if (item.data.type != TabViewItemType::kTab) {
            continue;
        }
        if (item.data.location.groupIndex == groupIndex && item.data.location.tabIndex == tabIndex) {
            *bounds = item.bounds;
            return true;
        }
    }
    return false;
}

RECT TabBandWindow::ComputeDropPreviewRect(const DropTarget& target) const {
    RECT rect{};
    if (!target.active || target.outside) {
        return rect;
    }

    RECT base{};
    if (target.group) {
        if (!TryGetGroupBounds(target.groupIndex, &base)) {
            return rect;
        }
    } else {
        if (target.newGroup || target.tabIndex < 0) {
            return rect;
        }
        if (!TryGetTabBounds(target.groupIndex, target.tabIndex, &base)) {
            return rect;
        }
    }

    RECT shifted = base;
    OffsetRect(&shifted, kDropPreviewOffset, 0);

    RECT combined{};
    if (!UnionRect(&combined, &base, &shifted)) {
        combined.left = std::min(base.left, shifted.left);
        combined.top = std::min(base.top, shifted.top);
        combined.right = std::max(base.right, shifted.right);
        combined.bottom = std::max(base.bottom, shifted.bottom);
    }

    InflateRect(&combined, kDropInvalidatePadding, kDropInvalidatePadding);

    RECT clipped{};
    if (ClipRectToClient(combined, m_clientRect, &clipped) && RectHasArea(clipped)) {
        return clipped;
    }
    return RECT{};
}

void TabBandWindow::InvalidateDropRegions(const RECT& previousIndicator, const RECT& currentIndicator,
                                          const RECT& previousPreview, const RECT& currentPreview) {
    if (!m_hwnd) {
        return;
    }

    RECT dirty{};
    bool hasDirty = false;
    const auto accumulate = [&](const RECT& rect) {
        if (!RectHasArea(rect)) {
            return;
        }
        RECT clipped{};
        if (!ClipRectToClient(rect, m_clientRect, &clipped) || !RectHasArea(clipped)) {
            return;
        }
        if (!hasDirty) {
            dirty = clipped;
            hasDirty = true;
            return;
        }
        RECT combined{};
        if (UnionRect(&combined, &dirty, &clipped)) {
            dirty = combined;
        } else {
            dirty.left = std::min(dirty.left, clipped.left);
            dirty.top = std::min(dirty.top, clipped.top);
            dirty.right = std::max(dirty.right, clipped.right);
            dirty.bottom = std::max(dirty.bottom, clipped.bottom);
        }
    };

    accumulate(previousIndicator);
    accumulate(currentIndicator);
    accumulate(previousPreview);
    accumulate(currentPreview);

    if (hasDirty) {
        InvalidateRect(m_hwnd, &dirty, FALSE);
    }
}

void TabBandWindow::ApplyDropTargetChange(const DropTarget& previous, const DropTarget& current,
                                          RECT& indicatorRectStorage, RECT& previewRectStorage) {
    OnDropPreviewTargetChanged(previous, current);

    const RECT newIndicator = ComputeDropIndicatorRect(current);
    const RECT newPreview = ComputeDropPreviewRect(current);

    InvalidateDropRegions(indicatorRectStorage, newIndicator, previewRectStorage, newPreview);

    indicatorRectStorage = newIndicator;
    previewRectStorage = newPreview;
}

void TabBandWindow::ApplyInternalDropTarget(const DropTarget& previous, const DropTarget& current) {
    ApplyDropTargetChange(previous, current, m_drag.indicatorRect, m_drag.previewRect);
    m_drag.target = current;
}

void TabBandWindow::ApplyExternalDropTarget(const DropTarget& previous, const DropTarget& current,
                                            TabBandWindow* sourceWindow) {
    ApplyDropTargetChange(previous, current, m_externalDrop.indicatorRect, m_externalDrop.previewRect);
    m_externalDrop.target = current;
    m_externalDrop.active = current.active && !current.outside;
    m_externalDrop.source = m_externalDrop.active ? sourceWindow : nullptr;
}

void TabBandWindow::UpdateDropTarget(const POINT& pt) {
    DropTarget previous = m_drag.target;
    DropTarget target = ComputeDropTarget(pt, m_drag.origin);
    ApplyInternalDropTarget(previous, target);
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

    if (m_drag.origin.type == HitType::kGroupHeader) {
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

    DropTarget previousExternalTarget = m_externalDrop.active ? m_externalDrop.target : DropTarget{};
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
        ApplyExternalDropTarget(previousExternalTarget, target, sourceWindow);
    } else {
        DropTarget cleared{};
        ApplyExternalDropTarget(previousExternalTarget, cleared, nullptr);
        m_externalDrop = {};
    }
}

void TabBandWindow::HandleExternalDragLeave() {
    DropTarget previous = m_externalDrop.active ? m_externalDrop.target : DropTarget{};

    {
        auto& state = GetSharedDragState();
        std::scoped_lock lock(state.mutex);
        if (state.hover == this) {
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
        }
    }
    DropTarget cleared{};
    ApplyExternalDropTarget(previous, cleared, nullptr);
    m_externalDrop = {};
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

    DropTarget previous = m_externalDrop.active ? m_externalDrop.target : DropTarget{};
    DropTarget cleared{};
    ApplyExternalDropTarget(previous, cleared, nullptr);
    m_externalDrop = {};
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
        if (origin.type == HitType::kGroupHeader) {
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

    if (origin.type == HitType::kGroupHeader) {
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

    if (hit.type == HitType::kTab && hit.location.IsValid()) {
        m_owner->OnTabSelected(hit.location);
    } else if (hit.type == HitType::kGroupHeader) {
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
            info.type = (item.data.type == TabViewItemType::kTab) ? HitType::kTab : HitType::kGroupHeader;
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

    if (m_newTabBounds.right > m_newTabBounds.left && PtInRect(&m_newTabBounds, pt)) {
        info.hit = true;
        info.type = HitType::kNewTab;
        info.itemIndex = std::numeric_limits<size_t>::max();
        info.location = {};
        info.before = false;
        info.after = false;
        info.closeButton = false;
        return info;
    }

    info.hit = true;
    info.type = HitType::kWhitespace;
    info.itemIndex = std::numeric_limits<size_t>::max();
    info.location = {};
    info.before = false;
    info.after = false;
    info.closeButton = false;
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

    if (hit.hit && hit.type != HitType::kWhitespace && hit.type != HitType::kNewTab) {
        if (hit.type == HitType::kTab) {
            AppendMenuW(menu, MF_STRING, IDM_CLOSE_TAB, L"Close Tab");
            AppendMenuW(menu, MF_STRING, IDM_HIDE_TAB, L"Hide Tab");
            bool pinned = false;
            if (hitVisual && hitVisual->data.type == TabViewItemType::kTab) {
                pinned = hitVisual->data.pinned;
            } else if (hit.itemIndex < m_items.size()) {
                pinned = m_items[hit.itemIndex].data.pinned;
            } else if (auto* manager = ResolveManager()) {
                if (const auto* tab = manager->Get(hit.location)) {
                    pinned = tab->pinned;
                }
            }
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_PIN_TAB, pinned ? L"Unpin Tab" : L"Pin Tab");
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
        } else if (hit.type == HitType::kGroupHeader && hit.itemIndex >= 0) {
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
        } else if (hit.type == HitType::kGroupHeader && hit.location.groupIndex >= 0) {
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
    AppendMenuW(menu, MF_STRING, IDM_CONTEXT_MENU_CUSTOMIZATIONS, L"Context Menu Customizations...");
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

    m_savedGroupCommands.clear();

    HMENU groupsMenu = CreatePopupMenu();
    if (!groupsMenu) {
        return;
    }

    const auto names = m_owner->GetSavedGroupNames();
    if (!names.has_value()) {
        AppendMenuW(groupsMenu, MF_STRING | MF_GRAYED, 0, L"Failed to load saved groups");
    } else if (names->empty()) {
        AppendMenuW(groupsMenu, MF_STRING | MF_GRAYED, 0, L"No Saved Groups");
    } else {
        UINT command = IDM_LOAD_SAVED_GROUP_BASE;
        for (const auto& name : *names) {
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

size_t TabBandWindow::FindTabDataIndex(TabLocation location) const {
    if (!location.IsValid()) {
        return kInvalidIndex;
    }
    const auto it = m_tabLocationIndex.find(location);
    if (it == m_tabLocationIndex.end()) {
        return kInvalidIndex;
    }
    return it->second;
}

size_t TabBandWindow::FindGroupHeaderIndex(int groupIndex) const {
    if (groupIndex < 0) {
        return kInvalidIndex;
    }
    for (size_t i = 0; i < m_tabData.size(); ++i) {
        const auto& item = m_tabData[i];
        if (item.type == TabViewItemType::kGroupHeader &&
            item.location.groupIndex == groupIndex) {
            return i;
        }
    }
    return kInvalidIndex;
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
    if (!hit.hit || hit.type == HitType::kWhitespace || hit.type == HitType::kNewTab) {
        return nullptr;
    }

    if (hit.type == HitType::kGroupHeader) {
        for (const auto& item : m_items) {
            if (item.data.type == TabViewItemType::kGroupHeader &&
                item.data.location.groupIndex == hit.location.groupIndex) {
                return &item;
            }
        }
    } else if (hit.type == HitType::kTab && hit.location.IsValid()) {
        for (const auto& item : m_items) {
            if (item.data.type == TabViewItemType::kTab && item.data.location.IsValid() &&
                item.data.location.groupIndex == hit.location.groupIndex &&
                item.data.location.tabIndex == hit.location.tabIndex) {
                return &item;
            }
        }
    }

    if (hit.type == HitType::kGroupHeader) {
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
        const UINT savedGroupsMessage = GetSavedGroupsChangedMessage();
        if (savedGroupsMessage != 0 && message == savedGroupsMessage) {
            self->OnSavedGroupsChanged();
            return 0;
        }
        const UINT progressMessage = GetProgressUpdateMessage();
        if (progressMessage != 0 && message == progressMessage) {
            std::unique_ptr<TabProgressUpdatePayload> payload(
                reinterpret_cast<TabProgressUpdatePayload*>(lParam));
            if (payload) {
                self->RefreshProgressState(payload.get());
            } else {
                self->RefreshProgressState();
            }
            return 0;
        }
        if (self->m_shellNotifyMessage != 0 && message == self->m_shellNotifyMessage) {
            self->OnShellNotify(wParam, lParam);
            return 0;
        }
        switch (message) {
            case WM_CREATE: {
                if (EnsureNewTabButtonClassRegistered()) {
                    self->m_newTabButton = CreateWindowExW(0, kNewTabButtonClassName, L"New tab",
                                                           WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                                           0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDC_NEW_TAB),
                                                           GetModuleHandleInstance(), self);
                }
                if (self->m_newTabButton) {
                    self->UpdateNewTabButtonTheme();
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

            case WM_DPICHANGED: {
                auto* suggested = reinterpret_cast<RECT*>(lParam);
                self->HandleDpiChanged(LOWORD(wParam), HIWORD(wParam), suggested);
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
                if (wParam == TabBandWindow::kSessionFlushTimerId) {
                    if (self->m_owner) {
                        self->m_owner->OnPeriodicSessionFlush();
                    }
                    return 0;
                }
                return fallback();
            }
            case WM_WTSSESSION_CHANGE: {
                if (self->m_themeNotifier.HandleSessionChange(wParam, lParam)) {
                    return 0;
                }
                return fallback();
            }
            case WM_SHELLTABS_THEME_CHANGED:
            case WM_THEMECHANGED:
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
            case WM_SHELLTABS_SHOW_HISTORY_MENU: {
                if (!self->m_owner) {
                    return 0;
                }
                const auto* request = reinterpret_cast<const HistoryMenuRequest*>(wParam);
                if (!request) {
                    return 0;
                }
                return self->m_owner->OnShowHistoryMenu(*request) ? 1 : 0;
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
            case WM_SHELLTABS_REGISTER_DRAGDROP: {
                self->m_dropTargetRegistrationPending = false;
                self->EnsureDropTargetRegistered();
                return 0;
            }
            case WM_SHELLTABS_PREVIEW_READY: {
                self->HandlePreviewReady(static_cast<uint64_t>(wParam));
                return 0;
            }
            case WM_SHELLTABS_INITIALIZATION_COMPLETE: {
                auto* payload = reinterpret_cast<TabBand::InitializationResult*>(lParam);
                std::unique_ptr<TabBand::InitializationResult> result(payload);
                if (self->m_owner && result) {
                    self->m_owner->HandleInitializationResult(std::move(result));
                }
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
                KillTimer(hwnd, TabBandWindow::kSessionFlushTimerId);
                self->ClearExplorerContext();
                self->ClearVisualItems();
                self->CloseThemeHandles();
                self->ClearDropHoverState();
                self->HidePreviewWindow(true);
                self->ReleaseBackBuffer();
                self->UnregisterShellNotifications();
                if (auto* manager = self->ResolveManager()) {
                    manager->UnregisterProgressListener(hwnd);
                }
                if (self->m_progressTimerActive) {
                    KillTimer(hwnd, TabBandWindow::kProgressTimerId);
                    self->m_progressTimerActive = false;
                }
                if (self->m_dropTargetRegistered) {
                    RevokeDragDrop(hwnd);
                    self->m_dropTargetRegistered = false;
                }
                self->m_dropTarget.Reset();
                self->m_dropTargetRegistrationPending = false;
                if (self->m_parentFrame) {
                    ClearAvailableDockMaskForFrame(self->m_parentFrame);
                    self->m_parentFrame = nullptr;
                }
                self->m_parentRebar = nullptr;
                self->m_rebarBandIndex = -1;
                self->InvalidateRebarIntegration();
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
