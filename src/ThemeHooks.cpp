#include "ThemeHooks.h"

#include <MinHook.h>

#include <CommCtrl.h>
#include <gdiplus.h>
#include <optional>
#include <mutex>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <algorithm>
#include <vector>
#include <cwchar>
#include <new>

#include <windowsx.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <wingdi.h>

#include "Logging.h"
#include "DpiUtils.h"
#include "CompositionIntercept.h"

// Note: Some theme part constants may not be available in older Windows SDKs
// They are conditionally included in the switch statement if defined

namespace shelltabs {

void UnregisterThemeSurface(HWND hwnd) noexcept;

namespace {

struct HwndHasher {
    size_t operator()(HWND hwnd) const noexcept { return reinterpret_cast<size_t>(hwnd); }
};

struct ThemeSurfaceRegistration {
    ExplorerGlowCoordinator* coordinator = nullptr;
    ExplorerSurfaceKind kind = ExplorerSurfaceKind::ListView;
};

// Forward declaration for PaintGlowSurface function
bool PaintGlowSurface(HDC dc, HWND hwnd, const RECT& rect, const GlowColorSet& colors, ExplorerSurfaceKind kind);

std::mutex g_registryMutex;
std::unordered_map<HWND, ThemeSurfaceRegistration, HwndHasher> g_surfaceRegistry;
std::unordered_set<HWND, HwndHasher> g_directUiHosts;

struct ScrollbarPaintMetrics {
    UINT dpiX = 96u;
    UINT dpiY = 96u;
    int baseLineThickness = 2;
    int baseHaloPadding = 3;
    int thumbAlongPadding = 6;
};

std::mutex g_scrollbarMetricsMutex;
std::unordered_map<HWND, ScrollbarPaintMetrics, HwndHasher> g_scrollbarMetrics;

std::mutex g_initializationMutex;
bool g_hooksActive = false;

using DrawThemeBackgroundFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, const RECT*, const RECT*);
using DrawThemeEdgeFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, const RECT*, UINT, UINT, RECT*);
using FillRectFn = int(WINAPI*)(HDC, const RECT*, HBRUSH);
using GdiGradientFillFn = BOOL(WINAPI*)(HDC, PTRIVERTEX, ULONG, PVOID, ULONG, ULONG);
using DirectUiDrawFn = HRESULT(STDMETHODCALLTYPE*)(void*, HDC, const RECT*, UINT, void*);
using TrackPopupMenuFn = BOOL(WINAPI*)(HMENU, UINT, int, int, int, HWND, const RECT*);
using TrackPopupMenuExFn = BOOL(WINAPI*)(HMENU, UINT, int, int, HWND, LPTPMPARAMS);
using CreateWindowExWFn = HWND(WINAPI*)(DWORD, LPCWSTR, LPCWSTR, DWORD, int, int, int, int, HWND, HMENU, HINSTANCE, LPVOID);

DrawThemeBackgroundFn g_originalDrawThemeBackground = nullptr;
DrawThemeEdgeFn g_originalDrawThemeEdge = nullptr;
FillRectFn g_originalFillRect = nullptr;
GdiGradientFillFn g_originalGdiGradientFill = nullptr;
TrackPopupMenuFn g_originalTrackPopupMenu = nullptr;
TrackPopupMenuExFn g_originalTrackPopupMenuEx = nullptr;
CreateWindowExWFn g_originalCreateWindowExW = nullptr;
void* g_drawThemeBackgroundTarget = nullptr;
void* g_drawThemeEdgeTarget = nullptr;
void* g_fillRectTarget = nullptr;
void* g_gdiGradientFillTarget = nullptr;
void* g_trackPopupMenuTarget = nullptr;
void* g_trackPopupMenuExTarget = nullptr;
void* g_createWindowExWTarget = nullptr;
std::vector<void*> g_directUiDetourTargets;

thread_local bool g_drawThemeBackgroundActive = false;
thread_local bool g_drawThemeEdgeActive = false;
thread_local bool g_fillRectActive = false;
thread_local bool g_gradientFillActive = false;
thread_local bool g_directUiDrawActive = false;

std::once_flag g_directUiThemeLogged;
std::once_flag g_directUiComLogged;

struct ThemePaintOverrideState {
    bool active = false;
    ExplorerSurfaceKind kind = ExplorerSurfaceKind::ListView;
    GlowColorSet colors{};
    HWND window = nullptr;
    bool suppressFallback = false;
};

thread_local ThemePaintOverrideState g_themePaintOverride{};

const ThemePaintOverrideState* GetThemePaintOverride() noexcept {
    return g_themePaintOverride.active ? &g_themePaintOverride : nullptr;
}

bool ActivateThemePaintOverride(HWND window, ExplorerSurfaceKind kind, const GlowColorSet& colors,
                                bool suppressFallback) noexcept {
    if (g_themePaintOverride.active) {
        return false;
    }
    g_themePaintOverride.active = true;
    g_themePaintOverride.kind = kind;
    g_themePaintOverride.colors = colors;
    g_themePaintOverride.window = window;
    g_themePaintOverride.suppressFallback = suppressFallback;
    return true;
}

void DeactivateThemePaintOverride() noexcept {
    g_themePaintOverride = {};
}

// Forward declarations
bool InstallHook(void* target, void* detour, void** original, const wchar_t* name);

struct PaintContextSnapshot {
    bool active = false;
    ThemeSurfaceRegistration registration{};
    HWND window = nullptr;
};

thread_local PaintContextSnapshot g_threadPaintContext{};

struct SurfaceLookupResult {
    ThemeSurfaceRegistration registration;
    HWND window = nullptr;
};

class ScopedPaintContext {
public:
    ScopedPaintContext() = default;

    explicit ScopedPaintContext(const SurfaceLookupResult& surface) {
        m_previous = g_threadPaintContext;
        g_threadPaintContext.active = true;
        g_threadPaintContext.registration = surface.registration;
        g_threadPaintContext.window = surface.window;
        m_active = true;
    }

    ~ScopedPaintContext() {
        if (m_active) {
            g_threadPaintContext = m_previous;
        }
    }

    ScopedPaintContext(const ScopedPaintContext&) = delete;
    ScopedPaintContext& operator=(const ScopedPaintContext&) = delete;

    ScopedPaintContext(ScopedPaintContext&& other) noexcept {
        m_previous = other.m_previous;
        m_active = other.m_active;
        other.m_active = false;
    }

    ScopedPaintContext& operator=(ScopedPaintContext&& other) noexcept {
        if (this != &other) {
            if (m_active) {
                g_threadPaintContext = m_previous;
            }
            m_previous = other.m_previous;
            m_active = other.m_active;
            other.m_active = false;
        }
        return *this;
    }

    bool Active() const noexcept { return m_active; }

private:
    PaintContextSnapshot m_previous{};
    bool m_active = false;
};

struct DirectUiElementInfo {
    ThemeSurfaceRegistration registration{};
    HWND host = nullptr;
    void* target = nullptr;
};

struct DirectUiTargetInfo {
    DirectUiDrawFn original = nullptr;
    bool installed = false;
};

std::mutex g_directUiDetourMutex;
std::unordered_map<void*, DirectUiElementInfo> g_directUiElementInfo;
std::unordered_map<void*, DirectUiTargetInfo> g_directUiTargetInfo;

struct PopupInvocationContext {
    ExplorerGlowCoordinator* coordinator = nullptr;
    DWORD threadId = 0;
    bool themingEnabled = false;
};

struct PopupSubclassData {
    ExplorerGlowCoordinator* coordinator = nullptr;
    ExplorerSurfaceKind kind = ExplorerSurfaceKind::PopupMenu;
};

thread_local std::vector<PopupInvocationContext> g_popupInvocationStack;

constexpr const wchar_t kPopupSubclassPropertyName[] = L"ShellTabs.PopupSubclass";
constexpr UINT_PTR kPopupSubclassId = 0x53545050;  // 'STPP'

class ReentrancyGuard {
public:
    explicit ReentrancyGuard(bool& flag) : m_flag(flag) {
        if (!m_flag) {
            m_flag = true;
            m_entered = true;
        }
    }

    ~ReentrancyGuard() {
        if (m_entered) {
            m_flag = false;
        }
    }

    bool Entered() const noexcept { return m_entered; }

private:
    bool& m_flag;
    bool m_entered = false;
};

std::optional<ThemeSurfaceRegistration> LookupRegistration(HWND hwnd) {
    std::lock_guard<std::mutex> guard(g_registryMutex);
    auto it = g_surfaceRegistry.find(hwnd);
    if (it == g_surfaceRegistry.end()) {
        return std::nullopt;
    }
    return it->second;
}

bool ShouldApplyPopupTheming(ExplorerGlowCoordinator* coordinator, ExplorerSurfaceKind kind) {
    return coordinator && coordinator->ShouldRenderSurface(kind);
}

std::optional<ThemeSurfaceRegistration> ResolveRegistrationForWindowHierarchy(HWND hwnd) {
    if (!hwnd) {
        return std::nullopt;
    }
    HWND current = hwnd;
    while (current) {
        auto registration = LookupRegistration(current);
        if (registration.has_value()) {
            return registration;
        }
        HWND parent = GetAncestor(current, GA_PARENT);
        if (!parent || parent == current) {
            break;
        }
        current = parent;
    }
    HWND owner = GetWindow(hwnd, GW_OWNER);
    if (owner && owner != hwnd) {
        return ResolveRegistrationForWindowHierarchy(owner);
    }
    return std::nullopt;
}

ExplorerGlowCoordinator* ResolveCoordinatorForWindow(HWND hwnd) {
    auto registration = ResolveRegistrationForWindowHierarchy(hwnd);
    if (registration.has_value()) {
        return registration->coordinator;
    }
    return nullptr;
}

PopupInvocationContext PreparePopupInvocation(HWND owner) {
    PopupInvocationContext context{};
    if (!owner) {
        return context;
    }
    auto registration = ResolveRegistrationForWindowHierarchy(owner);
    if (!registration.has_value()) {
        return context;
    }
    context.coordinator = registration->coordinator;
    context.threadId = GetCurrentThreadId();
    context.themingEnabled = ShouldApplyPopupTheming(context.coordinator, ExplorerSurfaceKind::PopupMenu);
    return context;
}

PopupInvocationContext* CurrentPopupInvocation() {
    if (g_popupInvocationStack.empty()) {
        return nullptr;
    }
    PopupInvocationContext& context = g_popupInvocationStack.back();
    if (context.threadId != GetCurrentThreadId()) {
        return nullptr;
    }
    return &context;
}

bool IsAtomResource(LPCWSTR value) {
    return reinterpret_cast<ULONG_PTR>(value) <= 0xFFFFu;
}

bool IsMenuClassName(LPCWSTR className) {
    if (!className) {
        return false;
    }
    if (IsAtomResource(className)) {
        return LOWORD(className) == 0x8000u;
    }
    return _wcsicmp(className, L"#32768") == 0;
}

bool IsTooltipClassName(LPCWSTR className) {
    if (!className || IsAtomResource(className)) {
        return false;
    }
    return _wcsicmp(className, TOOLTIPS_CLASSW) == 0 || _wcsicmp(className, L"tooltips_class32") == 0;
}

bool PaintPopupPaletteInternal(HWND hwnd, PopupSubclassData* data, HDC dc) {
    if (!hwnd || !data || !dc) {
        return false;
    }
    ExplorerGlowCoordinator* coordinator = data->coordinator;
    if (!ShouldApplyPopupTheming(coordinator, data->kind)) {
        return false;
    }
    RECT rect{};
    if (!GetClientRect(hwnd, &rect)) {
        return false;
    }
    GlowColorSet colors = coordinator->ResolveColors(data->kind);
    if (!colors.valid) {
        return false;
    }
    return PaintGlowSurface(dc, hwnd, rect, colors, data->kind);
}

LRESULT CALLBACK PopupSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                   UINT_PTR subclassId, DWORD_PTR refData);

bool InstallPopupSubclass(HWND hwnd, ExplorerGlowCoordinator* coordinator, ExplorerSurfaceKind kind) {
    if (!hwnd || !coordinator) {
        return false;
    }
    if (!ShouldApplyPopupTheming(coordinator, kind)) {
        return false;
    }

    auto* existing = reinterpret_cast<PopupSubclassData*>(GetPropW(hwnd, kPopupSubclassPropertyName));
    if (existing) {
        existing->coordinator = coordinator;
        existing->kind = kind;
        RegisterThemeSurface(hwnd, kind, coordinator);
        InvalidateRect(hwnd, nullptr, TRUE);
        return true;
    }

    auto* data = new (std::nothrow) PopupSubclassData();
    if (!data) {
        return false;
    }
    data->coordinator = coordinator;
    data->kind = kind;

    if (!SetWindowSubclass(hwnd, PopupSubclassProc, kPopupSubclassId, reinterpret_cast<DWORD_PTR>(data))) {
        delete data;
        return false;
    }
    if (!SetPropW(hwnd, kPopupSubclassPropertyName, reinterpret_cast<HANDLE>(data))) {
        RemoveWindowSubclass(hwnd, PopupSubclassProc, kPopupSubclassId);
        delete data;
        return false;
    }

    RegisterThemeSurface(hwnd, kind, coordinator);
    InvalidateRect(hwnd, nullptr, TRUE);
    return true;
}

LRESULT CALLBACK PopupSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                   UINT_PTR subclassId, DWORD_PTR refData) {
    auto* data = reinterpret_cast<PopupSubclassData*>(refData);
    switch (msg) {
        case WM_ERASEBKGND: {
            if (PaintPopupPaletteInternal(hwnd, data, reinterpret_cast<HDC>(wParam))) {
                return 1;
            }
            break;
        }
        case WM_PRINTCLIENT: {
            if (PaintPopupPaletteInternal(hwnd, data, reinterpret_cast<HDC>(wParam))) {
                return 0;
            }
            break;
        }
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED: {
            InvalidateRect(hwnd, nullptr, TRUE);
            break;
        }
        case WM_NCDESTROY: {
            if (data) {
                UnregisterThemeSurface(hwnd);
                RemovePropW(hwnd, kPopupSubclassPropertyName);
                delete data;
            }
            RemoveWindowSubclass(hwnd, PopupSubclassProc, subclassId);
            break;
        }
        default:
            break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

std::optional<SurfaceLookupResult> ResolveSurfaceFromWindowDc(HDC dc) {
    if (!dc) {
        return std::nullopt;
    }
    HWND window = WindowFromDC(dc);
    HWND current = window;
    while (current) {
        auto registration = LookupRegistration(current);
        if (registration.has_value()) {
            if (!IsWindow(current)) {
                UnregisterThemeSurface(current);
                current = GetAncestor(current, GA_PARENT);
                continue;
            }
            ExplorerGlowCoordinator* coordinator = registration->coordinator;
            if (coordinator && coordinator->ShouldRenderSurface(registration->kind)) {
                SurfaceLookupResult result{};
                result.registration = *registration;
                result.window = current;
                return result;
            }
        }
        HWND parent = GetAncestor(current, GA_PARENT);
        if (!parent || parent == current) {
            break;
        }
        current = parent;
    }
    return std::nullopt;
}

std::optional<SurfaceLookupResult> ResolveSurfaceForPainting(HDC dc) {
    auto surface = ResolveSurfaceFromWindowDc(dc);
    if (surface.has_value()) {
        return surface;
    }

    if (g_threadPaintContext.active) {
        SurfaceLookupResult result{};
        result.registration = g_threadPaintContext.registration;
        result.window = g_threadPaintContext.window;
        return result;
    }
    return std::nullopt;
}

std::optional<ExplorerSurfaceKind> ResolveSurfaceKindFromThemePart(int partId) noexcept {
    // Note: Theme part constants can have overlapping values between different theme classes
    // We use separate checks to avoid switch case conflicts

    // Check Rebar theme parts
    if (partId == RP_GRIPPER || partId == RP_GRIPPERVERT || partId == RP_BAND ||
        partId == RP_BACKGROUND || partId == RP_CHEVRON || partId == RP_CHEVRONVERT ||
        partId == RP_SPLITTER || partId == RP_SPLITTERVERT) {
        return ExplorerSurfaceKind::Rebar;
    }

    // Check Toolbar theme parts
    if (partId == TP_BUTTON || partId == TP_DROPDOWNBUTTON || partId == TP_SPLITBUTTON ||
        partId == TP_SPLITBUTTONDROPDOWN || partId == TP_SEPARATOR || partId == TP_SEPARATORVERT) {
        return ExplorerSurfaceKind::Toolbar;
    }

    return std::nullopt;
}

int ScaleByDpiInternal(int value, UINT dpi) {
    if (dpi == 0) {
        dpi = 96u;
    }
    return std::max(1, MulDiv(value, static_cast<int>(dpi), 96));
}

ScrollbarPaintMetrics ComputeScrollbarMetrics(HWND hwnd) {
    ScrollbarPaintMetrics metrics{};
    const UINT dpi = hwnd ? GetWindowDpi(hwnd) : 96u;
    metrics.dpiX = dpi;
    metrics.dpiY = dpi;
    metrics.baseLineThickness = ScaleByDpiInternal(2, metrics.dpiX);
    metrics.baseHaloPadding = ScaleByDpiInternal(3, metrics.dpiX);
    metrics.thumbAlongPadding = ScaleByDpiInternal(6, metrics.dpiY);
    return metrics;
}

ScrollbarPaintMetrics ResolveScrollbarMetrics(HWND hwnd) {
    if (!hwnd) {
        return ComputeScrollbarMetrics(nullptr);
    }

    std::lock_guard<std::mutex> guard(g_scrollbarMetricsMutex);
    auto it = g_scrollbarMetrics.find(hwnd);
    if (it != g_scrollbarMetrics.end()) {
        return it->second;
    }

    ScrollbarPaintMetrics metrics = ComputeScrollbarMetrics(hwnd);
    g_scrollbarMetrics.emplace(hwnd, metrics);
    return metrics;
}

void InvalidateScrollbarMetricsInternal(HWND hwnd) noexcept {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_scrollbarMetricsMutex);
    g_scrollbarMetrics.erase(hwnd);
}

void FillGradientRect(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& rect, BYTE alpha,
                      float angle = 90.0f) {
    if (!colors.valid) {
        return;
    }
    const int width = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    if (width <= 0 || height <= 0) {
        return;
    }

    const Gdiplus::Rect paintRect(rect.left, rect.top, width, height);
    const Gdiplus::Color start(alpha, GetRValue(colors.start), GetGValue(colors.start), GetBValue(colors.start));
    const Gdiplus::Color end(alpha, GetRValue(colors.end), GetGValue(colors.end), GetBValue(colors.end));

    if (!colors.gradient || colors.start == colors.end) {
        Gdiplus::SolidBrush brush(start);
        graphics.FillRectangle(&brush, paintRect);
        return;
    }

    Gdiplus::LinearGradientBrush brush(paintRect, start, end, angle);
    graphics.FillRectangle(&brush, paintRect);
}

enum class ScrollbarOrientation {
    Horizontal,
    Vertical,
};

enum class ScrollbarPartKind {
    None,
    Track,
    Thumb,
};

struct ScrollbarPartInfo {
    ScrollbarPartKind kind = ScrollbarPartKind::None;
    ScrollbarOrientation orientation = ScrollbarOrientation::Vertical;
};

std::optional<ScrollbarPartInfo> DescribeScrollbarPart(int partId) {
    ScrollbarPartInfo info{};
    switch (partId) {
        case SBP_THUMBBTNHORZ:
            info.kind = ScrollbarPartKind::Thumb;
            info.orientation = ScrollbarOrientation::Horizontal;
            return info;
        case SBP_THUMBBTNVERT:
            info.kind = ScrollbarPartKind::Thumb;
            info.orientation = ScrollbarOrientation::Vertical;
            return info;
        case SBP_LOWERTRACKHORZ:
        case SBP_UPPERTRACKHORZ:
            info.kind = ScrollbarPartKind::Track;
            info.orientation = ScrollbarOrientation::Horizontal;
            return info;
        case SBP_LOWERTRACKVERT:
        case SBP_UPPERTRACKVERT:
            info.kind = ScrollbarPartKind::Track;
            info.orientation = ScrollbarOrientation::Vertical;
            return info;
        default:
            break;
    }
    return std::nullopt;
}

bool PaintScrollbarPart(HDC dc, const RECT& partRect, const RECT& clipRect, HWND hwnd,
                        const ScrollbarPartInfo& info, const ScrollbarGlowDefinition& definition) {
    if (!definition.colors.valid) {
        return false;
    }
    RECT intersection{};
    if (!IntersectRect(&intersection, &partRect, &clipRect) || IsRectEmpty(&intersection)) {
        return true;
    }

    const ScrollbarPaintMetrics metrics = ResolveScrollbarMetrics(hwnd);
    const int crossExtent =
        (info.orientation == ScrollbarOrientation::Vertical) ? (partRect.right - partRect.left)
                                                             : (partRect.bottom - partRect.top);
    if (crossExtent <= 0) {
        return false;
    }

    const int clampedCross = std::max(crossExtent, 1);
    const int lineThickness = std::clamp(metrics.baseLineThickness, 1, clampedCross);
    const int haloPadding = std::max(lineThickness, metrics.baseHaloPadding);
    const int thumbAlongPadding = std::max(metrics.thumbAlongPadding, 1);

    Gdiplus::Graphics graphics(dc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    if (info.kind == ScrollbarPartKind::Track) {
        RECT haloRect = partRect;
        if (info.orientation == ScrollbarOrientation::Vertical) {
            haloRect.left -= haloPadding;
            haloRect.right += haloPadding;
        } else {
            haloRect.top -= haloPadding;
            haloRect.bottom += haloPadding;
        }

        RECT haloClip{};
        if (IntersectRect(&haloClip, &haloRect, &clipRect) && !IsRectEmpty(&haloClip)) {
            FillGradientRect(graphics, definition.colors, haloClip, definition.trackHaloAlpha);
        }

        RECT lineRect = partRect;
        if (info.orientation == ScrollbarOrientation::Vertical) {
            const int center = partRect.left + ((partRect.right - partRect.left) - lineThickness) / 2;
            lineRect.left = center;
            lineRect.right = center + lineThickness;
        } else {
            const int center = partRect.top + ((partRect.bottom - partRect.top) - lineThickness) / 2;
            lineRect.top = center;
            lineRect.bottom = center + lineThickness;
        }

        RECT lineClip{};
        if (IntersectRect(&lineClip, &lineRect, &clipRect) && !IsRectEmpty(&lineClip)) {
            FillGradientRect(graphics, definition.colors, lineClip, definition.trackLineAlpha);
        }
        return true;
    }

    if (info.kind == ScrollbarPartKind::Thumb) {
        RECT haloRect = partRect;
        if (info.orientation == ScrollbarOrientation::Vertical) {
            haloRect.left -= haloPadding;
            haloRect.right += haloPadding;
            haloRect.top -= thumbAlongPadding;
            haloRect.bottom += thumbAlongPadding;
        } else {
            haloRect.left -= thumbAlongPadding;
            haloRect.right += thumbAlongPadding;
            haloRect.top -= haloPadding;
            haloRect.bottom += haloPadding;
        }

        RECT haloClip{};
        if (IntersectRect(&haloClip, &haloRect, &clipRect) && !IsRectEmpty(&haloClip)) {
            FillGradientRect(graphics, definition.colors, haloClip, definition.thumbHaloAlpha);
        }

        RECT thumbClip{};
        if (IntersectRect(&thumbClip, &partRect, &clipRect) && !IsRectEmpty(&thumbClip)) {
            FillGradientRect(graphics, definition.colors, thumbClip, definition.thumbFillAlpha);
        }
        return true;
    }

    return false;
}

COLOR16 ToColor16(BYTE component) {
    return static_cast<COLOR16>(static_cast<unsigned int>(component) << 8);
}

enum class GradientOrientation {
    Vertical,
    Horizontal,
};

GradientOrientation DetermineOrientation(ExplorerSurfaceKind kind, HWND hwnd) {
    switch (kind) {
        case ExplorerSurfaceKind::Toolbar:
        case ExplorerSurfaceKind::Rebar: {
            RECT rect{};
            if (hwnd && GetClientRect(hwnd, &rect)) {
                if ((rect.bottom - rect.top) > (rect.right - rect.left)) {
                    return GradientOrientation::Horizontal;
                }
            }
            const LONG_PTR style = hwnd ? GetWindowLongPtr(hwnd, GWL_STYLE) : 0;
            if ((style & CCS_VERT) == CCS_VERT) {
                return GradientOrientation::Horizontal;
            }
            return GradientOrientation::Vertical;
        }
        case ExplorerSurfaceKind::Scrollbar: {
            const LONG_PTR style = hwnd ? GetWindowLongPtr(hwnd, GWL_STYLE) : 0;
            if ((style & SBS_VERT) == SBS_VERT) {
                return GradientOrientation::Vertical;
            }
            return GradientOrientation::Horizontal;
        }
        default:
            return GradientOrientation::Vertical;
    }
}

bool PaintSolid(HDC dc, const RECT& rect, COLORREF color) {
    if (!g_originalFillRect) {
        return false;
    }
    HBRUSH brush = CreateSolidBrush(color);
    if (!brush) {
        return false;
    }
    const int result = g_originalFillRect(dc, &rect, brush);
    DeleteObject(brush);
    return result != 0;
}

bool PaintGradient(HDC dc, const RECT& rect, const GlowColorSet& colors, GradientOrientation orientation) {
    if (!g_originalGdiGradientFill) {
        return false;
    }
    TRIVERTEX vertices[2] = {};
    vertices[0].x = rect.left;
    vertices[0].y = rect.top;
    vertices[0].Red = ToColor16(GetRValue(colors.start));
    vertices[0].Green = ToColor16(GetGValue(colors.start));
    vertices[0].Blue = ToColor16(GetBValue(colors.start));
    vertices[0].Alpha = 0;

    vertices[1].x = rect.right;
    vertices[1].y = rect.bottom;
    vertices[1].Red = ToColor16(GetRValue(colors.end));
    vertices[1].Green = ToColor16(GetGValue(colors.end));
    vertices[1].Blue = ToColor16(GetBValue(colors.end));
    vertices[1].Alpha = 0;

    GRADIENT_RECT gradientRect = {0, 1};
    ULONG mode = (orientation == GradientOrientation::Horizontal) ? GRADIENT_FILL_RECT_H : GRADIENT_FILL_RECT_V;
    return g_originalGdiGradientFill(dc, vertices, 2, &gradientRect, 1, mode) != FALSE;
}

bool PaintGlowSurface(HDC dc, HWND hwnd, const RECT& rect, const GlowColorSet& colors,
                      ExplorerSurfaceKind kind) {
    if (rect.right <= rect.left || rect.bottom <= rect.top) {
        return false;
    }
    if (!colors.valid) {
        return false;
    }
    if (colors.gradient && colors.start != colors.end) {
        if (PaintGradient(dc, rect, colors, DetermineOrientation(kind, hwnd))) {
            return true;
        }
    }
    return PaintSolid(dc, rect, colors.start);
}

HRESULT STDMETHODCALLTYPE DirectUiDrawDetour(void* element, HDC dc, const RECT* rect, UINT state, void* data) {
    DirectUiElementInfo elementInfo{};
    DirectUiDrawFn original = nullptr;

    {
        std::lock_guard<std::mutex> guard(g_directUiDetourMutex);
        auto elementIt = g_directUiElementInfo.find(element);
        if (elementIt != g_directUiElementInfo.end()) {
            elementInfo = elementIt->second;
            auto targetIt = g_directUiTargetInfo.find(elementInfo.target);
            if (targetIt != g_directUiTargetInfo.end()) {
                original = targetIt->second.original;
            }
        }
    }

    if (!original) {
        return S_OK;
    }

    ReentrancyGuard guard(g_directUiDrawActive);
    if (!guard.Entered()) {
        return original(element, dc, rect, state, data);
    }

    bool hostRegistered = false;
    {
        std::lock_guard<std::mutex> registryGuard(g_registryMutex);
        hostRegistered = g_directUiHosts.find(elementInfo.host) != g_directUiHosts.end();
    }
    if (!hostRegistered) {
        return original(element, dc, rect, state, data);
    }

    SurfaceLookupResult surface{};
    surface.registration = elementInfo.registration;
    surface.window = elementInfo.host;

    ScopedPaintContext context(surface);
    std::call_once(g_directUiComLogged,
                   []() { LogMessage(LogLevel::Info, L"ThemeHooks: DirectUI COM hook active"); });

    if (rect && elementInfo.registration.coordinator) {
        GlowColorSet colors = elementInfo.registration.coordinator->ResolveColors(elementInfo.registration.kind);
        if (colors.valid) {
            PaintGlowSurface(dc, elementInfo.host, *rect, colors, elementInfo.registration.kind);
        }
    }

    return original(element, dc, rect, state, data);
}

BOOL WINAPI TrackPopupMenuDetour(HMENU menu, UINT flags, int x, int y, int reserved, HWND hwnd, const RECT* rect) {
    PopupInvocationContext context = PreparePopupInvocation(hwnd);
    if (context.themingEnabled) {
        g_popupInvocationStack.push_back(context);
    }

    BOOL result = g_originalTrackPopupMenu ?
                      g_originalTrackPopupMenu(menu, flags, x, y, reserved, hwnd, rect)
                      : FALSE;

    if (context.themingEnabled) {
        g_popupInvocationStack.pop_back();
    }
    return result;
}

BOOL WINAPI TrackPopupMenuExDetour(HMENU menu, UINT flags, int x, int y, HWND hwnd, LPTPMPARAMS params) {
    PopupInvocationContext context = PreparePopupInvocation(hwnd);
    if (context.themingEnabled) {
        g_popupInvocationStack.push_back(context);
    }

    BOOL result = g_originalTrackPopupMenuEx ?
                      g_originalTrackPopupMenuEx(menu, flags, x, y, hwnd, params)
                      : FALSE;

    if (context.themingEnabled) {
        g_popupInvocationStack.pop_back();
    }
    return result;
}

HWND WINAPI CreateWindowExWDetour(DWORD exStyle, LPCWSTR className, LPCWSTR windowName, DWORD style, int x, int y,
                                  int width, int height, HWND hwndParent, HMENU menu, HINSTANCE instance,
                                  LPVOID param) {
    const bool isMenuClass = IsMenuClassName(className);
    const bool isTooltipClass = !isMenuClass && IsTooltipClassName(className);

    ExplorerGlowCoordinator* coordinator = nullptr;
    ExplorerSurfaceKind kind = ExplorerSurfaceKind::PopupMenu;

    if (isMenuClass) {
        if (auto* context = CurrentPopupInvocation()) {
            if (context->themingEnabled) {
                coordinator = context->coordinator;
            }
        }
        kind = ExplorerSurfaceKind::PopupMenu;
    } else if (isTooltipClass) {
        coordinator = ResolveCoordinatorForWindow(hwndParent);
        kind = ExplorerSurfaceKind::Tooltip;
    }

    HWND hwnd = g_originalCreateWindowExW ? g_originalCreateWindowExW(exStyle, className, windowName, style, x, y,
                                                                       width, height, hwndParent, menu, instance, param)
                                          : nullptr;

    if (hwnd && coordinator) {
        InstallPopupSubclass(hwnd, coordinator, kind);
    }
    return hwnd;
}

HRESULT WINAPI DrawThemeBackgroundDetour(HTHEME theme, HDC dc, int partId, int stateId, const RECT* rect,
                                         const RECT* clipRect) {
    ReentrancyGuard guard(g_drawThemeBackgroundActive);
    if (!guard.Entered()) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }
    if (!rect) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    auto surface = ResolveSurfaceFromWindowDc(dc);
    const ThemePaintOverrideState* override = GetThemePaintOverride();

    ExplorerGlowCoordinator* coordinator = surface.has_value() ? surface->registration.coordinator : nullptr;
    ExplorerSurfaceKind surfaceKind = ExplorerSurfaceKind::ListView;
    bool haveSurfaceKind = false;

    if (override) {
        surfaceKind = override->kind;
        haveSurfaceKind = true;
    } else if (surface.has_value()) {
        surfaceKind = surface->registration.kind;
        haveSurfaceKind = true;
    } else if (auto mapped = ResolveSurfaceKindFromThemePart(partId); mapped.has_value()) {
        surfaceKind = *mapped;
        haveSurfaceKind = true;
    }

    if (!haveSurfaceKind) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    if (surfaceKind == ExplorerSurfaceKind::Scrollbar && !override) {
        if (!coordinator) {
            return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
        }
        auto definition = coordinator->ResolveScrollbarDefinition();
        if (!definition.has_value()) {
            return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
        }

        auto partInfo = DescribeScrollbarPart(partId);
        if (!partInfo.has_value()) {
            return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
        }

        RECT clipBounds = *rect;
        if (clipRect) {
            RECT clipped{};
            if (!IntersectRect(&clipped, &clipBounds, clipRect) || IsRectEmpty(&clipped)) {
                return S_OK;
            }
            clipBounds = clipped;
        }

        if (PaintScrollbarPart(dc, *rect, clipBounds, surface->window, *partInfo, *definition)) {
            return S_OK;
        }

        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    GlowColorSet colors{};
    if (override && override->colors.valid) {
        colors = override->colors;
    } else if (coordinator) {
        colors = coordinator->ResolveColors(surfaceKind);
    }

    if (!colors.valid) {
        if (override && override->suppressFallback) {
            return S_OK;
        }
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    HWND paintWindow = nullptr;
    if (override && override->window) {
        paintWindow = override->window;
    } else if (surface.has_value()) {
        paintWindow = surface->window;
    }

    RECT paintRect = *rect;
    if (clipRect) {
        RECT clipped{};
        if (!IntersectRect(&clipped, &paintRect, clipRect) || clipped.right <= clipped.left ||
            clipped.bottom <= clipped.top) {
            if (surfaceKind == ExplorerSurfaceKind::DirectUi && surface.has_value()) {
                ScopedPaintContext context(*surface);
                return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
            }
            return S_OK;
        }
        paintRect = clipped;
    }

    if (surfaceKind == ExplorerSurfaceKind::DirectUi) {
        if (!surface.has_value()) {
            if (override && override->suppressFallback) {
                return S_OK;
            }
            return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
        }
        ScopedPaintContext context(*surface);
        std::call_once(g_directUiThemeLogged,
                       []() { LogMessage(LogLevel::Info, L"ThemeHooks: DirectUI theme hook active"); });
        PaintGlowSurface(dc, paintWindow, paintRect, colors, surfaceKind);
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    if (!PaintGlowSurface(dc, paintWindow, paintRect, colors, surfaceKind)) {
        if (override && override->suppressFallback) {
            return S_OK;
        }
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    return S_OK;
}

HRESULT WINAPI DrawThemeEdgeDetour(HTHEME theme, HDC dc, int partId, int stateId, const RECT* rect, UINT edge,
                                   UINT flags, RECT* contentRect) {
    ReentrancyGuard guard(g_drawThemeEdgeActive);
    if (!guard.Entered()) {
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }
    if (!rect) {
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }

    auto surface = ResolveSurfaceFromWindowDc(dc);
    const ThemePaintOverrideState* override = GetThemePaintOverride();

    ExplorerGlowCoordinator* coordinator = surface.has_value() ? surface->registration.coordinator : nullptr;
    ExplorerSurfaceKind surfaceKind = ExplorerSurfaceKind::ListView;
    bool haveSurfaceKind = false;

    if (override) {
        surfaceKind = override->kind;
        haveSurfaceKind = true;
    } else if (surface.has_value()) {
        surfaceKind = surface->registration.kind;
        haveSurfaceKind = true;
    } else if (auto mapped = ResolveSurfaceKindFromThemePart(partId); mapped.has_value()) {
        surfaceKind = *mapped;
        haveSurfaceKind = true;
    }

    if (!haveSurfaceKind) {
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }

    if (surfaceKind == ExplorerSurfaceKind::Scrollbar && !override) {
        auto partInfo = DescribeScrollbarPart(partId);
        if (partInfo.has_value()) {
            if (contentRect) {
                *contentRect = *rect;
            }
            return S_OK;
        }
    }

    GlowColorSet colors{};
    if (override && override->colors.valid) {
        colors = override->colors;
    } else if (coordinator) {
        colors = coordinator->ResolveColors(surfaceKind);
    }

    if (!colors.valid) {
        if (override && override->suppressFallback) {
            if (contentRect) {
                *contentRect = *rect;
            }
            return S_OK;
        }
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }

    HWND paintWindow = nullptr;
    if (override && override->window) {
        paintWindow = override->window;
    } else if (surface.has_value()) {
        paintWindow = surface->window;
    }

    if (surfaceKind == ExplorerSurfaceKind::DirectUi) {
        if (!surface.has_value()) {
            if (override && override->suppressFallback) {
                if (contentRect) {
                    *contentRect = *rect;
                }
                return S_OK;
            }
            return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
        }
        ScopedPaintContext context(*surface);
        std::call_once(g_directUiThemeLogged,
                       []() { LogMessage(LogLevel::Info, L"ThemeHooks: DirectUI theme hook active"); });
        PaintGlowSurface(dc, paintWindow, *rect, colors, surfaceKind);
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }

    if (!PaintGlowSurface(dc, paintWindow, *rect, colors, surfaceKind)) {
        if (override && override->suppressFallback) {
            if (contentRect) {
                *contentRect = *rect;
            }
            return S_OK;
        }
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }
    if (contentRect) {
        *contentRect = *rect;
    }
    return S_OK;
}

int WINAPI FillRectDetour(HDC dc, const RECT* rect, HBRUSH brush) {
    ReentrancyGuard guard(g_fillRectActive);
    if (!guard.Entered()) {
        return g_originalFillRect(dc, rect, brush);
    }

    auto surface = ResolveSurfaceForPainting(dc);
    if (!surface.has_value()) {
        return g_originalFillRect(dc, rect, brush);
    }

    RECT paintRect{};
    if (rect) {
        paintRect = *rect;
    } else if (!surface->window || !GetClientRect(surface->window, &paintRect)) {
        return g_originalFillRect(dc, rect, brush);
    }

    GlowColorSet colors = surface->registration.coordinator->ResolveColors(surface->registration.kind);
    if (!colors.valid) {
        return g_originalFillRect(dc, rect, brush);
    }

    if (!PaintGlowSurface(dc, surface->window, paintRect, colors, surface->registration.kind)) {
        return g_originalFillRect(dc, rect, brush);
    }
    return 1;
}

RECT ComputeBoundingRect(const TRIVERTEX* vertices, ULONG count) {
    RECT rect{0, 0, 0, 0};
    if (!vertices || count == 0) {
        return rect;
    }
    LONG left = vertices[0].x;
    LONG right = vertices[0].x;
    LONG top = vertices[0].y;
    LONG bottom = vertices[0].y;
    for (ULONG index = 1; index < count; ++index) {
        left = std::min<LONG>(left, vertices[index].x);
        right = std::max<LONG>(right, vertices[index].x);
        top = std::min<LONG>(top, vertices[index].y);
        bottom = std::max<LONG>(bottom, vertices[index].y);
    }
    rect.left = left;
    rect.right = right;
    rect.top = top;
    rect.bottom = bottom;
    return rect;
}

BOOL WINAPI GdiGradientFillDetour(HDC dc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh, ULONG meshCount,
                                  ULONG mode) {
    ReentrancyGuard guard(g_gradientFillActive);
    if (!guard.Entered()) {
        return g_originalGdiGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode);
    }
    if (!vertices || vertexCount < 2) {
        return g_originalGdiGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode);
    }

    auto surface = ResolveSurfaceForPainting(dc);
    if (!surface.has_value()) {
        return g_originalGdiGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode);
    }

    GlowColorSet colors = surface->registration.coordinator->ResolveColors(surface->registration.kind);
    if (!colors.valid) {
        return g_originalGdiGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode);
    }

    RECT paintRect = ComputeBoundingRect(vertices, vertexCount);
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        return g_originalGdiGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode);
    }

    GradientOrientation orientation = (mode == GRADIENT_FILL_RECT_H) ? GradientOrientation::Horizontal
                                                                     : GradientOrientation::Vertical;
    if (!PaintGradient(dc, paintRect, colors, orientation)) {
        return g_originalGdiGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode);
    }
    return TRUE;
}

void RegisterDirectUiHostInternal(HWND hwnd) noexcept {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_registryMutex);
    g_directUiHosts.insert(hwnd);
}

void UnregisterDirectUiHostInternal(HWND hwnd) noexcept {
    if (!hwnd) {
        return;
    }
    {
        std::lock_guard<std::mutex> guard(g_registryMutex);
        g_directUiHosts.erase(hwnd);
    }

    std::lock_guard<std::mutex> detourGuard(g_directUiDetourMutex);
    for (auto it = g_directUiElementInfo.begin(); it != g_directUiElementInfo.end();) {
        if (it->second.host == hwnd) {
            it = g_directUiElementInfo.erase(it);
        } else {
            ++it;
        }
    }
}

void RegisterDirectUiRenderInterfaceInternal(void* element, size_t drawIndex, HWND host,
                                             ExplorerGlowCoordinator* coordinator) noexcept {
    if (!element || !coordinator || !host) {
        return;
    }
    if (!g_hooksActive) {
        return;
    }
    auto** vtablePtr = reinterpret_cast<void***>(element);
    if (!vtablePtr || !*vtablePtr) {
        return;
    }

    constexpr size_t kMaxDirectUiVtableSlots = 128;
    if (drawIndex >= kMaxDirectUiVtableSlots) {
        return;
    }

    void** vtable = *vtablePtr;
    void* target = vtable[drawIndex];
    if (!target) {
        return;
    }

    std::lock_guard<std::mutex> guard(g_directUiDetourMutex);
    if (g_directUiElementInfo.find(element) != g_directUiElementInfo.end()) {
        auto& info = g_directUiElementInfo[element];
        info.registration.coordinator = coordinator;
        info.host = host;
        info.target = target;
        return;
    }

    auto targetIt = g_directUiTargetInfo.find(target);
    if (targetIt == g_directUiTargetInfo.end() || !targetIt->second.installed) {
        DirectUiTargetInfo info{};
        if (!InstallHook(target, reinterpret_cast<void*>(&DirectUiDrawDetour),
                         reinterpret_cast<void**>(&info.original), L"DirectUiDraw")) {
            return;
        }
        info.installed = true;
        g_directUiTargetInfo[target] = info;
        g_directUiDetourTargets.push_back(target);
    }

    DirectUiElementInfo elementInfo{};
    elementInfo.registration = ThemeSurfaceRegistration{coordinator, ExplorerSurfaceKind::DirectUi};
    elementInfo.host = host;
    elementInfo.target = target;
    g_directUiElementInfo[element] = elementInfo;
}

bool InstallHook(void* target, void* detour, void** original, const wchar_t* name) {
    if (!target) {
        LogMessage(LogLevel::Warning, L"ThemeHooks: missing target for %ls", name);
        return false;
    }
    MH_STATUS status = MH_CreateHook(target, detour, original);
    if (status != MH_OK) {
        LogMessage(LogLevel::Error, L"ThemeHooks: MH_CreateHook failed for %ls (status=%d)", name, status);
        return false;
    }
    status = MH_EnableHook(target);
    if (status != MH_OK) {
        LogMessage(LogLevel::Error, L"ThemeHooks: MH_EnableHook failed for %ls (status=%d)", name, status);
        MH_DisableHook(target);
        return false;
    }
    return true;
}

void DisableHook(void* target, const wchar_t* name) {
    if (!target) {
        return;
    }
    MH_STATUS status = MH_DisableHook(target);
    if (status != MH_OK) {
        LogMessage(LogLevel::Warning, L"ThemeHooks: MH_DisableHook failed for %ls (status=%d)", name, status);
    }
}

}  // namespace

void RegisterDirectUiHost(HWND hwnd) noexcept {
    RegisterDirectUiHostInternal(hwnd);
}

void UnregisterDirectUiHost(HWND hwnd) noexcept {
    UnregisterDirectUiHostInternal(hwnd);
}

void RegisterDirectUiRenderInterface(void* element, size_t drawIndex, HWND host,
                                     ExplorerGlowCoordinator* coordinator) noexcept {
    RegisterDirectUiRenderInterfaceInternal(element, drawIndex, host, coordinator);
}

void InvalidateScrollbarMetrics(HWND hwnd) noexcept { InvalidateScrollbarMetricsInternal(hwnd); }

bool InitializeThemeHooks() {
    std::lock_guard<std::mutex> guard(g_initializationMutex);
    if (g_hooksActive) {
        return true;
    }

    MH_STATUS status = MH_Initialize();
    if (status != MH_OK && status != MH_ERROR_ALREADY_INITIALIZED) {
        LogMessage(LogLevel::Error, L"ThemeHooks: MH_Initialize failed (status=%d)", status);
        return false;
    }
    const bool initializedHere = (status == MH_OK);

    HMODULE uxtheme = GetModuleHandleW(L"uxtheme.dll");
    if (!uxtheme) {
        uxtheme = LoadLibraryW(L"uxtheme.dll");
    }
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    if (!user32) {
        user32 = LoadLibraryW(L"user32.dll");
    }
    HMODULE gdi32 = GetModuleHandleW(L"gdi32.dll");
    if (!gdi32) {
        gdi32 = LoadLibraryW(L"gdi32.dll");
    }

    g_drawThemeBackgroundTarget = uxtheme ? reinterpret_cast<void*>(GetProcAddress(uxtheme, "DrawThemeBackground")) : nullptr;
    g_drawThemeEdgeTarget = uxtheme ? reinterpret_cast<void*>(GetProcAddress(uxtheme, "DrawThemeEdge")) : nullptr;
    g_fillRectTarget = user32 ? reinterpret_cast<void*>(GetProcAddress(user32, "FillRect")) : nullptr;
    g_gdiGradientFillTarget = gdi32 ? reinterpret_cast<void*>(GetProcAddress(gdi32, "GdiGradientFill")) : nullptr;
    g_trackPopupMenuTarget = user32 ? reinterpret_cast<void*>(GetProcAddress(user32, "TrackPopupMenu")) : nullptr;
    g_trackPopupMenuExTarget = user32 ? reinterpret_cast<void*>(GetProcAddress(user32, "TrackPopupMenuEx")) : nullptr;
    g_createWindowExWTarget = user32 ? reinterpret_cast<void*>(GetProcAddress(user32, "CreateWindowExW")) : nullptr;

    if (!InstallHook(g_drawThemeBackgroundTarget, reinterpret_cast<void*>(&DrawThemeBackgroundDetour),
                     reinterpret_cast<void**>(&g_originalDrawThemeBackground), L"DrawThemeBackground")) {
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    if (!InstallHook(g_drawThemeEdgeTarget, reinterpret_cast<void*>(&DrawThemeEdgeDetour),
                     reinterpret_cast<void**>(&g_originalDrawThemeEdge), L"DrawThemeEdge")) {
        DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    if (!InstallHook(g_fillRectTarget, reinterpret_cast<void*>(&FillRectDetour),
                     reinterpret_cast<void**>(&g_originalFillRect), L"FillRect")) {
        DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");
        DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    if (!InstallHook(g_gdiGradientFillTarget, reinterpret_cast<void*>(&GdiGradientFillDetour),
                     reinterpret_cast<void**>(&g_originalGdiGradientFill), L"GdiGradientFill")) {
        DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");
        DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
        DisableHook(g_fillRectTarget, L"FillRect");
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    if (!InstallHook(g_trackPopupMenuTarget, reinterpret_cast<void*>(&TrackPopupMenuDetour),
                     reinterpret_cast<void**>(&g_originalTrackPopupMenu), L"TrackPopupMenu")) {
        DisableHook(g_gdiGradientFillTarget, L"GdiGradientFill");
        DisableHook(g_fillRectTarget, L"FillRect");
        DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
        DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    if (!InstallHook(g_trackPopupMenuExTarget, reinterpret_cast<void*>(&TrackPopupMenuExDetour),
                     reinterpret_cast<void**>(&g_originalTrackPopupMenuEx), L"TrackPopupMenuEx")) {
        DisableHook(g_trackPopupMenuTarget, L"TrackPopupMenu");
        DisableHook(g_gdiGradientFillTarget, L"GdiGradientFill");
        DisableHook(g_fillRectTarget, L"FillRect");
        DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
        DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    if (!InstallHook(g_createWindowExWTarget, reinterpret_cast<void*>(&CreateWindowExWDetour),
                     reinterpret_cast<void**>(&g_originalCreateWindowExW), L"CreateWindowExW")) {
        DisableHook(g_trackPopupMenuExTarget, L"TrackPopupMenuEx");
        DisableHook(g_trackPopupMenuTarget, L"TrackPopupMenu");
        DisableHook(g_gdiGradientFillTarget, L"GdiGradientFill");
        DisableHook(g_fillRectTarget, L"FillRect");
        DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
        DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");
        if (initializedHere) {
            MH_Uninitialize();
        }
        return false;
    }

    g_hooksActive = true;
    LogMessage(LogLevel::Info, L"ThemeHooks: installed theme detours");
    return true;
}

void ShutdownThemeHooks() {
    std::lock_guard<std::mutex> guard(g_initializationMutex);
    if (!g_hooksActive) {
        return;
    }

    DisableHook(g_createWindowExWTarget, L"CreateWindowExW");
    DisableHook(g_trackPopupMenuExTarget, L"TrackPopupMenuEx");
    DisableHook(g_trackPopupMenuTarget, L"TrackPopupMenu");
    DisableHook(g_gdiGradientFillTarget, L"GdiGradientFill");
    DisableHook(g_fillRectTarget, L"FillRect");
    DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
    DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");

    {
        std::lock_guard<std::mutex> detourGuard(g_directUiDetourMutex);
        for (void* target : g_directUiDetourTargets) {
            DisableHook(target, L"DirectUiDraw");
        }
        g_directUiDetourTargets.clear();
        g_directUiTargetInfo.clear();
        g_directUiElementInfo.clear();
    }

    MH_STATUS status = MH_Uninitialize();
    if (status != MH_OK && status != MH_ERROR_NOT_INITIALIZED) {
        LogMessage(LogLevel::Warning, L"ThemeHooks: MH_Uninitialize failed (status=%d)", status);
    }

    {
        std::lock_guard<std::mutex> registryGuard(g_registryMutex);
        g_surfaceRegistry.clear();
        g_directUiHosts.clear();
    }

    {
        std::lock_guard<std::mutex> metricsGuard(g_scrollbarMetricsMutex);
        g_scrollbarMetrics.clear();
    }

    g_originalDrawThemeBackground = nullptr;
    g_originalDrawThemeEdge = nullptr;
    g_originalFillRect = nullptr;
    g_originalGdiGradientFill = nullptr;
    g_originalTrackPopupMenu = nullptr;
    g_originalTrackPopupMenuEx = nullptr;
    g_originalCreateWindowExW = nullptr;
    g_drawThemeBackgroundTarget = nullptr;
    g_drawThemeEdgeTarget = nullptr;
    g_fillRectTarget = nullptr;
    g_gdiGradientFillTarget = nullptr;
    g_trackPopupMenuTarget = nullptr;
    g_trackPopupMenuExTarget = nullptr;
    g_createWindowExWTarget = nullptr;
    g_hooksActive = false;
    LogMessage(LogLevel::Info, L"ThemeHooks: detached theme detours");
}

bool AreThemeHooksActive() noexcept { return g_hooksActive; }

void RegisterThemeSurface(HWND hwnd, ExplorerSurfaceKind kind, ExplorerGlowCoordinator* coordinator) noexcept {
    if (!hwnd || !coordinator) {
        return;
    }
    RegisterCompositionSurface(hwnd, coordinator);
    std::lock_guard<std::mutex> guard(g_registryMutex);
    g_surfaceRegistry[hwnd] = ThemeSurfaceRegistration{coordinator, kind};
}

void UnregisterThemeSurface(HWND hwnd) noexcept {
    if (!hwnd) {
        return;
    }
    InvalidateScrollbarMetricsInternal(hwnd);
    UnregisterCompositionSurface(hwnd);
    std::lock_guard<std::mutex> guard(g_registryMutex);
    g_surfaceRegistry.erase(hwnd);
}

ThemePaintOverrideGuard::ThemePaintOverrideGuard(HWND window, ExplorerSurfaceKind kind, GlowColorSet colors,
                                                 bool suppressFallback) noexcept {
    if (ActivateThemePaintOverride(window, kind, colors, suppressFallback)) {
        m_active = true;
    }
}

ThemePaintOverrideGuard::~ThemePaintOverrideGuard() {
    if (m_active) {
        DeactivateThemePaintOverride();
    }
}

ThemePaintOverrideGuard::ThemePaintOverrideGuard(ThemePaintOverrideGuard&& other) noexcept {
    m_active = other.m_active;
    other.m_active = false;
}

ThemePaintOverrideGuard& ThemePaintOverrideGuard::operator=(ThemePaintOverrideGuard&& other) noexcept {
    if (this != &other) {
        if (m_active) {
            DeactivateThemePaintOverride();
        }
        m_active = other.m_active;
        other.m_active = false;
    }
    return *this;
}

}  // namespace shelltabs

