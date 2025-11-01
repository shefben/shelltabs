#include "ThemeHooks.h"

#include <MinHook.h>

#include <CommCtrl.h>
#include <optional>
#include <mutex>
#include <unordered_map>
#include <utility>
#include <algorithm>

#include <windowsx.h>
#include <uxtheme.h>
#include <wingdi.h>

#include "Logging.h"

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

std::mutex g_registryMutex;
std::unordered_map<HWND, ThemeSurfaceRegistration, HwndHasher> g_surfaceRegistry;

std::mutex g_initializationMutex;
bool g_hooksActive = false;

using DrawThemeBackgroundFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, const RECT*, const RECT*);
using DrawThemeEdgeFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, const RECT*, UINT, UINT, RECT*);
using FillRectFn = int(WINAPI*)(HDC, const RECT*, HBRUSH);
using GdiGradientFillFn = BOOL(WINAPI*)(HDC, PTRIVERTEX, ULONG, PVOID, ULONG, ULONG);

DrawThemeBackgroundFn g_originalDrawThemeBackground = nullptr;
DrawThemeEdgeFn g_originalDrawThemeEdge = nullptr;
FillRectFn g_originalFillRect = nullptr;
GdiGradientFillFn g_originalGdiGradientFill = nullptr;

void* g_drawThemeBackgroundTarget = nullptr;
void* g_drawThemeEdgeTarget = nullptr;
void* g_fillRectTarget = nullptr;
void* g_gdiGradientFillTarget = nullptr;

thread_local bool g_drawThemeBackgroundActive = false;
thread_local bool g_drawThemeEdgeActive = false;
thread_local bool g_fillRectActive = false;
thread_local bool g_gradientFillActive = false;

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

struct SurfaceLookupResult {
    ThemeSurfaceRegistration registration;
    HWND window = nullptr;
};

std::optional<ThemeSurfaceRegistration> LookupRegistration(HWND hwnd) {
    std::lock_guard<std::mutex> guard(g_registryMutex);
    auto it = g_surfaceRegistry.find(hwnd);
    if (it == g_surfaceRegistry.end()) {
        return std::nullopt;
    }
    return it->second;
}

std::optional<SurfaceLookupResult> ResolveSurfaceFromDc(HDC dc) {
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

HRESULT WINAPI DrawThemeBackgroundDetour(HTHEME theme, HDC dc, int partId, int stateId, const RECT* rect,
                                         const RECT* clipRect) {
    ReentrancyGuard guard(g_drawThemeBackgroundActive);
    if (!guard.Entered()) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }
    if (!rect) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    auto surface = ResolveSurfaceFromDc(dc);
    if (!surface.has_value()) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    GlowColorSet colors = surface->registration.coordinator->ResolveColors(surface->registration.kind);
    if (!colors.valid) {
        return g_originalDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect);
    }

    RECT paintRect = *rect;
    if (clipRect) {
        RECT clipped{};
        if (!IntersectRect(&clipped, &paintRect, clipRect) || clipped.right <= clipped.left ||
            clipped.bottom <= clipped.top) {
            return S_OK;
        }
        paintRect = clipped;
    }

    if (!PaintGlowSurface(dc, surface->window, paintRect, colors, surface->registration.kind)) {
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

    auto surface = ResolveSurfaceFromDc(dc);
    if (!surface.has_value()) {
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }

    GlowColorSet colors = surface->registration.coordinator->ResolveColors(surface->registration.kind);
    if (!colors.valid) {
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }

    RECT paintRect = *rect;
    if (!PaintGlowSurface(dc, surface->window, paintRect, colors, surface->registration.kind)) {
        return g_originalDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect);
    }
    if (contentRect) {
        *contentRect = paintRect;
    }
    return S_OK;
}

int WINAPI FillRectDetour(HDC dc, const RECT* rect, HBRUSH brush) {
    ReentrancyGuard guard(g_fillRectActive);
    if (!guard.Entered()) {
        return g_originalFillRect(dc, rect, brush);
    }

    auto surface = ResolveSurfaceFromDc(dc);
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

    auto surface = ResolveSurfaceFromDc(dc);
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

    g_hooksActive = true;
    LogMessage(LogLevel::Info, L"ThemeHooks: installed theme detours");
    return true;
}

void ShutdownThemeHooks() {
    std::lock_guard<std::mutex> guard(g_initializationMutex);
    if (!g_hooksActive) {
        return;
    }

    DisableHook(g_gdiGradientFillTarget, L"GdiGradientFill");
    DisableHook(g_fillRectTarget, L"FillRect");
    DisableHook(g_drawThemeEdgeTarget, L"DrawThemeEdge");
    DisableHook(g_drawThemeBackgroundTarget, L"DrawThemeBackground");

    MH_STATUS status = MH_Uninitialize();
    if (status != MH_OK && status != MH_ERROR_NOT_INITIALIZED) {
        LogMessage(LogLevel::Warning, L"ThemeHooks: MH_Uninitialize failed (status=%d)", status);
    }

    {
        std::lock_guard<std::mutex> registryGuard(g_registryMutex);
        g_surfaceRegistry.clear();
    }

    g_originalDrawThemeBackground = nullptr;
    g_originalDrawThemeEdge = nullptr;
    g_originalFillRect = nullptr;
    g_originalGdiGradientFill = nullptr;
    g_drawThemeBackgroundTarget = nullptr;
    g_drawThemeEdgeTarget = nullptr;
    g_fillRectTarget = nullptr;
    g_gdiGradientFillTarget = nullptr;
    g_hooksActive = false;
    LogMessage(LogLevel::Info, L"ThemeHooks: detached theme detours");
}

void RegisterThemeSurface(HWND hwnd, ExplorerSurfaceKind kind, ExplorerGlowCoordinator* coordinator) noexcept {
    if (!hwnd || !coordinator) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_registryMutex);
    g_surfaceRegistry[hwnd] = ThemeSurfaceRegistration{coordinator, kind};
}

void UnregisterThemeSurface(HWND hwnd) noexcept {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_registryMutex);
    g_surfaceRegistry.erase(hwnd);
}

}  // namespace shelltabs

