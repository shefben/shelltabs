#include "ThemeHooks.h"

#include "ExplorerGlowSurfaces.h"
#include "ExplorerThemeUtils.h"
#include "Logging.h"
#include "OptionsStore.h"

#include <windows.h>
#include <CommCtrl.h>
#include <gdiplus.h>
#include <psapi.h>
#include <uxtheme.h>
#include <vssym32.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <cstddef>
#include <cstdlib>
#include <mutex>
#include <optional>
#include <string>
#include <utility>
#include <vector>

namespace shelltabs {
namespace {

using DrawThemeBackgroundFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, const RECT*, const RECT*);
using DrawThemeEdgeFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, const RECT*, UINT, UINT, RECT*);
using FillRectFn = int(WINAPI*)(HDC, const RECT*, HBRUSH);
using GradientFillFn = BOOL(WINAPI*)(HDC, PTRIVERTEX, ULONG, PVOID, ULONG, ULONG);

constexpr BYTE kLineAlpha = 220;
constexpr BYTE kHaloAlpha = 96;
constexpr BYTE kFrameAlpha = 210;
constexpr BYTE kFrameHaloAlpha = 110;

struct ImportPatch {
    HMODULE module = nullptr;
    void** address = nullptr;
    void* original = nullptr;
};

struct ThemeHookState {
    std::mutex mutex;
    bool installed = false;
    bool attempted = false;
    DrawThemeBackgroundFn drawThemeBackground = nullptr;
    DrawThemeEdgeFn drawThemeEdge = nullptr;
    FillRectFn fillRect = nullptr;
    GradientFillFn gradientFill = nullptr;
    std::vector<ImportPatch> patches;
    ExplorerGlowCoordinator coordinator;
    bool gdiplusInitialized = false;
    ULONG_PTR gdiplusToken = 0;
};

ThemeHookState& GetState() {
    static ThemeHookState state;
    return state;
}

std::atomic<uint32_t> g_surfaceMask{0};
thread_local bool g_inHook = false;

uint32_t SurfaceBit(ExplorerSurfaceKind kind) {
    return 1u << static_cast<unsigned>(kind);
}

class HookReentrancyGuard {
public:
    HookReentrancyGuard() : m_shouldBypass(g_inHook) {
        if (!m_shouldBypass) {
            g_inHook = true;
        }
    }

    ~HookReentrancyGuard() {
        if (!m_shouldBypass) {
            g_inHook = false;
        }
    }

    bool ShouldBypass() const noexcept { return m_shouldBypass; }

private:
    bool m_shouldBypass = false;
};

int ScaleByDpi(int value, UINT dpi) {
    if (dpi == 0) {
        dpi = 96;
    }
    return std::max(1, MulDiv(value, static_cast<int>(dpi), 96));
}

bool MatchesClass(HWND hwnd, const wchar_t* className) {
    if (!hwnd || !className) {
        return false;
    }

    wchar_t buffer[64] = {};
    const int length = GetClassNameW(hwnd, buffer, static_cast<int>(std::size(buffer)));
    if (length <= 0) {
        return false;
    }
    return _wcsicmp(buffer, className) == 0;
}

RECT GetClientRectSafe(HWND hwnd) {
    RECT rect{0, 0, 0, 0};
    if (hwnd && IsWindow(hwnd)) {
        GetClientRect(hwnd, &rect);
    }
    return rect;
}

bool ApproximatelyEqual(int lhs, int rhs, int tolerance) {
    return std::abs(lhs - rhs) <= tolerance;
}

Gdiplus::Rect RectToGdiplus(const RECT& rect) {
    return {rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top};
}

void FillGradientRect(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& rect, BYTE alpha,
                      float angle = 90.0f) {
    if (!colors.valid || rect.right <= rect.left || rect.bottom <= rect.top) {
        return;
    }

    const Gdiplus::Rect gdiplusRect = RectToGdiplus(rect);
    const Gdiplus::Color start(alpha, GetRValue(colors.start), GetGValue(colors.start), GetBValue(colors.start));
    const Gdiplus::Color end(alpha, GetRValue(colors.end), GetGValue(colors.end), GetBValue(colors.end));

    if (!colors.gradient || colors.start == colors.end) {
        Gdiplus::SolidBrush brush(start);
        graphics.FillRectangle(&brush, gdiplusRect);
        return;
    }

    Gdiplus::LinearGradientBrush brush(gdiplusRect, start, end, angle, true);
    graphics.FillRectangle(&brush, gdiplusRect);
}

void FillFrameRegion(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& outerRect, const RECT& innerRect,
                     BYTE alpha, float angle = 90.0f) {
    if (!colors.valid) {
        return;
    }

    RECT topRect = outerRect;
    topRect.bottom = std::min(topRect.bottom, innerRect.top);
    if (topRect.bottom > topRect.top) {
        FillGradientRect(graphics, colors, topRect, alpha, angle);
    }

    RECT bottomRect = outerRect;
    bottomRect.top = std::max(bottomRect.top, innerRect.bottom);
    if (bottomRect.bottom > bottomRect.top) {
        FillGradientRect(graphics, colors, bottomRect, alpha, angle);
    }

    RECT leftRect = outerRect;
    leftRect.right = std::min(leftRect.right, innerRect.left);
    leftRect.top = std::max(leftRect.top, topRect.bottom);
    leftRect.bottom = std::min(leftRect.bottom, bottomRect.top);
    if (leftRect.right > leftRect.left && leftRect.bottom > leftRect.top) {
        FillGradientRect(graphics, colors, leftRect, alpha, 0.0f);
    }

    RECT rightRect = outerRect;
    rightRect.left = std::max(rightRect.left, innerRect.right);
    rightRect.top = std::max(rightRect.top, topRect.bottom);
    rightRect.bottom = std::min(rightRect.bottom, bottomRect.top);
    if (rightRect.right > rightRect.left && rightRect.bottom > rightRect.top) {
        FillGradientRect(graphics, colors, rightRect, alpha, 0.0f);
    }
}

ExplorerSurfaceKind IdentifySurfaceFromWindow(HWND window) {
    if (!window || !IsWindow(window)) {
        return ExplorerSurfaceKind::DirectUi;  // placeholder
    }

    if (MatchesClass(window, REBARCLASSNAMEW)) {
        return ExplorerSurfaceKind::Rebar;
    }
    if (MatchesClass(window, WC_HEADERW)) {
        return ExplorerSurfaceKind::Header;
    }
    if (MatchesClass(window, L"Edit")) {
        return ExplorerSurfaceKind::Edit;
    }
    return ExplorerSurfaceKind::DirectUi;
}

bool ShouldHandleSurfaceLocked(const ThemeHookState& state, ExplorerSurfaceKind kind) {
    switch (kind) {
        case ExplorerSurfaceKind::Rebar:
        case ExplorerSurfaceKind::Header:
        case ExplorerSurfaceKind::Edit:
            return state.coordinator.ShouldRenderSurface(kind);
        default:
            return false;
    }
}

void UpdateSurfaceMaskLocked(const ThemeHookState& state) {
    uint32_t mask = 0;
    if (state.installed) {
        for (ExplorerSurfaceKind kind : {ExplorerSurfaceKind::Rebar, ExplorerSurfaceKind::Header, ExplorerSurfaceKind::Edit}) {
            if (ShouldHandleSurfaceLocked(state, kind)) {
                mask |= SurfaceBit(kind);
            }
        }
    }
    g_surfaceMask.store(mask, std::memory_order_release);
}

std::optional<GlowColorSet> ResolveColorsLocked(ThemeHookState& state, ExplorerSurfaceKind kind) {
    if (!ShouldHandleSurfaceLocked(state, kind)) {
        return std::nullopt;
    }
    GlowColorSet colors = state.coordinator.ResolveColors(kind);
    if (!colors.valid) {
        return std::nullopt;
    }
    return colors;
}

void EnsureGdiplusInitializedLocked(ThemeHookState& state) {
    if (state.gdiplusInitialized) {
        return;
    }

    Gdiplus::GdiplusStartupInput input;
    ULONG_PTR token = 0;
    if (Gdiplus::GdiplusStartup(&token, &input, nullptr) == Gdiplus::Ok) {
        state.gdiplusToken = token;
        state.gdiplusInitialized = true;
    }
}

struct ImportTarget {
    const char* moduleName;
    void* original;
    void* hook;
};

bool PatchModuleImports(HMODULE module, const ImportTarget& target, std::vector<ImportPatch>& patches) {
    if (!module) {
        return true;
    }

    auto* base = reinterpret_cast<std::byte*>(module);
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    if (!dos || dos->e_magic != IMAGE_DOS_SIGNATURE) {
        return true;
    }

    const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS*>(base + dos->e_lfanew);
    if (!nt || nt->Signature != IMAGE_NT_SIGNATURE) {
        return true;
    }

    const auto& importDirectory = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (importDirectory.VirtualAddress == 0 || importDirectory.Size == 0) {
        return true;
    }

    auto* importDescriptor = reinterpret_cast<IMAGE_IMPORT_DESCRIPTOR*>(base + importDirectory.VirtualAddress);
    for (; importDescriptor->Name != 0; ++importDescriptor) {
        const char* importedModuleName = reinterpret_cast<const char*>(base + importDescriptor->Name);
        if (!importedModuleName) {
            continue;
        }
        if (_stricmp(importedModuleName, target.moduleName) != 0) {
            continue;
        }

        auto* thunk = reinterpret_cast<IMAGE_THUNK_DATA*>(base + importDescriptor->FirstThunk);
        for (; thunk && thunk->u1.Function != 0; ++thunk) {
            void** function = reinterpret_cast<void**>(&thunk->u1.Function);
            if (*function != target.original) {
                continue;
            }

            DWORD oldProtect = 0;
            if (!VirtualProtect(function, sizeof(void*), PAGE_READWRITE, &oldProtect)) {
                continue;
            }

            ImportPatch patch{};
            patch.module = module;
            patch.address = function;
            patch.original = target.original;

            *function = target.hook;

            DWORD ignore = 0;
            VirtualProtect(function, sizeof(void*), oldProtect, &ignore);
            FlushInstructionCache(GetCurrentProcess(), function, sizeof(void*));

            patches.push_back(patch);
        }
    }

    return true;
}

bool InstallImportHooksLocked(ThemeHookState& state, const std::vector<ImportTarget>& targets) {
    HANDLE process = GetCurrentProcess();
    DWORD needed = 0;
    if (!EnumProcessModules(process, nullptr, 0, &needed) || needed == 0) {
        return false;
    }

    std::vector<HMODULE> modules(needed / sizeof(HMODULE));
    if (!EnumProcessModules(process, modules.data(), static_cast<DWORD>(modules.size() * sizeof(HMODULE)), &needed)) {
        return false;
    }

    modules.resize(needed / sizeof(HMODULE));
    for (HMODULE module : modules) {
        for (const auto& target : targets) {
            PatchModuleImports(module, target, state.patches);
        }
    }

    return true;
}

bool EnsureOriginalFunctionsLocked(ThemeHookState& state) {
    if (state.drawThemeBackground && state.drawThemeEdge && state.fillRect && state.gradientFill) {
        return true;
    }

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

    if (!uxtheme || !user32 || !gdi32) {
        return false;
    }

    state.drawThemeBackground = reinterpret_cast<DrawThemeBackgroundFn>(GetProcAddress(uxtheme, "DrawThemeBackground"));
    state.drawThemeEdge = reinterpret_cast<DrawThemeEdgeFn>(GetProcAddress(uxtheme, "DrawThemeEdge"));
    state.fillRect = reinterpret_cast<FillRectFn>(GetProcAddress(user32, "FillRect"));
    state.gradientFill = reinterpret_cast<GradientFillFn>(GetProcAddress(gdi32, "GdiGradientFill"));

    return state.drawThemeBackground && state.drawThemeEdge && state.fillRect && state.gradientFill;
}

bool DrawRebarGlow(HWND hwnd, HDC hdc, const RECT& clip, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, REBARCLASSNAMEW)) {
        return false;
    }

    const RECT client = GetClientRectSafe(hwnd);
    RECT paintRect = clip;
    if (!IntersectRect(&paintRect, &clip, &client)) {
        paintRect = client;
    }
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        return false;
    }

    const UINT dpi = GetDpiForWindow(hwnd);
    const int lineThickness = ScaleByDpi(1, dpi);
    const int haloThickness = std::max(lineThickness * 3, ScaleByDpi(4, dpi));
    const int rowTolerance = ScaleByDpi(4, dpi);

    DWORD style = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
    const bool vertical = (style & RBS_VERTICAL) != 0;

    const int bandCount = static_cast<int>(SendMessageW(hwnd, RB_GETBANDCOUNT, 0, 0));
    if (bandCount <= 1) {
        return false;
    }

    Gdiplus::Graphics graphics(hdc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
    graphics.SetClip(RectToGdiplus(paintRect));

    bool havePrevious = false;
    RECT previousRect{};
    int previousAxis = 0;

    for (int index = 0; index < bandCount; ++index) {
        RECT bandRect{};
        if (!SendMessageW(hwnd, RB_GETRECT, index, reinterpret_cast<LPARAM>(&bandRect))) {
            continue;
        }

        REBARBANDINFOW bandInfo{};
        bandInfo.cbSize = sizeof(bandInfo);
        bandInfo.fMask = RBBIM_STYLE;
        if (!SendMessageW(hwnd, RB_GETBANDINFO, index, reinterpret_cast<LPARAM>(&bandInfo))) {
            continue;
        }
        if ((bandInfo.fStyle & RBBS_HIDDEN) != 0) {
            continue;
        }

        if (!havePrevious) {
            previousRect = bandRect;
            previousAxis = vertical ? bandRect.left : bandRect.top;
            havePrevious = true;
            continue;
        }

        const int currentAxis = vertical ? bandRect.left : bandRect.top;
        const bool startsNewRow = (bandInfo.fStyle & RBBS_BREAK) != 0;
        if (startsNewRow || !ApproximatelyEqual(currentAxis, previousAxis, rowTolerance)) {
            previousRect = bandRect;
            previousAxis = currentAxis;
            continue;
        }

        RECT lineRect{};
        if (vertical) {
            const int boundary = previousRect.bottom;
            lineRect.left = std::min(previousRect.left, bandRect.left);
            lineRect.right = std::max(previousRect.right, bandRect.right);
            lineRect.top = boundary - lineThickness / 2;
            lineRect.bottom = lineRect.top + lineThickness;
        } else {
            const int boundary = previousRect.right;
            lineRect.top = std::min(previousRect.top, bandRect.top);
            lineRect.bottom = std::max(previousRect.bottom, bandRect.bottom);
            lineRect.left = boundary - lineThickness / 2;
            lineRect.right = lineRect.left + lineThickness;
        }

        RECT clipped{};
        if (!IntersectRect(&clipped, &lineRect, &paintRect) || clipped.right <= clipped.left || clipped.bottom <= clipped.top) {
            previousRect = bandRect;
            previousAxis = currentAxis;
            continue;
        }

        RECT haloRect = clipped;
        if (vertical) {
            const int centerY = (clipped.top + clipped.bottom) / 2;
            haloRect.top = centerY - haloThickness / 2;
            haloRect.bottom = haloRect.top + haloThickness;
        } else {
            const int centerX = (clipped.left + clipped.right) / 2;
            haloRect.left = centerX - haloThickness / 2;
            haloRect.right = haloRect.left + haloThickness;
        }

        RECT haloClip{};
        if (IntersectRect(&haloClip, &haloRect, &paintRect) && haloClip.right > haloClip.left && haloClip.bottom > haloClip.top) {
            FillGradientRect(graphics, colors, haloClip, kHaloAlpha, vertical ? 0.0f : 90.0f);
        }
        FillGradientRect(graphics, colors, clipped, kLineAlpha, vertical ? 0.0f : 90.0f);

        previousRect = bandRect;
        previousAxis = currentAxis;
    }

    graphics.ResetClip();
    return true;
}

bool DrawHeaderGlow(HWND hwnd, HDC hdc, const RECT& clip, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, WC_HEADERW)) {
        return false;
    }

    RECT client = GetClientRectSafe(hwnd);
    RECT paintRect = clip;
    if (!IntersectRect(&paintRect, &clip, &client)) {
        paintRect = client;
    }
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        return false;
    }

    const UINT dpi = GetDpiForWindow(hwnd);
    const int lineThickness = ScaleByDpi(1, dpi);
    const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, dpi));

    const int columnCount = Header_GetItemCount(hwnd);
    if (columnCount <= 0) {
        return false;
    }

    Gdiplus::Graphics graphics(hdc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
    graphics.SetClip(RectToGdiplus(paintRect));

    for (int index = 0; index < columnCount; ++index) {
        RECT itemRect{};
        if (!Header_GetItemRect(hwnd, index, &itemRect)) {
            continue;
        }

        RECT lineRect{itemRect.right - lineThickness, client.top, itemRect.right, client.bottom};
        RECT clipped{};
        if (!IntersectRect(&clipped, &lineRect, &paintRect) || clipped.right <= clipped.left || clipped.bottom <= clipped.top) {
            continue;
        }

        RECT haloRect{lineRect.left - (haloThickness - lineThickness) / 2, client.top,
                      lineRect.left - (haloThickness - lineThickness) / 2 + haloThickness, client.bottom};
        RECT haloClip{};
        if (IntersectRect(&haloClip, &haloRect, &paintRect) && haloClip.right > haloClip.left && haloClip.bottom > haloClip.top) {
            FillGradientRect(graphics, colors, haloClip, kHaloAlpha, 90.0f);
        }

        FillGradientRect(graphics, colors, clipped, kLineAlpha, 90.0f);
    }

    graphics.ResetClip();
    return true;
}

bool DrawEditGlow(HWND hwnd, HDC hdc, const RECT& clip, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, L"Edit")) {
        return false;
    }

    RECT client = GetClientRectSafe(hwnd);
    RECT paintRect = clip;
    if (!IntersectRect(&paintRect, &clip, &client)) {
        paintRect = client;
    }
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        return false;
    }

    const UINT dpi = GetDpiForWindow(hwnd);
    const int frameThickness = ScaleByDpi(1, dpi);
    const int haloThickness = std::max(frameThickness * 2, ScaleByDpi(3, dpi));

    RECT inner = client;
    InflateRect(&inner, -frameThickness, -frameThickness);
    RECT haloOuter = client;
    InflateRect(&haloOuter, haloThickness, haloThickness);

    Gdiplus::Graphics graphics(hdc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
    graphics.SetClip(RectToGdiplus(paintRect));

    FillFrameRegion(graphics, colors, haloOuter, inner, kFrameHaloAlpha);
    FillFrameRegion(graphics, colors, client, inner, kFrameAlpha);

    graphics.ResetClip();
    return true;
}

bool HandleDrawThemeBackground(ThemeHookState& state, HDC hdc, int partId, ExplorerSurfaceKind kind, const RECT* rect,
                               const RECT* clipRect) {
    if (!rect) {
        return false;
    }

    if (kind == ExplorerSurfaceKind::Rebar) {
        if (partId != RP_BAND && partId != RP_BACKGROUND) {
            return false;
        }
    } else if (kind == ExplorerSurfaceKind::Edit) {
        if (partId != EP_EDITTEXT && partId != EP_BACKGROUND && partId != EP_CARET) {
            return false;
        }
    } else {
        return false;
    }

    auto colors = ResolveColorsLocked(state, kind);
    if (!colors.has_value()) {
        return false;
    }

    EnsureGdiplusInitializedLocked(state);
    if (!state.gdiplusInitialized) {
        return false;
    }

    HWND window = WindowFromDC(hdc);
    if (!window) {
        return false;
    }

    RECT clip = rect ? *rect : RECT{};
    if (clipRect) {
        RECT intersect{};
        if (IntersectRect(&intersect, rect, clipRect)) {
            clip = intersect;
        }
    }

    if (kind == ExplorerSurfaceKind::Rebar) {
        return DrawRebarGlow(window, hdc, clip, colors.value());
    }
    if (kind == ExplorerSurfaceKind::Edit) {
        return DrawEditGlow(window, hdc, clip, colors.value());
    }
    return false;
}

bool HandleDrawThemeEdge(ThemeHookState& state, HDC hdc, ExplorerSurfaceKind kind, const RECT* rect) {
    if (kind != ExplorerSurfaceKind::Header || !rect) {
        return false;
    }

    auto colors = ResolveColorsLocked(state, kind);
    if (!colors.has_value()) {
        return false;
    }

    EnsureGdiplusInitializedLocked(state);
    if (!state.gdiplusInitialized) {
        return false;
    }

    HWND window = WindowFromDC(hdc);
    if (!window) {
        return false;
    }

    return DrawHeaderGlow(window, hdc, *rect, colors.value());
}

bool HandleFillRect(ThemeHookState& state, HDC hdc, const RECT* rect) {
    if (!rect) {
        return false;
    }

    HWND window = WindowFromDC(hdc);
    if (!window) {
        return false;
    }

    ExplorerSurfaceKind kind = IdentifySurfaceFromWindow(window);
    if (kind != ExplorerSurfaceKind::Header) {
        return false;
    }

    auto colors = ResolveColorsLocked(state, kind);
    if (!colors.has_value()) {
        return false;
    }

    EnsureGdiplusInitializedLocked(state);
    if (!state.gdiplusInitialized) {
        return false;
    }

    return DrawHeaderGlow(window, hdc, *rect, colors.value());
}

COLOR16 ToColorChannel(BYTE value) {
    return static_cast<COLOR16>(value) * 257;
}

BOOL HandleGradientFill(ThemeHookState& state, HDC hdc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh, ULONG meshCount,
                        ULONG mode) {
    if (!vertices || vertexCount < 2 || meshCount == 0) {
        return FALSE;
    }

    if (mode != GRADIENT_FILL_RECT_H && mode != GRADIENT_FILL_RECT_V) {
        return FALSE;
    }

    HWND window = WindowFromDC(hdc);
    if (!window) {
        return FALSE;
    }

    ExplorerSurfaceKind kind = IdentifySurfaceFromWindow(window);
    if (kind != ExplorerSurfaceKind::Rebar && kind != ExplorerSurfaceKind::Header) {
        return FALSE;
    }

    auto colors = ResolveColorsLocked(state, kind);
    if (!colors.has_value()) {
        return FALSE;
    }

    std::vector<TRIVERTEX> adjusted(vertices, vertices + vertexCount);
    const COLORREF start = colors->start;
    const COLORREF end = colors->gradient ? colors->end : colors->start;

    adjusted[0].Red = ToColorChannel(GetRValue(start));
    adjusted[0].Green = ToColorChannel(GetGValue(start));
    adjusted[0].Blue = ToColorChannel(GetBValue(start));

    adjusted[vertexCount - 1].Red = ToColorChannel(GetRValue(end));
    adjusted[vertexCount - 1].Green = ToColorChannel(GetGValue(end));
    adjusted[vertexCount - 1].Blue = ToColorChannel(GetBValue(end));

    return state.gradientFill(hdc, adjusted.data(), vertexCount, mesh, meshCount, mode);
}

HRESULT WINAPI DrawThemeBackgroundHook(HTHEME theme, HDC hdc, int partId, int stateId, const RECT* rect, const RECT* clipRect) {
    ThemeHookState& state = GetState();
    HookReentrancyGuard guard;
    if (guard.ShouldBypass() || !state.installed || !hdc) {
        return state.drawThemeBackground(theme, hdc, partId, stateId, rect, clipRect);
    }

    HWND window = WindowFromDC(hdc);
    ExplorerSurfaceKind kind = IdentifySurfaceFromWindow(window);

    std::lock_guard<std::mutex> lock(state.mutex);
    if (HandleDrawThemeBackground(state, hdc, partId, kind, rect, clipRect)) {
        return S_OK;
    }
    return state.drawThemeBackground(theme, hdc, partId, stateId, rect, clipRect);
}

HRESULT WINAPI DrawThemeEdgeHook(HTHEME theme, HDC hdc, int partId, int stateId, const RECT* rect, UINT edge, UINT flags,
                                 RECT* contentRect) {
    UNREFERENCED_PARAMETER(edge);
    UNREFERENCED_PARAMETER(flags);
    ThemeHookState& state = GetState();
    HookReentrancyGuard guard;
    if (guard.ShouldBypass() || !state.installed || !hdc) {
        return state.drawThemeEdge(theme, hdc, partId, stateId, rect, edge, flags, contentRect);
    }

    HWND window = WindowFromDC(hdc);
    ExplorerSurfaceKind kind = IdentifySurfaceFromWindow(window);

    std::lock_guard<std::mutex> lock(state.mutex);
    if (HandleDrawThemeEdge(state, hdc, kind, rect)) {
        if (contentRect && rect) {
            *contentRect = *rect;
        }
        return S_OK;
    }
    return state.drawThemeEdge(theme, hdc, partId, stateId, rect, edge, flags, contentRect);
}

int WINAPI FillRectHook(HDC hdc, const RECT* rect, HBRUSH brush) {
    ThemeHookState& state = GetState();
    HookReentrancyGuard guard;
    if (guard.ShouldBypass() || !state.installed || !hdc) {
        return state.fillRect(hdc, rect, brush);
    }

    std::lock_guard<std::mutex> lock(state.mutex);
    if (HandleFillRect(state, hdc, rect)) {
        return 1;
    }
    return state.fillRect(hdc, rect, brush);
}

BOOL WINAPI GradientFillHook(HDC hdc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh, ULONG meshCount, ULONG mode) {
    ThemeHookState& state = GetState();
    HookReentrancyGuard guard;
    if (guard.ShouldBypass() || !state.installed || !hdc) {
        return state.gradientFill(hdc, vertices, vertexCount, mesh, meshCount, mode);
    }

    std::lock_guard<std::mutex> lock(state.mutex);
    if (HandleGradientFill(state, hdc, vertices, vertexCount, mesh, meshCount, mode)) {
        return TRUE;
    }
    return state.gradientFill(hdc, vertices, vertexCount, mesh, meshCount, mode);
}

void RevertPatchesLocked(ThemeHookState& state) {
    for (auto it = state.patches.rbegin(); it != state.patches.rend(); ++it) {
        if (!it->module || !it->address || !it->original) {
            continue;
        }

        DWORD oldProtect = 0;
        if (!VirtualProtect(it->address, sizeof(void*), PAGE_READWRITE, &oldProtect)) {
            continue;
        }

        *it->address = it->original;
        DWORD ignore = 0;
        VirtualProtect(it->address, sizeof(void*), oldProtect, &ignore);
        FlushInstructionCache(GetCurrentProcess(), it->address, sizeof(void*));
    }
    state.patches.clear();
}

}  // namespace

bool InitializeThemeHooks(const ShellTabsOptions& options) {
    ThemeHookState& state = GetState();
    std::lock_guard<std::mutex> lock(state.mutex);

    state.coordinator.Configure(options);
    UpdateSurfaceMaskLocked(state);

    if (state.installed) {
        return true;
    }

    if (!EnsureOriginalFunctionsLocked(state)) {
        LogMessage(LogLevel::Warning, L"ThemeHooks failed to resolve system functions; hooks unavailable");
        return false;
    }

    const std::vector<ImportTarget> targets = {
        {"uxtheme.dll", reinterpret_cast<void*>(state.drawThemeBackground), reinterpret_cast<void*>(&DrawThemeBackgroundHook)},
        {"uxtheme.dll", reinterpret_cast<void*>(state.drawThemeEdge), reinterpret_cast<void*>(&DrawThemeEdgeHook)},
        {"user32.dll", reinterpret_cast<void*>(state.fillRect), reinterpret_cast<void*>(&FillRectHook)},
        {"gdi32.dll", reinterpret_cast<void*>(state.gradientFill), reinterpret_cast<void*>(&GradientFillHook)},
    };

    if (!InstallImportHooksLocked(state, targets)) {
        LogMessage(LogLevel::Warning, L"ThemeHooks failed to patch process imports; hooks unavailable");
        state.patches.clear();
        return false;
    }

    state.installed = true;
    UpdateSurfaceMaskLocked(state);
    return true;
}

void ShutdownThemeHooks() {
    ThemeHookState& state = GetState();
    std::lock_guard<std::mutex> lock(state.mutex);

    if (!state.installed) {
        g_surfaceMask.store(0, std::memory_order_release);
        return;
    }

    RevertPatchesLocked(state);
    state.installed = false;
    g_surfaceMask.store(0, std::memory_order_release);

    if (state.gdiplusInitialized) {
        Gdiplus::GdiplusShutdown(state.gdiplusToken);
        state.gdiplusInitialized = false;
        state.gdiplusToken = 0;
    }
}

void UpdateThemeHooks(const ShellTabsOptions& options) {
    ThemeHookState& state = GetState();
    std::lock_guard<std::mutex> lock(state.mutex);

    state.coordinator.Configure(options);
    UpdateSurfaceMaskLocked(state);
}

void NotifyThemeHooksThemeChanged() {
    ThemeHookState& state = GetState();
    std::lock_guard<std::mutex> lock(state.mutex);
    if (!state.installed) {
        return;
    }
    if (state.coordinator.HandleThemeChanged()) {
        UpdateSurfaceMaskLocked(state);
    }
}

void NotifyThemeHooksSettingChanged() {
    ThemeHookState& state = GetState();
    std::lock_guard<std::mutex> lock(state.mutex);
    if (!state.installed) {
        return;
    }
    if (state.coordinator.HandleSettingChanged()) {
        UpdateSurfaceMaskLocked(state);
    }
}

bool ThemeHooksOverrideSurface(ExplorerSurfaceKind kind) noexcept {
    const uint32_t mask = g_surfaceMask.load(std::memory_order_acquire);
    return (mask & SurfaceBit(kind)) != 0;
}

}  // namespace shelltabs

