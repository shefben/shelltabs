#include "ThemeHooks.h"

#include "ExplorerGlowSurfaces.h"
#include "Logging.h"

#include <psapi.h>
#include <gdiplus.h>
#include <uxtheme.h>
#include <vssym32.h>

#include <algorithm>
#include <vector>

namespace shelltabs {
namespace {

constexpr BYTE kTrackHaloAlpha = 96;
constexpr BYTE kTrackLineAlpha = 220;
constexpr BYTE kThumbFillAlpha = 210;
constexpr BYTE kThumbHaloAlpha = 110;
constexpr BYTE kSeparatorLineAlpha = 220;
constexpr BYTE kSeparatorHaloAlpha = 96;

bool IsScrollbarPart(int partId) {
    switch (partId) {
        case SBP_THUMBBTNHORZ:
        case SBP_THUMBBTNVERT:
        case SBP_UPPERTRACKHORZ:
        case SBP_UPPERTRACKVERT:
        case SBP_LOWERTRACKHORZ:
        case SBP_LOWERTRACKVERT:
        case SBP_GRIPPERHORZ:
        case SBP_GRIPPERVERT:
        case SBP_ARROWBTN:
            return true;
        default:
            return false;
    }
}

bool IsToolbarSeparatorPart(int partId) {
    switch (partId) {
        case TP_SEPARATOR:
        case TP_SEPARATORVERT:
            return true;
        default:
            return false;
    }
}

bool IsRebarBandPart(int partId) {
    switch (partId) {
        case RP_BAND:
        case RP_BANDVERT:
            return true;
        default:
            return false;
    }
}

bool IsTrackPart(int partId) {
    switch (partId) {
        case SBP_UPPERTRACKHORZ:
        case SBP_UPPERTRACKVERT:
        case SBP_LOWERTRACKHORZ:
        case SBP_LOWERTRACKVERT:
            return true;
        default:
            return false;
    }
}

bool IsThumbPart(int partId) {
    return partId == SBP_THUMBBTNHORZ || partId == SBP_THUMBBTNVERT;
}

bool IsVerticalPart(int partId) {
    switch (partId) {
        case SBP_THUMBBTNVERT:
        case SBP_UPPERTRACKVERT:
        case SBP_LOWERTRACKVERT:
        case SBP_GRIPPERVERT:
            return true;
        default:
            return false;
    }
}

int ScaleByDpi(int value, UINT dpi) {
    if (dpi == 0) {
        dpi = 96;
    }
    return std::max(1, MulDiv(value, static_cast<int>(dpi), 96));
}

RECT IntersectOrEmpty(const RECT& rect, const RECT* clip) {
    if (!clip) {
        return rect;
    }

    RECT intersection{};
    if (IntersectRect(&intersection, &rect, clip)) {
        return intersection;
    }
    return RECT{0, 0, 0, 0};
}

bool IsRectEmptyStrict(const RECT& rect) {
    return rect.left >= rect.right || rect.top >= rect.bottom;
}

void FillGradientRect(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& rect,
                      BYTE alpha, float angle = 90.0f) {
    if (!colors.valid || IsRectEmptyStrict(rect)) {
        return;
    }

    const Gdiplus::Rect gdRect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
    const Gdiplus::Color start(alpha, GetRValue(colors.start), GetGValue(colors.start), GetBValue(colors.start));
    const Gdiplus::Color end(alpha, GetRValue(colors.end), GetGValue(colors.end), GetBValue(colors.end));

    if (!colors.gradient || colors.start == colors.end) {
        Gdiplus::SolidBrush brush(start);
        graphics.FillRectangle(&brush, gdRect);
        return;
    }

    Gdiplus::LinearGradientBrush brush(gdRect, start, end, angle);
    graphics.FillRectangle(&brush, gdRect);
}

void PaintTrack(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& rect,
                const RECT* clip, bool vertical, UINT dpiX, UINT dpiY) {
    RECT clipRect = IntersectOrEmpty(rect, clip);
    if (IsRectEmptyStrict(clipRect)) {
        return;
    }

    FillGradientRect(graphics, colors, clipRect, kTrackHaloAlpha);

    RECT lineRect = rect;
    if (vertical) {
        const LONG width = rect.right - rect.left;
        const LONG available = std::max<LONG>(width, 1);
        const LONG lineThickness =
            std::clamp<LONG>(static_cast<LONG>(ScaleByDpi(2, dpiX)), 1, available);
        const LONG center = rect.left + (width - lineThickness) / 2;
        lineRect.left = center;
        lineRect.right = center + lineThickness;
    } else {
        const LONG height = rect.bottom - rect.top;
        const LONG available = std::max<LONG>(height, 1);
        const LONG lineThickness =
            std::clamp<LONG>(static_cast<LONG>(ScaleByDpi(2, dpiY)), 1, available);
        const LONG center = rect.top + (height - lineThickness) / 2;
        lineRect.top = center;
        lineRect.bottom = center + lineThickness;
    }

    RECT lineClip = IntersectOrEmpty(lineRect, clip);
    if (!IsRectEmptyStrict(lineClip)) {
        FillGradientRect(graphics, colors, lineClip, kTrackLineAlpha);
    }
}

void PaintThumb(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& rect,
                const RECT* clip, bool vertical, UINT dpiX, UINT dpiY) {
    RECT clipRect = IntersectOrEmpty(rect, clip);
    if (IsRectEmptyStrict(clipRect)) {
        return;
    }

    FillGradientRect(graphics, colors, clipRect, kThumbFillAlpha);

    RECT haloRect = rect;
    if (vertical) {
        const int alongPadding = ScaleByDpi(6, dpiY);
        const int crossPadding = ScaleByDpi(3, dpiX);
        InflateRect(&haloRect, crossPadding, alongPadding);
    } else {
        const int alongPadding = ScaleByDpi(6, dpiX);
        const int crossPadding = ScaleByDpi(3, dpiY);
        InflateRect(&haloRect, alongPadding, crossPadding);
    }

    RECT limitRect = rect;
    const RECT* limit = clip ? clip : &limitRect;
    RECT haloClip = IntersectOrEmpty(haloRect, limit);
    if (!IsRectEmptyStrict(haloClip)) {
        FillGradientRect(graphics, colors, haloClip, kThumbHaloAlpha);
    }
}

std::vector<HMODULE> EnumerateProcessModules() {
    std::vector<HMODULE> modules(128);
    DWORD needed = 0;

    HANDLE process = GetCurrentProcess();
    for (;;) {
        if (!EnumProcessModulesEx(process, modules.data(), static_cast<DWORD>(modules.size() * sizeof(HMODULE)),
                                  &needed, LIST_MODULES_ALL)) {
            return {};
        }
        if (needed <= modules.size() * sizeof(HMODULE)) {
            modules.resize(needed / sizeof(HMODULE));
            break;
        }
        modules.resize(needed / sizeof(HMODULE));
    }

    modules.erase(std::remove(modules.begin(), modules.end(), nullptr), modules.end());
    return modules;
}

}  // namespace

ThemeHooks& ThemeHooks::Instance() {
    static ThemeHooks instance;
    return instance;
}

void ThemeHooks::AttachCoordinator(ExplorerGlowCoordinator& coordinator) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_coordinator = &coordinator;
    UpdateActivationLocked();
}

void ThemeHooks::DetachCoordinator(const ExplorerGlowCoordinator& coordinator) {
    std::lock_guard<std::mutex> lock(m_mutex);
    if (m_coordinator == &coordinator) {
        m_coordinator = nullptr;
        UpdateActivationLocked();
    }
}

void ThemeHooks::NotifyCoordinatorUpdated() {
    std::lock_guard<std::mutex> lock(m_mutex);
    UpdateActivationLocked();
}

bool ThemeHooks::IsActive() const noexcept {
    std::lock_guard<std::mutex> lock(m_mutex);
    return m_active;
}

bool ThemeHooks::IsSurfaceHookActive(ExplorerSurfaceKind kind) const noexcept {
    std::lock_guard<std::mutex> lock(m_mutex);
    if (!m_active) {
        return false;
    }

    switch (kind) {
        case ExplorerSurfaceKind::Scrollbar:
            return m_expectScrollbar && m_scrollbarHookEngaged;
        case ExplorerSurfaceKind::Toolbar:
            return m_expectToolbar && m_toolbarHookEngaged;
        case ExplorerSurfaceKind::Rebar:
            return m_expectRebar && m_rebarHookEngaged;
        default:
            return false;
    }
}

HRESULT WINAPI ThemeHooks::HookedDrawThemeBackground(HTHEME theme, HDC dc, int partId, int stateId,
                                                     LPCRECT rect, LPCRECT clipRect) {
    if (ThemeHooks::Instance().OnDrawThemeBackground(theme, dc, partId, stateId, rect, clipRect)) {
        return S_OK;
    }

    auto* original = ThemeHooks::Instance().m_originalDrawThemeBackground;
    if (!original) {
        return E_FAIL;
    }
    return original(theme, dc, partId, stateId, rect, clipRect);
}

HRESULT WINAPI ThemeHooks::HookedDrawThemeEdge(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect,
                                               UINT edge, UINT flags, LPRECT contentRect) {
    if (ThemeHooks::Instance().OnDrawThemeEdge(theme, dc, partId, stateId, rect, edge, flags, contentRect)) {
        return S_OK;
    }

    auto* original = ThemeHooks::Instance().m_originalDrawThemeEdge;
    if (!original) {
        return E_FAIL;
    }
    return original(theme, dc, partId, stateId, rect, edge, flags, contentRect);
}

bool ThemeHooks::OnDrawThemeBackground(HTHEME, HDC dc, int partId, int stateId, LPCRECT rect, LPCRECT clipRect) {
    UNREFERENCED_PARAMETER(stateId);

    if (!dc || !rect) {
        return false;
    }

    if (IsScrollbarPart(partId)) {
        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Scrollbar, colors)) {
            return false;
        }

        const bool vertical = IsVerticalPart(partId);
        const UINT dpiX = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
        const UINT dpiY = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSY));

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        if (IsThumbPart(partId)) {
            PaintThumb(graphics, colors, *rect, clipRect, vertical, dpiX, dpiY);
            return true;
        }

        if (IsTrackPart(partId)) {
            PaintTrack(graphics, colors, *rect, clipRect, vertical, dpiX, dpiY);
            return true;
        }

        return false;
    }

    if (IsToolbarSeparatorPart(partId)) {
        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Toolbar, colors)) {
            return false;
        }

        RECT paintRect = *rect;
        if (clipRect && !IntersectRect(&paintRect, &paintRect, clipRect)) {
            return true;
        }
        if (IsRectEmptyStrict(paintRect)) {
            return true;
        }

        const bool verticalSeparator = (partId == TP_SEPARATOR);
        const UINT dpiX = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
        const UINT dpiY = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSY));
        const UINT crossDpi = verticalSeparator ? dpiX : dpiY;

        const int baseLine = ScaleByDpi(1, crossDpi);
        const int extent = verticalSeparator ? (paintRect.right - paintRect.left)
                                             : (paintRect.bottom - paintRect.top);
        const int lineThickness = std::clamp(baseLine, 1, std::max(extent, 1));
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, crossDpi));

        RECT lineRect = paintRect;
        if (verticalSeparator) {
            const int center = paintRect.left + ((paintRect.right - paintRect.left) - lineThickness) / 2;
            lineRect.left = center;
            lineRect.right = center + lineThickness;
        } else {
            const int center = paintRect.top + ((paintRect.bottom - paintRect.top) - lineThickness) / 2;
            lineRect.top = center;
            lineRect.bottom = center + lineThickness;
        }

        RECT haloRect = paintRect;
        if (verticalSeparator) {
            const int center = (lineRect.left + lineRect.right) / 2;
            const int halfHalo = haloThickness / 2;
            haloRect.left = std::max(paintRect.left, center - halfHalo);
            haloRect.right = std::min(paintRect.right, center + halfHalo);
        } else {
            const int center = (lineRect.top + lineRect.bottom) / 2;
            const int halfHalo = haloThickness / 2;
            haloRect.top = std::max(paintRect.top, center - halfHalo);
            haloRect.bottom = std::min(paintRect.bottom, center + halfHalo);
        }

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        if (!IsRectEmptyStrict(haloRect)) {
            FillGradientRect(graphics, colors, haloRect, kSeparatorHaloAlpha);
        }
        if (!IsRectEmptyStrict(lineRect)) {
            FillGradientRect(graphics, colors, lineRect, kSeparatorLineAlpha);
        }

        return true;
    }

    return false;
}

bool ThemeHooks::OnDrawThemeEdge(HTHEME, HDC dc, int partId, int stateId, LPCRECT rect, UINT edge, UINT flags,
                                 LPRECT contentRect) {
    UNREFERENCED_PARAMETER(stateId);
    UNREFERENCED_PARAMETER(edge);
    UNREFERENCED_PARAMETER(contentRect);

    if (!dc || !rect) {
        return false;
    }

    if (IsScrollbarPart(partId)) {
        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Scrollbar, colors)) {
            return false;
        }
        return true;
    }

    if (!IsRebarBandPart(partId)) {
        return false;
    }

    const bool hasRight = (flags & BF_RIGHT) != 0;
    const bool hasBottom = (flags & BF_BOTTOM) != 0;
    if ((!hasRight && !hasBottom) || (hasRight && hasBottom)) {
        return false;
    }

    GlowColorSet colors{};
    if (!ResolveColorsForHook(ExplorerSurfaceKind::Rebar, colors)) {
        return false;
    }

    RECT paintRect = *rect;
    if (IsRectEmptyStrict(paintRect)) {
        return true;
    }

    const bool verticalEdge = hasRight;
    const UINT dpiX = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
    const UINT dpiY = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSY));
    const UINT crossDpi = verticalEdge ? dpiX : dpiY;

    const int baseLine = ScaleByDpi(1, crossDpi);
    const int extent = verticalEdge ? (paintRect.right - paintRect.left) : (paintRect.bottom - paintRect.top);
    const int lineThickness = std::clamp(baseLine, 1, std::max(extent, 1));
    const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, crossDpi));

    RECT lineRect = paintRect;
    if (verticalEdge) {
        const int center = paintRect.left + ((paintRect.right - paintRect.left) - lineThickness) / 2;
        lineRect.left = center;
        lineRect.right = center + lineThickness;
    } else {
        const int center = paintRect.top + ((paintRect.bottom - paintRect.top) - lineThickness) / 2;
        lineRect.top = center;
        lineRect.bottom = center + lineThickness;
    }

    RECT haloRect = paintRect;
    if (verticalEdge) {
        const int center = (lineRect.left + lineRect.right) / 2;
        const int halfHalo = haloThickness / 2;
        haloRect.left = std::max(paintRect.left, center - halfHalo);
        haloRect.right = std::min(paintRect.right, center + halfHalo);
    } else {
        const int center = (lineRect.top + lineRect.bottom) / 2;
        const int halfHalo = haloThickness / 2;
        haloRect.top = std::max(paintRect.top, center - halfHalo);
        haloRect.bottom = std::min(paintRect.bottom, center + halfHalo);
    }

    Gdiplus::Graphics graphics(dc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    if (!IsRectEmptyStrict(haloRect)) {
        FillGradientRect(graphics, colors, haloRect, kSeparatorHaloAlpha);
    }
    if (!IsRectEmptyStrict(lineRect)) {
        FillGradientRect(graphics, colors, lineRect, kSeparatorLineAlpha);
    }

    return true;
}

bool ThemeHooks::ResolveColorsForHook(ExplorerSurfaceKind kind, GlowColorSet& colors) {
    std::lock_guard<std::mutex> lock(m_mutex);
    if (!m_active || !m_coordinator) {
        return false;
    }

    if (!ExpectHookFor(kind)) {
        return false;
    }

    if (!m_coordinator->ShouldRenderSurface(kind)) {
        return false;
    }

    colors = m_coordinator->ResolveColors(kind);
    if (!colors.valid) {
        return false;
    }

    switch (kind) {
        case ExplorerSurfaceKind::Scrollbar:
            m_scrollbarHookEngaged = true;
            break;
        case ExplorerSurfaceKind::Toolbar:
            m_toolbarHookEngaged = true;
            break;
        case ExplorerSurfaceKind::Rebar:
            m_rebarHookEngaged = true;
            break;
        default:
            break;
    }

    return true;
}

bool ThemeHooks::ExpectHookFor(ExplorerSurfaceKind kind) const noexcept {
    switch (kind) {
        case ExplorerSurfaceKind::Scrollbar:
            return m_expectScrollbar;
        case ExplorerSurfaceKind::Toolbar:
            return m_expectToolbar;
        case ExplorerSurfaceKind::Rebar:
            return m_expectRebar;
        default:
            return false;
    }
}

void ThemeHooks::UpdateActivationLocked() {
    const bool shouldHookScrollbar =
        m_coordinator && m_coordinator->ShouldRenderSurface(ExplorerSurfaceKind::Scrollbar);
    const bool shouldHookToolbar =
        m_coordinator && m_coordinator->ShouldRenderSurface(ExplorerSurfaceKind::Toolbar);
    const bool shouldHookRebar =
        m_coordinator && m_coordinator->ShouldRenderSurface(ExplorerSurfaceKind::Rebar);

    m_expectScrollbar = shouldHookScrollbar;
    m_expectToolbar = shouldHookToolbar;
    m_expectRebar = shouldHookRebar;

    if (!shouldHookScrollbar) {
        m_scrollbarHookEngaged = false;
    }
    if (!shouldHookToolbar) {
        m_toolbarHookEngaged = false;
    }
    if (!shouldHookRebar) {
        m_rebarHookEngaged = false;
    }

    const bool shouldActivate = shouldHookScrollbar || shouldHookToolbar || shouldHookRebar;
    if (!shouldActivate) {
        if (m_active) {
            UninstallLocked();
        }
        m_active = false;
        return;
    }

    const bool wasActive = m_active;
    const bool installed = InstallLocked();
    if (installed || wasActive) {
        if (!wasActive) {
            m_scrollbarHookEngaged = false;
            m_toolbarHookEngaged = false;
            m_rebarHookEngaged = false;
        }
        m_active = true;
    } else {
        m_active = false;
    }
}

bool ThemeHooks::InstallLocked() {
    if (!m_originalDrawThemeBackground || !m_originalDrawThemeEdge) {
        HMODULE themeModule = GetModuleHandleW(L"uxtheme.dll");
        if (!themeModule) {
            themeModule = LoadLibraryW(L"uxtheme.dll");
        }
        if (!themeModule) {
            LogLastError(L"LoadLibraryW(uxtheme.dll)", GetLastError());
            return false;
        }

        m_originalDrawThemeBackground =
            reinterpret_cast<DrawThemeBackgroundFn>(GetProcAddress(themeModule, "DrawThemeBackground"));
        m_originalDrawThemeEdge =
            reinterpret_cast<DrawThemeEdgeFn>(GetProcAddress(themeModule, "DrawThemeEdge"));
    }

    if (!m_originalDrawThemeBackground || !m_originalDrawThemeEdge) {
        return false;
    }

    auto modules = EnumerateProcessModules();
    bool patched = false;
    for (HMODULE module : modules) {
        if (!module) {
            continue;
        }
        patched |= HookModuleImportsLocked(module);
    }

    return patched;
}

void ThemeHooks::UninstallLocked() {
    for (auto& patch : m_backgroundPatches) {
        if (!patch.slot || !patch.original) {
            continue;
        }
        DWORD oldProtect = 0;
        if (VirtualProtect(patch.slot, sizeof(FARPROC), PAGE_EXECUTE_READWRITE, &oldProtect)) {
            *patch.slot = patch.original;
            DWORD ignore = 0;
            VirtualProtect(patch.slot, sizeof(FARPROC), oldProtect, &ignore);
            FlushInstructionCache(GetCurrentProcess(), patch.slot, sizeof(FARPROC));
        }
    }
    for (auto& patch : m_edgePatches) {
        if (!patch.slot || !patch.original) {
            continue;
        }
        DWORD oldProtect = 0;
        if (VirtualProtect(patch.slot, sizeof(FARPROC), PAGE_EXECUTE_READWRITE, &oldProtect)) {
            *patch.slot = patch.original;
            DWORD ignore = 0;
            VirtualProtect(patch.slot, sizeof(FARPROC), oldProtect, &ignore);
            FlushInstructionCache(GetCurrentProcess(), patch.slot, sizeof(FARPROC));
        }
    }

    m_backgroundPatches.clear();
    m_edgePatches.clear();
    m_expectScrollbar = false;
    m_expectToolbar = false;
    m_expectRebar = false;
    m_scrollbarHookEngaged = false;
    m_toolbarHookEngaged = false;
    m_rebarHookEngaged = false;
    m_active = false;
}

bool ThemeHooks::HookModuleImportsLocked(HMODULE module) {
    auto* dos = reinterpret_cast<PIMAGE_DOS_HEADER>(module);
    if (!dos || dos->e_magic != IMAGE_DOS_SIGNATURE) {
        return false;
    }

    auto* nt = reinterpret_cast<PIMAGE_NT_HEADERS>(reinterpret_cast<BYTE*>(module) + dos->e_lfanew);
    if (!nt || nt->Signature != IMAGE_NT_SIGNATURE) {
        return false;
    }

    const auto& importDirectory = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (importDirectory.VirtualAddress == 0) {
        return false;
    }

    auto* descriptor = reinterpret_cast<PIMAGE_IMPORT_DESCRIPTOR>(
        reinterpret_cast<BYTE*>(module) + importDirectory.VirtualAddress);

    bool updated = false;
    for (; descriptor && descriptor->Name; ++descriptor) {
        const char* moduleName = reinterpret_cast<const char*>(reinterpret_cast<BYTE*>(module) + descriptor->Name);
        if (!moduleName || _stricmp(moduleName, "uxtheme.dll") != 0) {
            continue;
        }

        auto* thunk = reinterpret_cast<PIMAGE_THUNK_DATA>(reinterpret_cast<BYTE*>(module) + descriptor->FirstThunk);
        auto* originalThunk = descriptor->OriginalFirstThunk
                                   ? reinterpret_cast<PIMAGE_THUNK_DATA>(
                                         reinterpret_cast<BYTE*>(module) + descriptor->OriginalFirstThunk)
                                   : thunk;

        for (; originalThunk && originalThunk->u1.AddressOfData; ++originalThunk, ++thunk) {
            if (IMAGE_SNAP_BY_ORDINAL(originalThunk->u1.Ordinal)) {
                continue;
            }

            auto* import = reinterpret_cast<PIMAGE_IMPORT_BY_NAME>(
                reinterpret_cast<BYTE*>(module) + originalThunk->u1.AddressOfData);
            if (!import) {
                continue;
            }

            const char* functionName = reinterpret_cast<const char*>(import->Name);
            if (!functionName) {
                continue;
            }

            FARPROC* slot = reinterpret_cast<FARPROC*>(&thunk->u1.Function);
            if (!slot || !*slot) {
                continue;
            }

            if (_stricmp(functionName, "DrawThemeBackground") == 0) {
                if (*slot != reinterpret_cast<FARPROC>(&HookedDrawThemeBackground)) {
                    DWORD oldProtect = 0;
                    if (VirtualProtect(slot, sizeof(FARPROC), PAGE_EXECUTE_READWRITE, &oldProtect)) {
                        PatchedImport patch{};
                        patch.module = module;
                        patch.slot = slot;
                        patch.original = *slot;
                        *slot = reinterpret_cast<FARPROC>(&HookedDrawThemeBackground);
                        DWORD ignore = 0;
                        VirtualProtect(slot, sizeof(FARPROC), oldProtect, &ignore);
                        FlushInstructionCache(GetCurrentProcess(), slot, sizeof(FARPROC));
                        m_backgroundPatches.push_back(patch);
                        updated = true;
                    }
                }
            } else if (_stricmp(functionName, "DrawThemeEdge") == 0) {
                if (*slot != reinterpret_cast<FARPROC>(&HookedDrawThemeEdge)) {
                    DWORD oldProtect = 0;
                    if (VirtualProtect(slot, sizeof(FARPROC), PAGE_EXECUTE_READWRITE, &oldProtect)) {
                        PatchedImport patch{};
                        patch.module = module;
                        patch.slot = slot;
                        patch.original = *slot;
                        *slot = reinterpret_cast<FARPROC>(&HookedDrawThemeEdge);
                        DWORD ignore = 0;
                        VirtualProtect(slot, sizeof(FARPROC), oldProtect, &ignore);
                        FlushInstructionCache(GetCurrentProcess(), slot, sizeof(FARPROC));
                        m_edgePatches.push_back(patch);
                        updated = true;
                    }
                }
            }
        }
    }

    return updated;
}

}  // namespace shelltabs

