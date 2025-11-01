#include "ThemeHooks.h"

#include "ExplorerGlowSurfaces.h"
#include "Logging.h"

#include "MinHook.h"

#include <gdiplus.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <CommCtrl.h>

#include <algorithm>
#include <vector>
#include <cwchar>

namespace shelltabs {
namespace {

constexpr BYTE kTrackHaloAlpha = 96;
constexpr BYTE kTrackLineAlpha = 220;
constexpr BYTE kThumbFillAlpha = 210;
constexpr BYTE kThumbHaloAlpha = 110;
constexpr BYTE kSeparatorLineAlpha = 220;
constexpr BYTE kSeparatorHaloAlpha = 96;
constexpr BYTE kRebarFillAlpha = 110;
constexpr BYTE kHeaderLineAlpha = kSeparatorLineAlpha;
constexpr BYTE kHeaderHaloAlpha = kSeparatorHaloAlpha;
constexpr BYTE kEditFrameAlpha = kThumbFillAlpha;
constexpr BYTE kEditHaloAlpha = kThumbHaloAlpha;

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

bool IsHeaderPart(int partId) {
    switch (partId) {
        case HP_HEADERITEM:
        case HP_HEADERITEMLEFT:
        case HP_HEADERITEMRIGHT:
        case HP_HEADERSORTARROW:
            return true;
        default:
            return false;
    }
}

bool IsEditBorderPart(int partId) {
    switch (partId) {
        case EP_EDITBORDER_NOSCROLL:
        case EP_EDITBORDER_HSCROLL:
        case EP_EDITBORDER_VSCROLL:
        case EP_EDITBORDER_HVSCROLL:
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

Gdiplus::Rect RectToGdiplus(const RECT& rect) {
    return {rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top};
}

void FillFrameRegion(Gdiplus::Graphics& graphics, const GlowColorSet& colors, const RECT& outerRect,
                     const RECT& innerRect, BYTE alpha, float angle = 90.0f) {
    if (outerRect.left >= outerRect.right || outerRect.top >= outerRect.bottom) {
        return;
    }

    RECT clippedInner = innerRect;
    clippedInner.left = std::max(clippedInner.left, outerRect.left);
    clippedInner.top = std::max(clippedInner.top, outerRect.top);
    clippedInner.right = std::min(clippedInner.right, outerRect.right);
    clippedInner.bottom = std::min(clippedInner.bottom, outerRect.bottom);

    const LONG width = std::max<LONG>(outerRect.right - outerRect.left, 0);
    const LONG height = std::max<LONG>(outerRect.bottom - outerRect.top, 0);
    if (width == 0 || height == 0) {
        return;
    }

    FillGradientRect(graphics, colors, outerRect, alpha, angle);

    const LONG innerWidth = std::max<LONG>(clippedInner.right - clippedInner.left, 0);
    const LONG innerHeight = std::max<LONG>(clippedInner.bottom - clippedInner.top, 0);
    if (innerWidth <= 0 || innerHeight <= 0) {
        return;
    }

    Gdiplus::Region innerRegion(RectToGdiplus(clippedInner));
    graphics.ExcludeClip(&innerRegion);
    FillGradientRect(graphics, colors, outerRect, alpha, angle);
    graphics.ResetClip();
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

}  // namespace

ThemeHooks& ThemeHooks::Instance() {
    static ThemeHooks instance;
    return instance;
}

bool ThemeHooks::Initialize() {
    std::lock_guard<std::mutex> lock(m_mutex);
    return EnsureHooksInitializedLocked();
}

void ThemeHooks::Shutdown() {
    std::lock_guard<std::mutex> lock(m_mutex);
    DisableHooksLocked();
    DestroyHooksLocked();
    m_coordinator = nullptr;
    m_active = false;
    m_expectScrollbar = false;
    m_expectToolbar = false;
    m_expectRebar = false;
    m_expectHeader = false;
    m_expectEdit = false;
    m_scrollbarHookEngaged = false;
    m_toolbarHookEngaged = false;
    m_rebarHookEngaged = false;
    m_headerHookEngaged = false;
    m_editHookEngaged = false;
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
        case ExplorerSurfaceKind::Header:
            return m_expectHeader && m_headerHookEngaged;
        case ExplorerSurfaceKind::Edit:
            return m_expectEdit && m_editHookEngaged;
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

int WINAPI ThemeHooks::HookedFillRect(HDC dc, const RECT* rect, HBRUSH brush) {
    if (rect && ThemeHooks::Instance().OnFillRect(dc, *rect, brush)) {
        return 1;
    }

    auto* original = ThemeHooks::Instance().m_originalFillRect;
    if (!original) {
        return 0;
    }
    return original(dc, rect, brush);
}

BOOL WINAPI ThemeHooks::HookedGdiGradientFill(HDC dc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh,
                                              ULONG meshCount, ULONG mode) {
    if (ThemeHooks::Instance().OnGradientFill(dc, vertices, vertexCount, mesh, meshCount, mode)) {
        return TRUE;
    }

    auto* original = ThemeHooks::Instance().m_originalGradientFill;
    if (!original) {
        return FALSE;
    }
    return original(dc, vertices, vertexCount, mesh, meshCount, mode);
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

    if (IsRebarBandPart(partId)) {
        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Rebar, colors)) {
            return false;
        }

        RECT paintRect = *rect;
        if (clipRect) {
            RECT intersect{};
            if (IntersectRect(&intersect, &paintRect, clipRect)) {
                paintRect = intersect;
            } else {
                return true;
            }
        }

        if (IsRectEmptyStrict(paintRect)) {
            return true;
        }

        const bool verticalBand = (partId == RP_BANDVERT);
        const float angle = verticalBand ? 0.0f : 90.0f;

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
        FillGradientRect(graphics, colors, paintRect, kRebarFillAlpha, angle);
        return true;
    }

    if (IsHeaderPart(partId)) {
        HWND target = WindowFromDC(dc);
        if (!MatchesClass(target, WC_HEADERW)) {
            return false;
        }

        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Header, colors)) {
            return false;
        }

        RECT paintRect = *rect;
        if (clipRect) {
            RECT intersect{};
            if (IntersectRect(&intersect, &paintRect, clipRect)) {
                paintRect = intersect;
            } else {
                return true;
            }
        }

        if (IsRectEmptyStrict(paintRect)) {
            return true;
        }

        const UINT dpiX = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
        const int lineThickness = std::clamp(ScaleByDpi(1, dpiX), 1, paintRect.right - paintRect.left);
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, dpiX));

        RECT lineRect = paintRect;
        lineRect.left = std::max(lineRect.left, lineRect.right - lineThickness);

        RECT haloRect = paintRect;
        const int haloOffset = (haloThickness - lineThickness) / 2;
        haloRect.left = std::max(paintRect.left, lineRect.left - haloOffset);

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        FillGradientRect(graphics, colors, haloRect, kHeaderHaloAlpha, 90.0f);
        FillGradientRect(graphics, colors, lineRect, kHeaderLineAlpha, 90.0f);
        return true;
    }

    if (IsEditBorderPart(partId)) {
        HWND target = WindowFromDC(dc);
        if (!MatchesClass(target, L"Edit")) {
            return false;
        }

        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Edit, colors)) {
            return false;
        }

        RECT paintRect = *rect;
        if (clipRect) {
            RECT intersect{};
            if (IntersectRect(&intersect, &paintRect, clipRect)) {
                paintRect = intersect;
            } else {
                return true;
            }
        }

        if (IsRectEmptyStrict(paintRect)) {
            return true;
        }

        const UINT dpiX = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
        const UINT dpiY = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSY));

        RECT inner = paintRect;
        const int frameThicknessX = ScaleByDpi(1, dpiX);
        const int frameThicknessY = ScaleByDpi(1, dpiY);
        if ((inner.right - inner.left) > frameThicknessX * 2) {
            inner.left += frameThicknessX;
            inner.right -= frameThicknessX;
        }
        if ((inner.bottom - inner.top) > frameThicknessY * 2) {
            inner.top += frameThicknessY;
            inner.bottom -= frameThicknessY;
        }

        RECT haloRect = paintRect;
        InflateRect(&haloRect, ScaleByDpi(3, dpiX), ScaleByDpi(3, dpiY));

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        RECT haloPaint = IntersectOrEmpty(haloRect, clipRect);
        if (!IsRectEmptyStrict(haloPaint)) {
            FillFrameRegion(graphics, colors, haloPaint, paintRect, kEditHaloAlpha);
        }

        RECT framePaint = IntersectOrEmpty(paintRect, clipRect);
        if (!IsRectEmptyStrict(framePaint)) {
            FillFrameRegion(graphics, colors, framePaint, inner, kEditFrameAlpha);
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

bool ThemeHooks::OnFillRect(HDC dc, const RECT& rect, HBRUSH brush) {
    UNREFERENCED_PARAMETER(brush);

    if (!dc || IsRectEmptyStrict(rect)) {
        return false;
    }

    HWND target = WindowFromDC(dc);
    if (!target) {
        return false;
    }

    if (MatchesClass(target, L"ReBarWindow32")) {
        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Rebar, colors)) {
            return false;
        }

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
        FillGradientRect(graphics, colors, rect, kRebarFillAlpha, 90.0f);
        return true;
    }

    if (MatchesClass(target, L"Edit")) {
        GlowColorSet colors{};
        if (!ResolveColorsForHook(ExplorerSurfaceKind::Edit, colors)) {
            return false;
        }

        const UINT dpiX = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
        const UINT dpiY = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSY));

        RECT inner = rect;
        const int frameThicknessX = ScaleByDpi(1, dpiX);
        const int frameThicknessY = ScaleByDpi(1, dpiY);
        if ((inner.right - inner.left) > frameThicknessX * 2) {
            inner.left += frameThicknessX;
            inner.right -= frameThicknessX;
        }
        if ((inner.bottom - inner.top) > frameThicknessY * 2) {
            inner.top += frameThicknessY;
            inner.bottom -= frameThicknessY;
        }

        Gdiplus::Graphics graphics(dc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
        FillFrameRegion(graphics, colors, rect, inner, kEditFrameAlpha);
        return true;
    }

    return false;
}

bool ThemeHooks::OnGradientFill(HDC dc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh, ULONG meshCount,
                                ULONG mode) {
    if (!dc || !vertices || !mesh || vertexCount == 0 || meshCount == 0) {
        return false;
    }

    HWND target = WindowFromDC(dc);
    if (!target || !MatchesClass(target, L"ReBarWindow32")) {
        return false;
    }

    GlowColorSet colors{};
    if (!ResolveColorsForHook(ExplorerSurfaceKind::Rebar, colors)) {
        return false;
    }

    if (mode != GRADIENT_FILL_RECT_H && mode != GRADIENT_FILL_RECT_V) {
        return false;
    }

    const auto* rects = reinterpret_cast<const GRADIENT_RECT*>(mesh);
    if (!rects) {
        return false;
    }

    const GRADIENT_RECT& first = rects[0];
    if (first.UpperLeft >= vertexCount || first.LowerRight >= vertexCount) {
        return false;
    }

    const TRIVERTEX& topLeft = vertices[first.UpperLeft];
    const TRIVERTEX& bottomRight = vertices[first.LowerRight];
    RECT fillRect{topLeft.x, topLeft.y, bottomRight.x, bottomRight.y};
    if (IsRectEmptyStrict(fillRect)) {
        return false;
    }

    const float angle = (mode == GRADIENT_FILL_RECT_H) ? 0.0f : 90.0f;
    Gdiplus::Graphics graphics(dc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
    FillGradientRect(graphics, colors, fillRect, kRebarFillAlpha, angle);
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
        case ExplorerSurfaceKind::Header:
            m_headerHookEngaged = true;
            break;
        case ExplorerSurfaceKind::Edit:
            m_editHookEngaged = true;
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
        case ExplorerSurfaceKind::Header:
            return m_expectHeader;
        case ExplorerSurfaceKind::Edit:
            return m_expectEdit;
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
    const bool shouldHookHeader =
        m_coordinator && m_coordinator->ShouldRenderSurface(ExplorerSurfaceKind::Header);
    const bool shouldHookEdit =
        m_coordinator && m_coordinator->ShouldRenderSurface(ExplorerSurfaceKind::Edit);

    m_expectScrollbar = shouldHookScrollbar;
    m_expectToolbar = shouldHookToolbar;
    m_expectRebar = shouldHookRebar;
    m_expectHeader = shouldHookHeader;
    m_expectEdit = shouldHookEdit;

    if (!shouldHookScrollbar) {
        m_scrollbarHookEngaged = false;
    }
    if (!shouldHookToolbar) {
        m_toolbarHookEngaged = false;
    }
    if (!shouldHookRebar) {
        m_rebarHookEngaged = false;
    }
    if (!shouldHookHeader) {
        m_headerHookEngaged = false;
    }
    if (!shouldHookEdit) {
        m_editHookEngaged = false;
    }

    const bool shouldActivate =
        shouldHookScrollbar || shouldHookToolbar || shouldHookRebar || shouldHookHeader || shouldHookEdit;
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
            m_headerHookEngaged = false;
            m_editHookEngaged = false;
        }
        m_active = true;
    } else {
        m_active = false;
    }
}

bool ThemeHooks::InstallLocked() {
    if (!EnsureHooksInitializedLocked()) {
        return false;
    }

    if (m_hooksEnabled) {
        return true;
    }

    auto enable = [&](void* target) {
        if (!target) {
            return true;
        }
        const MH_STATUS status = MH_EnableHook(target);
        if (status != MH_OK && status != MH_ERROR_ENABLED) {
            LogMessage(LogLevel::Error, L"MH_EnableHook failed for target %p (status=%d)", target,
                       static_cast<int>(status));
            return false;
        }
        return true;
    };

    if (!enable(m_drawThemeBackgroundTarget) || !enable(m_drawThemeEdgeTarget) || !enable(m_fillRectTarget) ||
        !enable(m_gradientFillTarget)) {
        return false;
    }

    m_hooksEnabled = true;
    return true;
}

void ThemeHooks::UninstallLocked() {
    DisableHooksLocked();

    m_expectScrollbar = false;
    m_expectToolbar = false;
    m_expectRebar = false;
    m_expectHeader = false;
    m_expectEdit = false;

    m_scrollbarHookEngaged = false;
    m_toolbarHookEngaged = false;
    m_rebarHookEngaged = false;
    m_headerHookEngaged = false;
    m_editHookEngaged = false;

    m_active = false;
}

bool ThemeHooks::EnsureHooksInitializedLocked() {
    if (m_hooksInitialized) {
        return true;
    }

    bool initializedLibrary = false;
    MH_STATUS status = MH_Initialize();
    if (status == MH_OK) {
        initializedLibrary = true;
    } else if (status != MH_ERROR_ALREADY_INITIALIZED) {
        LogMessage(LogLevel::Error, L"MH_Initialize failed (status=%d)", static_cast<int>(status));
        return false;
    }

    auto ensureModule = [](const wchar_t* name) -> HMODULE {
        HMODULE module = GetModuleHandleW(name);
        if (!module) {
            module = LoadLibraryW(name);
        }
        return module;
    };

    HMODULE themeModule = ensureModule(L"uxtheme.dll");
    HMODULE userModule = ensureModule(L"user32.dll");
    HMODULE gdiModule = ensureModule(L"gdi32.dll");

    if (!themeModule || !userModule || !gdiModule) {
        LogMessage(LogLevel::Error, L"Failed to resolve required theme modules");
        if (initializedLibrary) {
            MH_Uninitialize();
        }
        return false;
    }

    auto resolveTarget = [&](HMODULE module, const char* name, void** storage) -> bool {
        FARPROC proc = GetProcAddress(module, name);
        if (!proc) {
            LogMessage(LogLevel::Error, L"GetProcAddress failed for %hs", name);
            return false;
        }
        *storage = reinterpret_cast<void*>(proc);
        return true;
    };

    if (!resolveTarget(themeModule, "DrawThemeBackground", &m_drawThemeBackgroundTarget) ||
        !resolveTarget(themeModule, "DrawThemeEdge", &m_drawThemeEdgeTarget) ||
        !resolveTarget(userModule, "FillRect", &m_fillRectTarget) ||
        !resolveTarget(gdiModule, "GdiGradientFill", &m_gradientFillTarget)) {
        if (initializedLibrary) {
            MH_Uninitialize();
        }
        m_drawThemeBackgroundTarget = nullptr;
        m_drawThemeEdgeTarget = nullptr;
        m_fillRectTarget = nullptr;
        m_gradientFillTarget = nullptr;
        return false;
    }

    MH_STATUS backgroundStatus = MH_CreateHook(m_drawThemeBackgroundTarget,
                                               reinterpret_cast<LPVOID>(&HookedDrawThemeBackground),
                                               reinterpret_cast<LPVOID*>(&m_originalDrawThemeBackground));
    MH_STATUS edgeStatus = MH_CreateHook(m_drawThemeEdgeTarget, reinterpret_cast<LPVOID>(&HookedDrawThemeEdge),
                                         reinterpret_cast<LPVOID*>(&m_originalDrawThemeEdge));
    MH_STATUS fillStatus = MH_CreateHook(m_fillRectTarget, reinterpret_cast<LPVOID>(&HookedFillRect),
                                         reinterpret_cast<LPVOID*>(&m_originalFillRect));
    MH_STATUS gradientStatus = MH_CreateHook(m_gradientFillTarget,
                                             reinterpret_cast<LPVOID>(&HookedGdiGradientFill),
                                             reinterpret_cast<LPVOID*>(&m_originalGradientFill));

    if ((backgroundStatus != MH_OK && backgroundStatus != MH_ERROR_ALREADY_CREATED) ||
        (edgeStatus != MH_OK && edgeStatus != MH_ERROR_ALREADY_CREATED) ||
        (fillStatus != MH_OK && fillStatus != MH_ERROR_ALREADY_CREATED) ||
        (gradientStatus != MH_OK && gradientStatus != MH_ERROR_ALREADY_CREATED)) {
        if (initializedLibrary) {
            MH_Uninitialize();
        }
        m_drawThemeBackgroundTarget = nullptr;
        m_drawThemeEdgeTarget = nullptr;
        m_fillRectTarget = nullptr;
        m_gradientFillTarget = nullptr;
        m_originalDrawThemeBackground = nullptr;
        m_originalDrawThemeEdge = nullptr;
        m_originalFillRect = nullptr;
        m_originalGradientFill = nullptr;
        return false;
    }

    m_hooksInitialized = true;
    return true;
}

void ThemeHooks::DisableHooksLocked() {
    if (!m_hooksInitialized || !m_hooksEnabled) {
        return;
    }

    auto disable = [&](void* target) {
        if (!target) {
            return;
        }
        const MH_STATUS status = MH_DisableHook(target);
        if (status != MH_OK && status != MH_ERROR_DISABLED) {
            LogMessage(LogLevel::Warning, L"MH_DisableHook failed for target %p (status=%d)", target,
                       static_cast<int>(status));
        }
    };

    disable(m_drawThemeBackgroundTarget);
    disable(m_drawThemeEdgeTarget);
    disable(m_fillRectTarget);
    disable(m_gradientFillTarget);
    m_hooksEnabled = false;
}

void ThemeHooks::DestroyHooksLocked() {
    if (!m_hooksInitialized) {
        return;
    }

    DisableHooksLocked();

    auto remove = [&](void*& target) {
        if (!target) {
            return;
        }
        const MH_STATUS status = MH_RemoveHook(target);
        if (status != MH_OK && status != MH_ERROR_NOT_CREATED) {
            LogMessage(LogLevel::Warning, L"MH_RemoveHook failed for target %p (status=%d)", target,
                       static_cast<int>(status));
        }
        target = nullptr;
    };

    remove(m_drawThemeBackgroundTarget);
    remove(m_drawThemeEdgeTarget);
    remove(m_fillRectTarget);
    remove(m_gradientFillTarget);

    m_originalDrawThemeBackground = nullptr;
    m_originalDrawThemeEdge = nullptr;
    m_originalFillRect = nullptr;
    m_originalGradientFill = nullptr;

    MH_Uninitialize();
    m_hooksInitialized = false;
    m_hooksEnabled = false;
}

}  // namespace shelltabs

