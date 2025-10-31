#include "GlowRenderer.h"

#include <algorithm>
#include <array>
#include <dwmapi.h>
#include <gdiplus.h>
#include <uxtheme.h>
#include <windowsx.h>

namespace shelltabs {
namespace {

constexpr BYTE kLineAlpha = 220;
constexpr BYTE kHaloAlpha = 96;
constexpr BYTE kFrameAlpha = 210;
constexpr BYTE kFrameHaloAlpha = 110;

Gdiplus::Color ToColor(BYTE alpha, COLORREF value) {
    return Gdiplus::Color(alpha, GetRValue(value), GetGValue(value), GetBValue(value));
}

int ScaleByDpi(int value, UINT dpi) {
    if (dpi == 0) {
        dpi = 96;
    }
    return std::max(1, MulDiv(value, static_cast<int>(dpi), 96));
}

void FillGradientRect(Gdiplus::Graphics& graphics, const ExplorerGlowRenderer::GlowColorSet& colors,
                      const Gdiplus::Rect& rect, BYTE alpha) {
    if (rect.Width <= 0 || rect.Height <= 0) {
        return;
    }

    if (colors.gradient) {
        Gdiplus::LinearGradientBrush brush(Gdiplus::Point(rect.X, rect.Y),
                                           Gdiplus::Point(rect.X, rect.GetBottom()),
                                           ToColor(alpha, colors.start), ToColor(alpha, colors.end));
        graphics.FillRectangle(&brush, rect);
    } else {
        Gdiplus::SolidBrush brush(ToColor(alpha, colors.start));
        graphics.FillRectangle(&brush, rect);
    }
}

void FillFrameRegion(Gdiplus::Graphics& graphics, const ExplorerGlowRenderer::GlowColorSet& colors,
                     const RECT& outerRect, const RECT& innerRect, BYTE alpha) {
    if (outerRect.left >= outerRect.right || outerRect.top >= outerRect.bottom) {
        return;
    }

    const Gdiplus::Rect outer(outerRect.left, outerRect.top, outerRect.right - outerRect.left,
                              outerRect.bottom - outerRect.top);
    const Gdiplus::Rect inner(innerRect.left, innerRect.top, innerRect.right - innerRect.left,
                              innerRect.bottom - innerRect.top);

    if (inner.Width <= 0 || inner.Height <= 0) {
        FillGradientRect(graphics, colors, outer, alpha);
        return;
    }

    if (inner.X < outer.X || inner.Y < outer.Y || inner.GetRight() > outer.GetRight() ||
        inner.GetBottom() > outer.GetBottom()) {
        FillGradientRect(graphics, colors, outer, alpha);
        return;
    }

    const std::array<Gdiplus::Rect, 4> regions = {
        Gdiplus::Rect(outer.X, outer.Y, outer.Width, inner.Y - outer.Y),
        Gdiplus::Rect(outer.X, inner.GetBottom(), outer.Width, outer.GetBottom() - inner.GetBottom()),
        Gdiplus::Rect(outer.X, inner.Y, inner.X - outer.X, inner.Height),
        Gdiplus::Rect(inner.GetRight(), inner.Y, outer.GetRight() - inner.GetRight(), inner.Height)};

    for (const auto& region : regions) {
        if (region.Width > 0 && region.Height > 0) {
            FillGradientRect(graphics, colors, region, alpha);
        }
    }
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

}  // namespace

ExplorerGlowRenderer::ExplorerGlowRenderer() = default;

void ExplorerGlowRenderer::Configure(const ShellTabsOptions& options) {
    m_glowEnabled = options.enableNeonGlow;
    m_palette = options.glowPalette;
    UpdateAccentColor();
}

void ExplorerGlowRenderer::RegisterSurface(HWND hwnd, ExplorerSurfaceKind kind) {
    if (!hwnd || !IsWindow(hwnd)) {
        return;
    }

    SurfaceState& state = m_surfaces[hwnd];
    state.kind = kind;
    state.dpiX = GetDpiForWindow(hwnd);
    state.dpiY = state.dpiX;
}

void ExplorerGlowRenderer::UnregisterSurface(HWND hwnd) {
    m_surfaces.erase(hwnd);
}

void ExplorerGlowRenderer::Reset() {
    m_surfaces.clear();
}

void ExplorerGlowRenderer::InvalidateRegisteredSurfaces() const {
    for (const auto& entry : m_surfaces) {
        HWND hwnd = entry.first;
        if (hwnd && IsWindow(hwnd)) {
            InvalidateRect(hwnd, nullptr, FALSE);
        }
    }
}

void ExplorerGlowRenderer::HandleThemeChanged(HWND hwnd) {
    UpdateAccentColor();
    if (hwnd && IsWindow(hwnd)) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void ExplorerGlowRenderer::HandleSettingChanged(HWND hwnd) {
    UpdateAccentColor();
    if (hwnd && IsWindow(hwnd)) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void ExplorerGlowRenderer::HandleDpiChanged(HWND hwnd, UINT dpiX, UINT dpiY) {
    auto it = m_surfaces.find(hwnd);
    if (it != m_surfaces.end()) {
        it->second.dpiX = dpiX;
        it->second.dpiY = dpiY;
    }
}

void ExplorerGlowRenderer::EnsureSurfaceState(HWND hwnd, ExplorerSurfaceKind kind) {
    if (!hwnd) {
        return;
    }
    auto it = m_surfaces.find(hwnd);
    if (it == m_surfaces.end()) {
        RegisterSurface(hwnd, kind);
    } else {
        it->second.kind = kind;
        if (!IsWindow(hwnd)) {
            return;
        }
        const UINT dpi = GetDpiForWindow(hwnd);
        it->second.dpiX = dpi;
        it->second.dpiY = dpi;
    }
}

ExplorerGlowRenderer::GlowColorSet ExplorerGlowRenderer::ResolveColors(ExplorerSurfaceKind kind) const {
    GlowColorSet colors{};
    if (!m_glowEnabled) {
        return colors;
    }

    const GlowSurfaceOptions* options = nullptr;
    switch (kind) {
        case ExplorerSurfaceKind::ListView:
        case ExplorerSurfaceKind::DirectUi:
            options = &m_palette.listView;
            break;
        case ExplorerSurfaceKind::Header:
            options = &m_palette.header;
            break;
        case ExplorerSurfaceKind::Rebar:
            options = &m_palette.rebar;
            break;
        case ExplorerSurfaceKind::Toolbar:
            options = &m_palette.toolbar;
            break;
        case ExplorerSurfaceKind::Edit:
            options = &m_palette.edits;
            break;
        default:
            break;
    }

    if (!options) {
        return colors;
    }

    colors.valid = true;
    switch (options->mode) {
        case GlowSurfaceMode::kExplorerAccent:
            colors.gradient = false;
            colors.start = m_accentColor;
            colors.end = m_accentColor;
            break;
        case GlowSurfaceMode::kSolid:
            colors.gradient = false;
            colors.start = options->solidColor;
            colors.end = options->solidColor;
            break;
        case GlowSurfaceMode::kGradient:
            colors.gradient = true;
            colors.start = options->gradientStartColor;
            colors.end = options->gradientEndColor;
            break;
        default:
            colors.valid = false;
            break;
    }

    return colors;
}

void ExplorerGlowRenderer::UpdateAccentColor() {
    DWORD color = 0;
    BOOL opaque = FALSE;
    if (SUCCEEDED(DwmGetColorizationColor(&color, &opaque))) {
        m_accentColor = RGB((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
    } else {
        m_accentColor = GetSysColor(COLOR_HOTLIGHT);
    }
}

void ExplorerGlowRenderer::PaintSurface(HWND hwnd, ExplorerSurfaceKind kind, HDC targetDc,
                                        const RECT& clipRect) {
    if (!hwnd || !targetDc) {
        return;
    }

    EnsureSurfaceState(hwnd, kind);
    auto it = m_surfaces.find(hwnd);
    if (it == m_surfaces.end()) {
        return;
    }

    const GlowColorSet colors = ResolveColors(kind);
    if (!colors.valid) {
        return;
    }

    const SurfaceState& state = it->second;

    switch (kind) {
        case ExplorerSurfaceKind::ListView:
            PaintListView(hwnd, state, targetDc, clipRect, colors);
            break;
        case ExplorerSurfaceKind::Header:
            PaintHeader(hwnd, state, targetDc, clipRect, colors);
            break;
        case ExplorerSurfaceKind::Rebar:
            PaintRebar(hwnd, state, targetDc, clipRect, colors);
            break;
        case ExplorerSurfaceKind::Toolbar:
            PaintToolbar(hwnd, state, targetDc, clipRect, colors);
            break;
        case ExplorerSurfaceKind::Edit:
            PaintEdit(hwnd, state, targetDc, clipRect, colors);
            break;
        case ExplorerSurfaceKind::DirectUi:
            PaintDirectUi(hwnd, state, targetDc, clipRect, colors);
            break;
        default:
            break;
    }
}

void ExplorerGlowRenderer::PaintListView(HWND hwnd, const SurfaceState& state, HDC targetDc,
                                         const RECT& clipRect, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, WC_LISTVIEWW)) {
        return;
    }

    const DWORD viewStyle = static_cast<DWORD>(SendMessageW(hwnd, LVM_GETVIEW, 0, 0));
    if (viewStyle != LV_VIEW_DETAILS) {
        return;
    }

    RECT clientRect = GetClientRectSafe(hwnd);
    if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
        return;
    }

    RECT paintRect = clipRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        paintRect = clientRect;
    }

    HDC bufferDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(targetDc, &paintRect, BPBF_TOPDOWNDIB, nullptr, &bufferDc);
    if (!buffer || !bufferDc) {
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        return;
    }

    BufferedPaintClear(buffer, nullptr);
    Gdiplus::Graphics graphics(bufferDc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    const int lineThickness = ScaleByDpi(1, state.dpiY);
    const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, state.dpiY));

    HWND header = ListView_GetHeader(hwnd);
    if (header && IsWindow(header)) {
        const int columnCount = Header_GetItemCount(header);
        for (int index = 0; index < columnCount; ++index) {
            RECT headerRect{};
            if (!Header_GetItemRect(header, index, &headerRect)) {
                continue;
            }
            MapWindowPoints(header, hwnd, reinterpret_cast<POINT*>(&headerRect), 2);
            const int lineLeft = headerRect.right - lineThickness;
            const int haloLeft = lineLeft - (haloThickness - lineThickness) / 2;
            const int height = clientRect.bottom - clientRect.top;
            Gdiplus::Rect haloRect(haloLeft, clientRect.top, haloThickness, height);
            FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
            Gdiplus::Rect lineRect(lineLeft, clientRect.top, lineThickness, height);
            FillGradientRect(graphics, colors, lineRect, kLineAlpha);
        }
    }

    const int topIndex = static_cast<int>(SendMessageW(hwnd, LVM_GETTOPINDEX, 0, 0));
    const int countPerPage = static_cast<int>(SendMessageW(hwnd, LVM_GETCOUNTPERPAGE, 0, 0));
    const int totalCount = ListView_GetItemCount(hwnd);
    const int endIndex = std::min(totalCount, topIndex + countPerPage + 1);

    for (int index = topIndex; index < endIndex; ++index) {
        RECT itemRect{};
        if (!ListView_GetItemRect(hwnd, index, &itemRect, LVIR_BOUNDS)) {
            continue;
        }
        if (itemRect.bottom <= itemRect.top) {
            continue;
        }
        const int y = itemRect.bottom - lineThickness;
        const int haloTop = y - (haloThickness - lineThickness) / 2;
        Gdiplus::Rect haloRect(clientRect.left, haloTop, clientRect.right - clientRect.left, haloThickness);
        FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
        Gdiplus::Rect lineRect(clientRect.left, y, clientRect.right - clientRect.left, lineThickness);
        FillGradientRect(graphics, colors, lineRect, kLineAlpha);
    }

    const bool hasFocus = (GetFocus() == hwnd);
    if (hasFocus) {
        const int focused = ListView_GetNextItem(hwnd, -1, LVNI_FOCUSED);
        if (focused >= 0) {
            RECT focusRect{};
            if (ListView_GetItemRect(hwnd, focused, &focusRect, LVIR_BOUNDS)) {
                RECT inner = focusRect;
                InflateRect(&inner, -ScaleByDpi(1, state.dpiX), -ScaleByDpi(1, state.dpiY));
                RECT frame = inner;
                InflateRect(&frame, lineThickness, lineThickness);
                RECT halo = inner;
                InflateRect(&halo, haloThickness, haloThickness);
                FillFrameRegion(graphics, colors, halo, inner, kFrameHaloAlpha);
                FillFrameRegion(graphics, colors, frame, inner, kFrameAlpha);
            }
        }
    }

    EndBufferedPaint(buffer, TRUE);
}

void ExplorerGlowRenderer::PaintHeader(HWND hwnd, const SurfaceState& state, HDC targetDc,
                                       const RECT& clipRect, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, WC_HEADERW)) {
        return;
    }

    RECT clientRect = GetClientRectSafe(hwnd);
    if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
        return;
    }

    RECT paintRect = clipRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        paintRect = clientRect;
    }

    HDC bufferDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(targetDc, &paintRect, BPBF_TOPDOWNDIB, nullptr, &bufferDc);
    if (!buffer || !bufferDc) {
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        return;
    }

    BufferedPaintClear(buffer, nullptr);
    Gdiplus::Graphics graphics(bufferDc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    const int lineThickness = ScaleByDpi(1, state.dpiY);
    const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, state.dpiY));

    const int width = clientRect.right - clientRect.left;
    const int bottom = clientRect.bottom - lineThickness;
    const int haloBottom = bottom - (haloThickness - lineThickness) / 2;
    Gdiplus::Rect haloRect(clientRect.left, haloBottom, width, haloThickness);
    FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
    Gdiplus::Rect lineRect(clientRect.left, bottom, width, lineThickness);
    FillGradientRect(graphics, colors, lineRect, kLineAlpha);

    const int itemCount = Header_GetItemCount(hwnd);
    for (int index = 0; index < itemCount; ++index) {
        RECT itemRect{};
        if (!Header_GetItemRect(hwnd, index, &itemRect)) {
            continue;
        }
        const int lineLeft = itemRect.right - lineThickness;
        const int haloLeft = lineLeft - (haloThickness - lineThickness) / 2;
        Gdiplus::Rect itemHalo(haloLeft, clientRect.top, haloThickness, clientRect.bottom - clientRect.top);
        FillGradientRect(graphics, colors, itemHalo, kHaloAlpha);
        Gdiplus::Rect itemLine(lineLeft, clientRect.top, lineThickness, clientRect.bottom - clientRect.top);
        FillGradientRect(graphics, colors, itemLine, kLineAlpha);
    }

    EndBufferedPaint(buffer, TRUE);
}

void ExplorerGlowRenderer::PaintRebar(HWND hwnd, const SurfaceState& state, HDC targetDc,
                                      const RECT& clipRect, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, REBARCLASSNAMEW)) {
        return;
    }

    RECT clientRect = GetClientRectSafe(hwnd);
    if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
        return;
    }

    RECT paintRect = clipRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        paintRect = clientRect;
    }

    HDC bufferDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(targetDc, &paintRect, BPBF_TOPDOWNDIB, nullptr, &bufferDc);
    if (!buffer || !bufferDc) {
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        return;
    }

    BufferedPaintClear(buffer, nullptr);
    Gdiplus::Graphics graphics(bufferDc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    const int lineThickness = ScaleByDpi(1, state.dpiY);
    const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, state.dpiY));

    const int bandCount = static_cast<int>(SendMessageW(hwnd, RB_GETBANDCOUNT, 0, 0));
    for (int index = 0; index < bandCount - 1; ++index) {
        RECT bandRect{};
        if (!SendMessageW(hwnd, RB_GETRECT, index, reinterpret_cast<LPARAM>(&bandRect))) {
            continue;
        }
        const int lineLeft = bandRect.right - lineThickness;
        const int haloLeft = lineLeft - (haloThickness - lineThickness) / 2;
        Gdiplus::Rect haloRect(haloLeft, clientRect.top, haloThickness, clientRect.bottom - clientRect.top);
        FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
        Gdiplus::Rect lineRect(lineLeft, clientRect.top, lineThickness, clientRect.bottom - clientRect.top);
        FillGradientRect(graphics, colors, lineRect, kLineAlpha);
    }

    const int width = clientRect.right - clientRect.left;
    const int topHalo = clientRect.top;
    Gdiplus::Rect topRect(clientRect.left, topHalo, width, haloThickness);
    FillGradientRect(graphics, colors, topRect, kHaloAlpha);
    Gdiplus::Rect topLine(clientRect.left, clientRect.top, width, lineThickness);
    FillGradientRect(graphics, colors, topLine, kLineAlpha);

    const int bottomLine = clientRect.bottom - lineThickness;
    const int bottomHalo = bottomLine - (haloThickness - lineThickness) / 2;
    Gdiplus::Rect haloRect(clientRect.left, bottomHalo, width, haloThickness);
    FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
    Gdiplus::Rect lineRect(clientRect.left, bottomLine, width, lineThickness);
    FillGradientRect(graphics, colors, lineRect, kLineAlpha);

    EndBufferedPaint(buffer, TRUE);
}

void ExplorerGlowRenderer::PaintToolbar(HWND hwnd, const SurfaceState& state, HDC targetDc,
                                        const RECT& clipRect, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, TOOLBARCLASSNAMEW)) {
        return;
    }

    RECT clientRect = GetClientRectSafe(hwnd);
    if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
        return;
    }

    RECT paintRect = clipRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        paintRect = clientRect;
    }

    HDC bufferDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(targetDc, &paintRect, BPBF_TOPDOWNDIB, nullptr, &bufferDc);
    if (!buffer || !bufferDc) {
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        return;
    }

    BufferedPaintClear(buffer, nullptr);
    Gdiplus::Graphics graphics(bufferDc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    const int lineThickness = ScaleByDpi(1, state.dpiY);
    const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, state.dpiY));

    const int buttonCount = static_cast<int>(SendMessageW(hwnd, TB_BUTTONCOUNT, 0, 0));
    for (int index = 0; index < buttonCount; ++index) {
        RECT itemRect{};
        if (!SendMessageW(hwnd, TB_GETITEMRECT, index, reinterpret_cast<LPARAM>(&itemRect))) {
            continue;
        }
        const int lineLeft = itemRect.right - lineThickness;
        const int haloLeft = lineLeft - (haloThickness - lineThickness) / 2;
        Gdiplus::Rect haloRect(haloLeft, clientRect.top, haloThickness, clientRect.bottom - clientRect.top);
        FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
        Gdiplus::Rect lineRect(lineLeft, clientRect.top, lineThickness, clientRect.bottom - clientRect.top);
        FillGradientRect(graphics, colors, lineRect, kLineAlpha);
    }

    EndBufferedPaint(buffer, TRUE);
}

void ExplorerGlowRenderer::PaintEdit(HWND hwnd, const SurfaceState& state, HDC targetDc,
                                     const RECT& clipRect, const GlowColorSet& colors) {
    if (!MatchesClass(hwnd, L"Edit")) {
        return;
    }

    RECT clientRect = GetClientRectSafe(hwnd);
    if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
        return;
    }

    RECT paintRect = clipRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        paintRect = clientRect;
    }

    HDC bufferDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(targetDc, &paintRect, BPBF_TOPDOWNDIB, nullptr, &bufferDc);
    if (!buffer || !bufferDc) {
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        return;
    }

    BufferedPaintClear(buffer, nullptr);
    Gdiplus::Graphics graphics(bufferDc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    RECT inner = clientRect;
    InflateRect(&inner, -ScaleByDpi(2, state.dpiX), -ScaleByDpi(2, state.dpiY));
    RECT frame = inner;
    InflateRect(&frame, ScaleByDpi(1, state.dpiX), ScaleByDpi(1, state.dpiY));
    RECT halo = inner;
    InflateRect(&halo, ScaleByDpi(3, state.dpiX), ScaleByDpi(3, state.dpiY));

    FillFrameRegion(graphics, colors, halo, inner, kFrameHaloAlpha);
    FillFrameRegion(graphics, colors, frame, inner, kFrameAlpha);

    EndBufferedPaint(buffer, TRUE);
}

void ExplorerGlowRenderer::PaintDirectUi(HWND hwnd, const SurfaceState& state, HDC targetDc,
                                         const RECT& clipRect, const GlowColorSet& colors) {
    RECT clientRect = GetClientRectSafe(hwnd);
    if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
        return;
    }

    RECT paintRect = clipRect;
    if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
        paintRect = clientRect;
    }

    HDC bufferDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(targetDc, &paintRect, BPBF_TOPDOWNDIB, nullptr, &bufferDc);
    if (!buffer || !bufferDc) {
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        return;
    }

    BufferedPaintClear(buffer, nullptr);
    Gdiplus::Graphics graphics(bufferDc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

    RECT inner = clientRect;
    InflateRect(&inner, -ScaleByDpi(2, state.dpiX), -ScaleByDpi(2, state.dpiY));
    RECT frame = inner;
    InflateRect(&frame, ScaleByDpi(1, state.dpiX), ScaleByDpi(1, state.dpiY));
    RECT halo = inner;
    InflateRect(&halo, ScaleByDpi(3, state.dpiX), ScaleByDpi(3, state.dpiY));

    FillFrameRegion(graphics, colors, halo, inner, kFrameHaloAlpha);
    FillFrameRegion(graphics, colors, frame, inner, kFrameAlpha);

    EndBufferedPaint(buffer, TRUE);
}

}  // namespace shelltabs

