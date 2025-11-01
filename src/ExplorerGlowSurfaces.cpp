#include "ExplorerGlowSurfaces.h"

#include "ExplorerThemeUtils.h"

#include <algorithm>
#include <array>
#include <utility>
#include <vector>

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

void FillGradientRect(Gdiplus::Graphics& graphics, const GlowColorSet& colors,
                      const Gdiplus::Rect& rect, BYTE alpha) {
    if (!colors.valid || rect.Width <= 0 || rect.Height <= 0) {
        return;
    }

    const Gdiplus::Color start(alpha, GetRValue(colors.start), GetGValue(colors.start),
                               GetBValue(colors.start));
    const Gdiplus::Color end(alpha, GetRValue(colors.end), GetGValue(colors.end),
                             GetBValue(colors.end));

    if (!colors.gradient || colors.start == colors.end) {
        Gdiplus::SolidBrush brush(start);
        graphics.FillRectangle(&brush, rect);
        return;
    }

    Gdiplus::LinearGradientBrush brush(rect, start, end, 90.0f);
    graphics.FillRectangle(&brush, rect);
}

void FillFrameRegion(Gdiplus::Graphics& graphics, const GlowColorSet& colors,
                     const RECT& outerRect, const RECT& innerRect, BYTE alpha) {
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

    const Gdiplus::Rect outer(outerRect.left, outerRect.top, width, height);
    FillGradientRect(graphics, colors, outer, alpha);

    const LONG innerWidth = std::max<LONG>(clippedInner.right - clippedInner.left, 0);
    const LONG innerHeight = std::max<LONG>(clippedInner.bottom - clippedInner.top, 0);
    if (innerWidth <= 0 || innerHeight <= 0) {
        return;
    }

    Gdiplus::Region innerRegion(Gdiplus::Rect(clippedInner.left, clippedInner.top, innerWidth,
                                              innerHeight));
    graphics.ExcludeClip(&innerRegion);
    FillGradientRect(graphics, colors, outer, alpha);
    graphics.ResetClip();
}

class ListViewGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    void OnPaint(HDC targetDc, const RECT& clipRect, const GlowColorSet& colors) override {
        HWND hwnd = Handle();
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
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &paintRect, &clientRect)) {
                return;
            }
            paintRect = intersect;
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
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintRect.left),
                                    -static_cast<Gdiplus::REAL>(paintRect.top));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        const int lineThickness = ScaleByDpi(1, DpiY());
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, DpiY()));

        if (HWND header = ListView_GetHeader(hwnd); header && IsWindow(header)) {
            const int columnCount = Header_GetItemCount(header);
            for (int index = 0; index < columnCount; ++index) {
                RECT headerRect{};
                if (!Header_GetItemRect(header, index, &headerRect)) {
                    continue;
                }
                MapWindowPoints(header, hwnd, reinterpret_cast<POINT*>(&headerRect), 2);
                if (headerRect.right <= paintRect.left || headerRect.left >= paintRect.right) {
                    continue;
                }
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
            if (itemRect.top >= paintRect.bottom || itemRect.bottom <= paintRect.top) {
                continue;
            }
            const int y = itemRect.bottom - lineThickness;
            const int haloTop = y - (haloThickness - lineThickness) / 2;
            Gdiplus::Rect haloRect(clientRect.left, haloTop, clientRect.right - clientRect.left,
                                   haloThickness);
            FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
            Gdiplus::Rect lineRect(clientRect.left, y, clientRect.right - clientRect.left, lineThickness);
            FillGradientRect(graphics, colors, lineRect, kLineAlpha);
        }

        if (GetFocus() == hwnd) {
            const int focused = ListView_GetNextItem(hwnd, -1, LVNI_FOCUSED);
            if (focused >= 0) {
                RECT focusRect{};
                if (ListView_GetItemRect(hwnd, focused, &focusRect, LVIR_BOUNDS)) {
                    RECT inner = focusRect;
                    InflateRect(&inner, -ScaleByDpi(1, DpiX()), -ScaleByDpi(1, DpiY()));
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
};

class HeaderGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    bool UsesCustomDraw() const noexcept override { return true; }

    bool HandleNotify(const NMHDR& header, LRESULT* result) override {
        if (!result || header.hwndFrom != Handle()) {
            return false;
        }

        const auto* custom = reinterpret_cast<const NMCUSTOMDRAW*>(&header);
        if (!custom) {
            return false;
        }

        switch (custom->dwDrawStage) {
            case CDDS_PREPAINT:
                *result = CDRF_NOTIFYPOSTPAINT;
                return true;
            case CDDS_POSTPAINT: {
                RECT clip = GetClientRectSafe(header.hwndFrom);
                PaintHeaderGlow(custom->hdc, clip);
                *result = CDRF_DODEFAULT;
                return true;
            }
            default:
                break;
        }
        return false;
    }

    void OnPaint(HDC, const RECT&, const GlowColorSet&) override {}

private:
    void PaintHeaderGlow(HDC targetDc, const RECT& clipRect) {
        if (!Coordinator().ShouldRenderSurface(Kind())) {
            return;
        }
        GlowColorSet colors = Coordinator().ResolveColors(Kind());
        if (!colors.valid) {
            return;
        }

        HWND hwnd = Handle();
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
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &paintRect, &clientRect)) {
                return;
            }
            paintRect = intersect;
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
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintRect.left),
                                    -static_cast<Gdiplus::REAL>(paintRect.top));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        const int lineThickness = ScaleByDpi(1, DpiY());
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, DpiY()));

        const int columnCount = Header_GetItemCount(hwnd);
        for (int index = 0; index < columnCount; ++index) {
            RECT itemRect{};
            if (!Header_GetItemRect(hwnd, index, &itemRect)) {
                continue;
            }
            if (itemRect.right <= paintRect.left || itemRect.left >= paintRect.right) {
                continue;
            }
            const int lineLeft = itemRect.right - lineThickness;
            const int haloLeft = lineLeft - (haloThickness - lineThickness) / 2;
            const int height = clientRect.bottom - clientRect.top;
            Gdiplus::Rect haloRect(haloLeft, clientRect.top, haloThickness, height);
            FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
            Gdiplus::Rect lineRect(lineLeft, clientRect.top, lineThickness, height);
            FillGradientRect(graphics, colors, lineRect, kLineAlpha);
        }

        EndBufferedPaint(buffer, TRUE);
    }
};

class RebarGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    bool UsesCustomDraw() const noexcept override { return true; }

    bool HandleNotify(const NMHDR& header, LRESULT* result) override {
        if (!result || header.hwndFrom != Handle()) {
            return false;
        }

        const auto* custom = reinterpret_cast<const NMCUSTOMDRAW*>(&header);
        if (!custom) {
            return false;
        }

        switch (custom->dwDrawStage) {
            case CDDS_PREPAINT:
                *result = CDRF_NOTIFYPOSTPAINT;
                return true;
            case CDDS_POSTPAINT: {
                RECT clip = GetClientRectSafe(header.hwndFrom);
                PaintRebarGlow(custom->hdc, clip);
                *result = CDRF_DODEFAULT;
                return true;
            }
            default:
                break;
        }

        return false;
    }

    void OnPaint(HDC, const RECT&, const GlowColorSet&) override {}

private:
    void PaintRebarGlow(HDC targetDc, const RECT& clipRect) {
        if (!Coordinator().ShouldRenderSurface(Kind())) {
            return;
        }
        GlowColorSet colors = Coordinator().ResolveColors(Kind());
        if (!colors.valid) {
            return;
        }

        HWND hwnd = Handle();
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
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &paintRect, &clientRect)) {
                return;
            }
            paintRect = intersect;
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
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintRect.left),
                                    -static_cast<Gdiplus::REAL>(paintRect.top));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        const int lineThickness = ScaleByDpi(1, DpiY());
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, DpiY()));

        const UINT bandCount = static_cast<UINT>(SendMessageW(hwnd, RB_GETBANDCOUNT, 0, 0));
        for (UINT index = 0; index < bandCount; ++index) {
            RECT bandRect{};
            if (!SendMessageW(hwnd, RB_GETRECT, index, reinterpret_cast<LPARAM>(&bandRect))) {
                continue;
            }
            if (bandRect.bottom <= bandRect.top || bandRect.right <= bandRect.left) {
                continue;
            }
            if (bandRect.bottom <= paintRect.top || bandRect.top >= paintRect.bottom) {
                continue;
            }
            const int y = bandRect.bottom - lineThickness;
            const int haloTop = y - (haloThickness - lineThickness) / 2;
            Gdiplus::Rect haloRect(clientRect.left, haloTop, clientRect.right - clientRect.left,
                                   haloThickness);
            FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
            Gdiplus::Rect lineRect(clientRect.left, y, clientRect.right - clientRect.left, lineThickness);
            FillGradientRect(graphics, colors, lineRect, kLineAlpha);
        }

        EndBufferedPaint(buffer, TRUE);
    }
};

class ToolbarGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    bool UsesCustomDraw() const noexcept override { return true; }

    bool HandleNotify(const NMHDR& header, LRESULT* result) override {
        if (!result || header.hwndFrom != Handle()) {
            return false;
        }

        const auto* custom = reinterpret_cast<const NMCUSTOMDRAW*>(&header);
        if (!custom) {
            return false;
        }

        switch (custom->dwDrawStage) {
            case CDDS_PREPAINT:
                *result = CDRF_NOTIFYPOSTPAINT;
                return true;
            case CDDS_POSTPAINT: {
                RECT clip = GetClientRectSafe(header.hwndFrom);
                PaintToolbarGlow(custom->hdc, clip);
                *result = CDRF_DODEFAULT;
                return true;
            }
            default:
                break;
        }
        return false;
    }

    void OnPaint(HDC, const RECT&, const GlowColorSet&) override {}

private:
    void PaintToolbarGlow(HDC targetDc, const RECT& clipRect) {
        if (!Coordinator().ShouldRenderSurface(Kind())) {
            return;
        }
        GlowColorSet colors = Coordinator().ResolveColors(Kind());
        if (!colors.valid) {
            return;
        }

        HWND hwnd = Handle();
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
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &paintRect, &clientRect)) {
                return;
            }
            paintRect = intersect;
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
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintRect.left),
                                    -static_cast<Gdiplus::REAL>(paintRect.top));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        const int lineThickness = ScaleByDpi(1, DpiY());
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, DpiY()));

        const int buttonCount = static_cast<int>(SendMessageW(hwnd, TB_BUTTONCOUNT, 0, 0));
        for (int index = 0; index < buttonCount; ++index) {
            RECT buttonRect{};
            if (!SendMessageW(hwnd, TB_GETITEMRECT, index, reinterpret_cast<LPARAM>(&buttonRect))) {
                continue;
            }
            if (buttonRect.bottom <= buttonRect.top || buttonRect.right <= buttonRect.left) {
                continue;
            }
            if (buttonRect.bottom <= paintRect.top || buttonRect.top >= paintRect.bottom) {
                continue;
            }
            const int y = buttonRect.bottom - lineThickness;
            const int haloTop = y - (haloThickness - lineThickness) / 2;
            Gdiplus::Rect haloRect(clientRect.left, haloTop, clientRect.right - clientRect.left,
                                   haloThickness);
            FillGradientRect(graphics, colors, haloRect, kHaloAlpha);
            Gdiplus::Rect lineRect(clientRect.left, y, clientRect.right - clientRect.left, lineThickness);
            FillGradientRect(graphics, colors, lineRect, kLineAlpha);
        }

        EndBufferedPaint(buffer, TRUE);
    }
};

class EditGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    void OnPaint(HDC targetDc, const RECT& clipRect, const GlowColorSet& colors) override {
        HWND hwnd = Handle();
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
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &paintRect, &clientRect)) {
                return;
            }
            paintRect = intersect;
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
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintRect.left),
                                    -static_cast<Gdiplus::REAL>(paintRect.top));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        RECT inner = clientRect;
        InflateRect(&inner, -ScaleByDpi(1, DpiX()), -ScaleByDpi(1, DpiY()));
        RECT frame = inner;
        InflateRect(&frame, ScaleByDpi(1, DpiX()), ScaleByDpi(1, DpiY()));
        RECT halo = inner;
        InflateRect(&halo, ScaleByDpi(3, DpiX()), ScaleByDpi(3, DpiY()));

        FillFrameRegion(graphics, colors, halo, inner, kFrameHaloAlpha);
        FillFrameRegion(graphics, colors, frame, inner, kFrameAlpha);

        EndBufferedPaint(buffer, TRUE);
    }
};

class DirectUiGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    void OnPaint(HDC targetDc, const RECT& clipRect, const GlowColorSet& colors) override {
        HWND hwnd = Handle();
        if (!MatchesClass(hwnd, L"DirectUIHWND")) {
            return;
        }

        RECT clientRect = GetClientRectSafe(hwnd);
        if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
            return;
        }

        RECT paintRect = clipRect;
        if (paintRect.right <= paintRect.left || paintRect.bottom <= paintRect.top) {
            paintRect = clientRect;
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &paintRect, &clientRect)) {
                return;
            }
            paintRect = intersect;
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
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintRect.left),
                                    -static_cast<Gdiplus::REAL>(paintRect.top));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        RECT inner = clientRect;
        InflateRect(&inner, -ScaleByDpi(1, DpiX()), -ScaleByDpi(1, DpiY()));
        RECT frame = inner;
        InflateRect(&frame, ScaleByDpi(1, DpiX()), ScaleByDpi(1, DpiY()));
        RECT halo = inner;
        InflateRect(&halo, ScaleByDpi(3, DpiX()), ScaleByDpi(3, DpiY()));

        FillFrameRegion(graphics, colors, halo, inner, kFrameHaloAlpha);
        FillFrameRegion(graphics, colors, frame, inner, kFrameAlpha);

        EndBufferedPaint(buffer, TRUE);
    }
};

}  // namespace

ExplorerGlowCoordinator::ExplorerGlowCoordinator() = default;

void ExplorerGlowCoordinator::Configure(const ShellTabsOptions& options) {
    m_glowEnabled = options.enableNeonGlow;
    m_palette = options.glowPalette;
    RefreshAccessibilityState();
    UpdateAccentColor();
}

bool ExplorerGlowCoordinator::HandleThemeChanged() {
    const bool accessibilityChanged = RefreshAccessibilityState();
    const COLORREF previousAccent = m_accentColor;
    UpdateAccentColor();
    return accessibilityChanged || (previousAccent != m_accentColor);
}

bool ExplorerGlowCoordinator::HandleSettingChanged() {
    const bool accessibilityChanged = RefreshAccessibilityState();
    const COLORREF previousAccent = m_accentColor;
    UpdateAccentColor();
    return accessibilityChanged || (previousAccent != m_accentColor);
}

bool ExplorerGlowCoordinator::ShouldRenderSurface(ExplorerSurfaceKind kind) const noexcept {
    if (!ShouldRender()) {
        return false;
    }
    const GlowSurfaceOptions* options = ResolveSurfaceOptions(kind);
    return options && options->enabled;
}

GlowColorSet ExplorerGlowCoordinator::ResolveColors(ExplorerSurfaceKind kind) const {
    GlowColorSet colors{};
    if (!ShouldRender()) {
        return colors;
    }

    const GlowSurfaceOptions* options = ResolveSurfaceOptions(kind);
    if (!options || !options->enabled) {
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

const GlowSurfaceOptions* ExplorerGlowCoordinator::ResolveSurfaceOptions(ExplorerSurfaceKind kind) const noexcept {
    switch (kind) {
        case ExplorerSurfaceKind::ListView:
            return &m_palette.listView;
        case ExplorerSurfaceKind::Header:
            return &m_palette.header;
        case ExplorerSurfaceKind::Rebar:
            return &m_palette.rebar;
        case ExplorerSurfaceKind::Toolbar:
            return &m_palette.toolbar;
        case ExplorerSurfaceKind::Edit:
            return &m_palette.edits;
        case ExplorerSurfaceKind::DirectUi:
            return &m_palette.directUi;
        default:
            return nullptr;
    }
}

void ExplorerGlowCoordinator::UpdateAccentColor() {
    DWORD color = 0;
    BOOL opaque = FALSE;
    if (SUCCEEDED(DwmGetColorizationColor(&color, &opaque))) {
        m_accentColor = RGB((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
    } else {
        m_accentColor = GetSysColor(COLOR_HOTLIGHT);
    }
}

bool ExplorerGlowCoordinator::RefreshAccessibilityState() {
    const bool isHighContrast = IsSystemHighContrastActive();
    if (m_highContrastActive != isHighContrast) {
        m_highContrastActive = isHighContrast;
        return true;
    }
    return false;
}

ExplorerGlowSurface::ExplorerGlowSurface(ExplorerSurfaceKind kind, ExplorerGlowCoordinator& coordinator)
    : m_kind(kind), m_coordinator(coordinator) {}

ExplorerGlowSurface::~ExplorerGlowSurface() { Detach(); }

bool ExplorerGlowSurface::Attach(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    if (m_subclassInstalled) {
        Detach();
    }

    const UINT dpi = GetDpiForWindow(hwnd);
    m_dpiX = dpi;
    m_dpiY = dpi;

    if (!SetWindowSubclass(hwnd, &ExplorerGlowSurface::SubclassProc, reinterpret_cast<UINT_PTR>(this),
                           reinterpret_cast<DWORD_PTR>(this))) {
        return false;
    }

    m_hwnd = hwnd;
    m_subclassInstalled = true;
    OnAttached();
    return true;
}

void ExplorerGlowSurface::Detach() {
    if (!m_subclassInstalled) {
        m_hwnd = nullptr;
        return;
    }

    HWND hwnd = m_hwnd;
    m_hwnd = nullptr;
    m_subclassInstalled = false;
    if (hwnd && IsWindow(hwnd)) {
        RemoveWindowSubclass(hwnd, &ExplorerGlowSurface::SubclassProc, reinterpret_cast<UINT_PTR>(this));
    }
    OnDetached();
}

bool ExplorerGlowSurface::IsAttached() const noexcept {
    return m_hwnd && IsWindow(m_hwnd) && m_subclassInstalled;
}

void ExplorerGlowSurface::RequestRepaint() const {
    if (m_hwnd && IsWindow(m_hwnd)) {
        InvalidateRect(m_hwnd, nullptr, FALSE);
    }
}

bool ExplorerGlowSurface::SupportsImmediatePainting() const noexcept {
    return !UsesCustomDraw();
}

bool ExplorerGlowSurface::PaintImmediately(HDC targetDc, const RECT& clipRect) {
    if (!SupportsImmediatePainting() || !targetDc) {
        return false;
    }
    if (!IsAttached()) {
        return false;
    }
    if (!Coordinator().ShouldRenderSurface(Kind())) {
        return false;
    }

    RECT bounds = clipRect;
    if (bounds.right <= bounds.left || bounds.bottom <= bounds.top) {
        return false;
    }

    PaintInternal(targetDc, bounds);
    return true;
}

bool ExplorerGlowSurface::HandleNotify(const NMHDR&, LRESULT*) { return false; }

void ExplorerGlowSurface::OnAttached() {}

void ExplorerGlowSurface::OnDetached() {}

void ExplorerGlowSurface::OnDpiChanged(UINT, UINT) {}

void ExplorerGlowSurface::OnThemeChanged() {}

void ExplorerGlowSurface::OnSettingsChanged() {}

bool ExplorerGlowSurface::UsesCustomDraw() const noexcept { return false; }

std::optional<LRESULT> ExplorerGlowSurface::HandleMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_NCDESTROY:
            Detach();
            break;
        case WM_DPICHANGED: {
            const UINT dpiX = LOWORD(wParam);
            const UINT dpiY = HIWORD(wParam);
            m_dpiX = dpiX;
            m_dpiY = dpiY;
            OnDpiChanged(dpiX, dpiY);
            break;
        }
        case WM_THEMECHANGED: {
            if (Coordinator().HandleThemeChanged()) {
                RequestRepaint();
            } else if (Coordinator().ShouldRenderSurface(Kind())) {
                RequestRepaint();
            }
            OnThemeChanged();
            break;
        }
        case WM_SETTINGCHANGE: {
            if (Coordinator().HandleSettingChanged()) {
                RequestRepaint();
            } else if (Coordinator().ShouldRenderSurface(Kind())) {
                RequestRepaint();
            }
            OnSettingsChanged();
            break;
        }
        case WM_PAINT:
            if (!UsesCustomDraw()) {
                return HandlePaintMessage(hwnd, msg, wParam, lParam);
            }
            break;
        case WM_PRINTCLIENT:
            if (!UsesCustomDraw()) {
                return HandlePrintClient(hwnd, wParam, lParam);
            }
            break;
        default:
            break;
    }
    return std::nullopt;
}

LRESULT ExplorerGlowSurface::HandlePaintMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    RECT updateRect{0, 0, 0, 0};
    bool hasUpdateRect = false;
    if (wParam == 0) {
        if (GetUpdateRect(hwnd, &updateRect, FALSE)) {
            hasUpdateRect = !IsRectEmpty(&updateRect);
        }
    }

    LRESULT defResult = DefSubclassProc(hwnd, msg, wParam, lParam);

    if (!Coordinator().ShouldRenderSurface(Kind())) {
        return defResult;
    }

    HDC targetDc = reinterpret_cast<HDC>(wParam);
    bool releaseDc = false;
    RECT clipRect{0, 0, 0, 0};
    bool hasClip = false;

    if (!targetDc) {
        UINT flags = DCX_CACHE | DCX_CLIPSIBLINGS | DCX_CLIPCHILDREN | DCX_WINDOW;
        targetDc = GetDCEx(hwnd, nullptr, flags);
        releaseDc = (targetDc != nullptr);
    }

    if (!targetDc) {
        return defResult;
    }

    RECT clip{};
    if (GetClipBox(targetDc, &clip) != ERROR && !IsRectEmpty(&clip)) {
        clipRect = clip;
        hasClip = true;
    }

    if (!hasClip && hasUpdateRect) {
        clipRect = updateRect;
        hasClip = true;
    }

    if (!hasClip) {
        clipRect = GetClientRectSafe(hwnd);
        hasClip = !IsRectEmpty(&clipRect);
    }

    if (hasClip) {
        PaintInternal(targetDc, clipRect);
    }

    if (releaseDc) {
        ReleaseDC(hwnd, targetDc);
    }

    return defResult;
}

LRESULT ExplorerGlowSurface::HandlePrintClient(HWND hwnd, WPARAM wParam, LPARAM lParam) {
    LRESULT defResult = DefSubclassProc(hwnd, WM_PRINTCLIENT, wParam, lParam);
    if (!Coordinator().ShouldRenderSurface(Kind())) {
        return defResult;
    }

    HDC targetDc = reinterpret_cast<HDC>(wParam);
    if (!targetDc) {
        return defResult;
    }

    RECT clip{};
    if (GetClipBox(targetDc, &clip) == ERROR || IsRectEmpty(&clip)) {
        clip = GetClientRectSafe(hwnd);
    }

    if (!IsRectEmpty(&clip)) {
        PaintInternal(targetDc, clip);
    }

    return defResult;
}

void ExplorerGlowSurface::PaintInternal(HDC targetDc, const RECT& clipRect) {
    GlowColorSet colors = Coordinator().ResolveColors(Kind());
    if (!colors.valid) {
        return;
    }
    OnPaint(targetDc, clipRect, colors);
}

LRESULT CALLBACK ExplorerGlowSurface::SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR subclassId, DWORD_PTR refData) {
    UNREFERENCED_PARAMETER(subclassId);
    auto* self = reinterpret_cast<ExplorerGlowSurface*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    std::optional<LRESULT> handled = self->HandleMessage(hwnd, msg, wParam, lParam);
    if (handled.has_value()) {
        return handled.value();
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

std::unique_ptr<ExplorerGlowSurface> CreateGlowSurfaceWrapper(ExplorerSurfaceKind kind,
                                                              ExplorerGlowCoordinator& coordinator) {
    switch (kind) {
        case ExplorerSurfaceKind::ListView:
            return std::make_unique<ListViewGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::Header:
            return std::make_unique<HeaderGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::Rebar:
            return std::make_unique<RebarGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::Toolbar:
            return std::make_unique<ToolbarGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::Edit:
            return std::make_unique<EditGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::DirectUi:
            return std::make_unique<DirectUiGlowSurface>(kind, coordinator);
        default:
            return nullptr;
    }
}

}  // namespace shelltabs

