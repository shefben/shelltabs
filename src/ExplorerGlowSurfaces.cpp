#include "ExplorerGlowSurfaces.h"

#include "BreadcrumbGradient.h"
#include "EditGradientRenderer.h"

#include "ExplorerThemeUtils.h"
#include "ShellTabsListView.h"

#include <algorithm>
#include <array>
#include <cstdlib>
#include <mutex>
#include <utility>
#include <vector>
#include <string>
#include <cwchar>
#include <cmath>

#include <dwmapi.h>
#include <gdiplus.h>
#include <UIAutomation.h>
#include <oleacc.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <windowsx.h>
#include <wrl/client.h>

namespace shelltabs {
namespace {

constexpr BYTE kLineAlpha = 220;
constexpr BYTE kHaloAlpha = 96;
constexpr BYTE kFrameAlpha = 210;
constexpr BYTE kFrameHaloAlpha = 110;

bool ApproximatelyEqual(int lhs, int rhs, int tolerance) {
    return std::abs(lhs - rhs) <= tolerance;
}

Microsoft::WRL::ComPtr<IUIAutomation> GetAutomationInstance() {
    static std::once_flag automationInitFlag;
    static Microsoft::WRL::ComPtr<IUIAutomation> automation;
    static HRESULT automationInitResult = E_FAIL;

    std::call_once(automationInitFlag, []() {
        Microsoft::WRL::ComPtr<IUIAutomation> instance;
        const HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER,
                                            IID_PPV_ARGS(&instance));
        if (SUCCEEDED(hr)) {
            automation = std::move(instance);
            automationInitResult = S_OK;
        } else {
            automation.Reset();
            automationInitResult = hr;
        }
    });

    if (FAILED(automationInitResult)) {
        return nullptr;
    }

    return automation;
}

void CollectDirectUiDescendants(IUIAutomationElement* element, IUIAutomationTreeWalker* walker, HWND host,
                                const RECT& clientRect, std::vector<RECT>& rectangles);

bool MatchesClass(HWND hwnd, const wchar_t* className);

void AppendDirectUiRectangle(IUIAutomationElement* element, HWND host, const RECT& clientRect,
                             std::vector<RECT>& rectangles) {
    if (!element || !host || !IsWindow(host)) {
        return;
    }

    RECT bounding{};
    if (FAILED(element->get_CurrentBoundingRectangle(&bounding))) {
        return;
    }

    if (bounding.right <= bounding.left || bounding.bottom <= bounding.top) {
        return;
    }

    BOOL offscreen = FALSE;
    const HRESULT offscreenHr = element->get_CurrentIsOffscreen(&offscreen);
    if (FAILED(offscreenHr) || offscreen) {
        return;
    }

    POINT points[2] = {{bounding.left, bounding.top}, {bounding.right, bounding.bottom}};
    MapWindowPoints(nullptr, host, points, 2);

    RECT local = {points[0].x, points[0].y, points[1].x, points[1].y};
    RECT clipped{};
    if (!IntersectRect(&clipped, &local, &clientRect)) {
        return;
    }

    if (clipped.right <= clipped.left || clipped.bottom <= clipped.top) {
        return;
    }

    const auto duplicate = std::find_if(rectangles.begin(), rectangles.end(), [&](const RECT& existing) {
        return EqualRect(&existing, &clipped);
    });

    if (duplicate == rectangles.end()) {
        rectangles.push_back(clipped);
    }
}

void CollectDirectUiDescendants(IUIAutomationElement* element, IUIAutomationTreeWalker* walker, HWND host,
                                const RECT& clientRect, std::vector<RECT>& rectangles) {
    if (!element || !walker) {
        return;
    }

    Microsoft::WRL::ComPtr<IUIAutomationElement> child;
    HRESULT hr = walker->GetFirstChildElement(element, &child);
    if (FAILED(hr)) {
        return;
    }

    while (child) {
        AppendDirectUiRectangle(child.Get(), host, clientRect, rectangles);
        CollectDirectUiDescendants(child.Get(), walker, host, clientRect, rectangles);

        Microsoft::WRL::ComPtr<IUIAutomationElement> next;
        hr = walker->GetNextSiblingElement(child.Get(), &next);
        if (FAILED(hr)) {
            break;
        }
        child = std::move(next);
    }
}

bool EnumerateDirectUiRectangles(HWND host, const RECT& clientRect, std::vector<RECT>& rectangles) {
    rectangles.clear();

    auto automation = GetAutomationInstance();
    if (!automation) {
        return false;
    }

    Microsoft::WRL::ComPtr<IUIAutomationElement> root;
    if (FAILED(automation->ElementFromHandle(host, &root)) || !root) {
        return false;
    }

    Microsoft::WRL::ComPtr<IUIAutomationTreeWalker> walker;
    if (FAILED(automation->get_ControlViewWalker(&walker)) || !walker) {
        return false;
    }

    CollectDirectUiDescendants(root.Get(), walker.Get(), host, clientRect, rectangles);
    return true;
}

RECT MapScreenRectToWindow(HWND hwnd, const RECT& rect) {
    POINT points[2] = {{rect.left, rect.top}, {rect.right, rect.bottom}};
    MapWindowPoints(nullptr, hwnd, points, 2);
    return {points[0].x, points[0].y, points[1].x, points[1].y};
}

bool GetHeaderItem(HWND header, int index, HDITEMW* item) {
    if (!header || !item) {
        return false;
    }
    const LRESULT result = SendMessageW(header, HDM_GETITEMW, static_cast<WPARAM>(index),
                                        reinterpret_cast<LPARAM>(item));
    return result != FALSE;
}

HWND FindEditHostWindow(HWND edit) {
    if (!edit || !IsWindow(edit)) {
        return nullptr;
    }

    HWND current = GetParent(edit);
    while (current && IsWindow(current)) {
        if (MatchesClass(current, L"ReBarWindow32") || MatchesClass(current, L"DirectUIHWND")) {
            return current;
        }
        current = GetParent(current);
    }

    return GetParent(edit);
}

struct EditSiblingContext {
    HWND reference = nullptr;
    RECT referenceRect{};
    int verticalTolerance = 0;
    LONG minimumOverlap = 0;
    std::vector<RECT>* results = nullptr;
};

struct ToolbarHitTestInfo {
    HWND hwnd = nullptr;
    POINT pt{};
    UINT flags = 0;
};

BOOL CALLBACK EnumSiblingEditProc(HWND child, LPARAM param) {
    auto* context = reinterpret_cast<EditSiblingContext*>(param);
    if (!context || !context->results) {
        return TRUE;
    }

    if (!MatchesClass(child, L"Edit") || !IsWindowVisible(child)) {
        return TRUE;
    }

    RECT rect{};
    if (!GetWindowRect(child, &rect)) {
        return TRUE;
    }

    if (rect.right <= rect.left || rect.bottom <= rect.top) {
        return TRUE;
    }

    const LONG topDiff = std::abs(rect.top - context->referenceRect.top);
    const LONG bottomDiff = std::abs(rect.bottom - context->referenceRect.bottom);
    const LONG overlapTop = std::max(rect.top, context->referenceRect.top);
    const LONG overlapBottom = std::min(rect.bottom, context->referenceRect.bottom);
    const LONG overlap = overlapBottom - overlapTop;

    if (topDiff > context->verticalTolerance || bottomDiff > context->verticalTolerance) {
        if (overlap < context->minimumOverlap) {
            return TRUE;
        }
    }

    const auto duplicate = std::find_if(context->results->begin(), context->results->end(),
                                        [&](const RECT& existing) { return EqualRect(&existing, &rect); });
    if (duplicate != context->results->end()) {
        return TRUE;
    }

    context->results->push_back(rect);
    return TRUE;
}

void CollectSiblingEditRects(HWND editHwnd, const RECT& referenceRect, int verticalTolerance,
                             std::vector<RECT>& rects) {
    rects.clear();

    HWND host = FindEditHostWindow(editHwnd);
    if (!host || !IsWindow(host)) {
        rects.push_back(referenceRect);
        return;
    }

    EditSiblingContext context;
    context.reference = editHwnd;
    context.referenceRect = referenceRect;
    context.verticalTolerance = verticalTolerance;
    const LONG referenceHeight = referenceRect.bottom - referenceRect.top;
    context.minimumOverlap = std::max<LONG>(referenceHeight / 2, 1);
    context.results = &rects;

    EnumChildWindows(host, EnumSiblingEditProc, reinterpret_cast<LPARAM>(&context));

    if (rects.empty()) {
        rects.push_back(referenceRect);
        return;
    }

    const auto selfIt = std::find_if(rects.begin(), rects.end(), [&](const RECT& rect) {
        return EqualRect(&rect, &referenceRect);
    });

    if (selfIt == rects.end()) {
        rects.push_back(referenceRect);
    }
}

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
                      const Gdiplus::Rect& rect, BYTE alpha, float angle = 90.0f) {
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

    Gdiplus::LinearGradientBrush brush(rect, start, end, angle);
    graphics.FillRectangle(&brush, rect);
}

Gdiplus::Rect RectToGdiplus(const RECT& rect) {
    return {rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top};
}

void FillFrameRegion(Gdiplus::Graphics& graphics, const GlowColorSet& colors,
                     const RECT& outerRect, const RECT& innerRect, BYTE alpha,
                     float angle = 90.0f) {
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
    FillGradientRect(graphics, colors, outer, alpha, angle);

    const LONG innerWidth = std::max<LONG>(clippedInner.right - clippedInner.left, 0);
    const LONG innerHeight = std::max<LONG>(clippedInner.bottom - clippedInner.top, 0);
    if (innerWidth <= 0 || innerHeight <= 0) {
        return;
    }

    Gdiplus::Region innerRegion(Gdiplus::Rect(clippedInner.left, clippedInner.top, innerWidth,
                                              innerHeight));
    graphics.ExcludeClip(&innerRegion);
    FillGradientRect(graphics, colors, outer, alpha, angle);
    graphics.ResetClip();
}

class ListViewGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    void OnPaint(HDC targetDc, const RECT& clipRect, const GlowColorSet& colors) override {
        HWND hwnd = Handle();
        HWND listView = hwnd;
        if (!MatchesClass(listView, WC_LISTVIEWW)) {
            HWND child = FindWindowExW(hwnd, nullptr, WC_LISTVIEWW, nullptr);
            if (!child || !MatchesClass(child, WC_LISTVIEWW)) {
                return;
            }
            listView = child;
        }

        const DWORD viewStyle = static_cast<DWORD>(SendMessageW(listView, LVM_GETVIEW, 0, 0));
        if (viewStyle != LV_VIEW_DETAILS) {
            return;
        }

        RECT clientRect = GetClientRectSafe(listView);
        if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
            return;
        }

        RECT listViewClip = clipRect;
        if (listView != hwnd) {
            MapWindowPoints(hwnd, listView, reinterpret_cast<POINT*>(&listViewClip), 2);
        }

        RECT listViewPaint = listViewClip;
        if (listViewPaint.right <= listViewPaint.left || listViewPaint.bottom <= listViewPaint.top) {
            listViewPaint = clientRect;
        } else {
            RECT intersect{};
            if (!IntersectRect(&intersect, &listViewPaint, &clientRect)) {
                return;
            }
            listViewPaint = intersect;
        }

        RECT paintRect = listViewPaint;
        if (listView != hwnd) {
            MapWindowPoints(listView, hwnd, reinterpret_cast<POINT*>(&paintRect), 2);
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
        POINT paintOrigin = {paintRect.left, paintRect.top};
        if (listView != hwnd) {
            MapWindowPoints(hwnd, listView, &paintOrigin, 1);
        }
        graphics.TranslateTransform(-static_cast<Gdiplus::REAL>(paintOrigin.x),
                                    -static_cast<Gdiplus::REAL>(paintOrigin.y));
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        const int lineThickness = ScaleByDpi(1, DpiY());
        const int haloThickness = std::max(lineThickness * 2, ScaleByDpi(3, DpiY()));

        if (HWND header = ListView_GetHeader(listView); header && IsWindow(header)) {
            const int columnCount = Header_GetItemCount(header);
            for (int index = 0; index < columnCount; ++index) {
                RECT headerRect{};
                if (!Header_GetItemRect(header, index, &headerRect)) {
                    continue;
                }
                MapWindowPoints(header, listView, reinterpret_cast<POINT*>(&headerRect), 2);
                if (headerRect.right <= listViewPaint.left || headerRect.left >= listViewPaint.right) {
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

        const int topIndex = static_cast<int>(SendMessageW(listView, LVM_GETTOPINDEX, 0, 0));
        const int countPerPage = static_cast<int>(SendMessageW(listView, LVM_GETCOUNTPERPAGE, 0, 0));
        const int totalCount = ListView_GetItemCount(listView);
        const int endIndex = std::min(totalCount, topIndex + countPerPage + 1);

        for (int index = topIndex; index < endIndex; ++index) {
            RECT itemRect{};
            if (!ListView_GetItemRect(listView, index, &itemRect, LVIR_BOUNDS)) {
                continue;
            }
            if (itemRect.bottom <= itemRect.top) {
                continue;
            }
            if (itemRect.top >= listViewPaint.bottom || itemRect.bottom <= listViewPaint.top) {
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
            int focused = -1;
            if (ShellTabsListView* control = ShellTabsListView::FromListView(hwnd)) {
                ShellTabsListView::SelectionItem item;
                if (control->TryGetFocusedItem(&item)) {
                    focused = item.index;
                }
            } else {
                focused = ListView_GetNextItem(hwnd, -1, LVNI_FOCUSED);
            }

            if (focused >= 0) {
                RECT focusRect{};
                if (ListView_GetItemRect(listView, focused, &focusRect, LVIR_BOUNDS)) {
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
                *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT;
                return true;
            case CDDS_ITEMPREPAINT:
                if (HandleItemPrepaint(*custom, result)) {
                    return true;
                }
                break;
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
    bool HandleItemPrepaint(const NMCUSTOMDRAW& custom, LRESULT* result) const {
        if (!result) {
            return false;
        }

        if (IsDividerSlot(custom)) {
            *result = CDRF_SKIPDEFAULT | CDRF_NOTIFYPOSTPAINT;
            return true;
        }

        const auto& gradientConfig = Coordinator().BreadcrumbFontGradient();
        if (!gradientConfig.enabled) {
            return false;
        }

        if (!DrawGradientHeaderItem(custom, gradientConfig)) {
            return false;
        }

        *result = CDRF_SKIPDEFAULT;
        return true;
    }

    bool IsDividerSlot(const NMCUSTOMDRAW& custom) const {
        HWND hwnd = Handle();
        if (!hwnd || !IsWindow(hwnd)) {
            return false;
        }

        const RECT& rect = custom.rc;
        if (rect.right <= rect.left || rect.bottom <= rect.top) {
            return false;
        }

        HDHITTESTINFO hit{};
        hit.pt.x = rect.left + (rect.right - rect.left) / 2;
        hit.pt.y = rect.top + (rect.bottom - rect.top) / 2;

        const LRESULT hitIndex = SendMessageW(hwnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hit));
        if (hitIndex < 0) {
            return false;
        }

        if ((hit.flags & (HHT_ONDIVIDER | HHT_ONDIVOPEN)) == 0) {
            return false;
        }

        return true;
    }

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

    bool DrawGradientHeaderItem(const NMCUSTOMDRAW& custom,
                                const BreadcrumbGradientConfig& gradientConfig) const {
        if (custom.dwItemSpec < 0) {
            return false;
        }

        HWND header = Handle();
        if (!header || !IsWindow(header)) {
            return false;
        }

        HWND parent = GetParent(header);
        if (!MatchesClass(parent, WC_LISTVIEWW)) {
            return false;
        }

        const DWORD viewStyle = static_cast<DWORD>(SendMessageW(parent, LVM_GETVIEW, 0, 0));
        if (viewStyle != LV_VIEW_DETAILS) {
            return false;
        }

        const int columnIndex = static_cast<int>(custom.dwItemSpec);

        std::wstring text(256, L'\0');
        HDITEMW item{};
        item.mask = HDI_FORMAT | HDI_TEXT;
        item.pszText = text.data();
        item.cchTextMax = static_cast<int>(text.size());
        if (!GetHeaderItem(header, columnIndex, &item)) {
            return false;
        }

        if (item.cchTextMax > 0 && item.pszText) {
            size_t length = std::wcslen(text.c_str());
            while (length + 1 >= static_cast<size_t>(item.cchTextMax)) {
                text.assign(text.size() * 2, L'\0');
                item.pszText = text.data();
                item.cchTextMax = static_cast<int>(text.size());
                if (!GetHeaderItem(header, columnIndex, &item)) {
                    break;
                }
                length = std::wcslen(text.c_str());
            }
            text.resize(length);
        } else {
            text.clear();
        }

        if (text.empty()) {
            DrawHeaderBackground(custom, item);
            DrawSortArrowIfNeeded(custom, item);
            return true;
        }

        if ((item.fmt & (HDF_OWNERDRAW | HDF_BITMAP | HDF_BITMAP_ON_RIGHT)) != 0) {
            return false;
        }

        const HFONT font = reinterpret_cast<HFONT>(SendMessageW(header, WM_GETFONT, 0, 0));
        HFONT oldFont = nullptr;
        if (font) {
            oldFont = static_cast<HFONT>(SelectObject(custom.hdc, font));
        }

        DrawHeaderBackground(custom, item);

        const BreadcrumbGradientPalette palette = ResolveBreadcrumbGradientPalette(gradientConfig);

        RECT textRect = custom.rc;
        const int horizontalPadding = ScaleByDpi(8, DpiX());
        textRect.left += horizontalPadding;
        textRect.right -= horizontalPadding;

        const bool hasSortArrow = (item.fmt & (HDF_SORTDOWN | HDF_SORTUP)) != 0;
        RECT arrowRect = {};
        if (hasSortArrow) {
            arrowRect = textRect;
            const int arrowWidth = ScaleByDpi(12, DpiX());
            arrowRect.left = std::max<LONG>(textRect.left, textRect.right - arrowWidth);
            arrowRect.right = textRect.right;
            textRect.right = std::max<LONG>(textRect.left, arrowRect.left - ScaleByDpi(4, DpiX()));
        }

        if (textRect.right <= textRect.left) {
            if (oldFont) {
                SelectObject(custom.hdc, oldFont);
            }
            DrawSortArrowIfNeeded(custom, item, &arrowRect);
            return true;
        }

        TEXTMETRICW metrics{};
        GetTextMetricsW(custom.hdc, &metrics);
        const int textHeight = metrics.tmHeight > 0 ? metrics.tmHeight : (textRect.bottom - textRect.top);
        const LONG availableHeight = textRect.bottom - textRect.top;
        const LONG calculatedOffset = (availableHeight - static_cast<LONG>(textHeight)) / 2;
        const int verticalOffset = static_cast<int>(std::max<LONG>(0, calculatedOffset));
        const int textBaseline = textRect.top + verticalOffset;

        std::vector<double> characterWidths(text.size(), 0.0);
        for (size_t i = 0; i < text.size(); ++i) {
            SIZE extent{};
            if (GetTextExtentPoint32W(custom.hdc, &text[i], 1, &extent) && extent.cx > 0) {
                characterWidths[i] = static_cast<double>(extent.cx);
            } else {
                characterWidths[i] = 1.0;
            }
        }

        double totalWidth = 0.0;
        for (double width : characterWidths) {
            totalWidth += width;
        }

        int textStartX = textRect.left;
        if ((item.fmt & HDF_CENTER) != 0) {
            const int available = textRect.right - textRect.left;
            const int offset = static_cast<int>(std::max(0.0, (static_cast<double>(available) - totalWidth) / 2.0));
            textStartX += offset;
        } else if ((item.fmt & HDF_RIGHT) != 0) {
            textStartX = static_cast<int>(std::max<LONG>(textRect.left,
                                                        textRect.right - static_cast<LONG>(std::ceil(totalWidth))));
        }

        struct CharacterMetrics {
            size_t index = 0;
            double x = 0.0;
            double width = 0.0;
        };

        std::vector<CharacterMetrics> characters;
        characters.reserve(text.size());

        double currentX = static_cast<double>(textStartX);
        for (size_t i = 0; i < text.size(); ++i) {
            const double width = std::max(1.0, characterWidths[i]);
            characters.push_back(CharacterMetrics{i, currentX, width});
            currentX += width;
        }

        if (characters.empty()) {
            if (oldFont) {
                SelectObject(custom.hdc, oldFont);
            }
            DrawSortArrowIfNeeded(custom, item, &arrowRect);
            return true;
        }

        double gradientLeft = characters.front().x - characters.front().width * 0.5;
        double gradientRight = characters.back().x + characters.back().width * 0.5;
        gradientLeft = std::clamp(gradientLeft, static_cast<double>(textRect.left),
                                  static_cast<double>(textRect.right));
        gradientRight = std::clamp(gradientRight, static_cast<double>(textRect.left),
                                   static_cast<double>(textRect.right));
        if (gradientRight <= gradientLeft) {
            gradientRight = static_cast<double>(textRect.right);
            gradientLeft = static_cast<double>(textRect.left);
        }
        const double gradientWidth = std::max(1.0, gradientRight - gradientLeft);

        const int previousBkMode = SetBkMode(custom.hdc, TRANSPARENT);
        const UINT previousAlign = GetTextAlign(custom.hdc);
        SetTextAlign(custom.hdc, TA_LEFT | TA_TOP);

        RECT clipRect = textRect;

        for (const CharacterMetrics& character : characters) {
            const double centerX = character.x + character.width / 2.0;
            const double position = (centerX - gradientLeft) / gradientWidth;
            const COLORREF color = EvaluateBreadcrumbGradientColor(palette, position);
            SetTextColor(custom.hdc, color);

            const int drawX = static_cast<int>(std::lround(character.x));
            ExtTextOutW(custom.hdc, drawX, textBaseline, ETO_CLIPPED, &clipRect,
                        text.c_str() + character.index, 1, nullptr);
        }

        SetTextAlign(custom.hdc, previousAlign);
        SetBkMode(custom.hdc, previousBkMode);

        if (oldFont) {
            SelectObject(custom.hdc, oldFont);
        }

        DrawSortArrowIfNeeded(custom, item, &arrowRect);
        return true;
    }

    void DrawHeaderBackground(const NMCUSTOMDRAW& custom, const HDITEMW& item) const {
        RECT rect = custom.rc;
        HTHEME theme = OpenThemeData(Handle(), L"HEADER");
        if (theme) {
            const int state = ResolveThemeState(custom, item);
            DrawThemeBackground(theme, custom.hdc, HP_HEADERITEM, state, &rect, nullptr);
            CloseThemeData(theme);
            return;
        }

        FillRect(custom.hdc, &rect, GetSysColorBrush(COLOR_BTNFACE));
        DrawEdge(custom.hdc, &rect, BDR_RAISEDOUTER, BF_RECT);
    }

    int ResolveThemeState(const NMCUSTOMDRAW& custom, const HDITEMW& item) const {
        const bool hot = (custom.uItemState & CDIS_HOT) != 0;
        const bool pressed = (custom.uItemState & CDIS_SELECTED) != 0;
        const bool sorted = (item.fmt & (HDF_SORTDOWN | HDF_SORTUP)) != 0;

        if (sorted) {
            if (pressed) {
                return HIS_SORTEDPRESSED;
            }
            if (hot) {
                return HIS_SORTEDHOT;
            }
            return HIS_SORTEDNORMAL;
        }

        if (pressed) {
            return HIS_PRESSED;
        }
        if (hot) {
            return HIS_HOT;
        }
        return HIS_NORMAL;
    }

    void DrawSortArrowIfNeeded(const NMCUSTOMDRAW& custom, const HDITEMW& item,
                               const RECT* overrideRect = nullptr) const {
        if ((item.fmt & (HDF_SORTDOWN | HDF_SORTUP)) == 0) {
            return;
        }

        RECT arrowRect = overrideRect ? *overrideRect : custom.rc;
        const int padding = ScaleByDpi(8, DpiX());
        arrowRect.right = custom.rc.right - padding;
        arrowRect.left = std::max<LONG>(custom.rc.left, arrowRect.right - ScaleByDpi(12, DpiX()));
        arrowRect.top = custom.rc.top + padding / 2;
        arrowRect.bottom = custom.rc.bottom - padding / 2;
        if (arrowRect.right <= arrowRect.left || arrowRect.bottom <= arrowRect.top) {
            return;
        }

        const bool ascending = (item.fmt & HDF_SORTUP) != 0;
        const int centerX = arrowRect.left + (arrowRect.right - arrowRect.left) / 2;
        const int centerY = arrowRect.top + (arrowRect.bottom - arrowRect.top) / 2;
        const int halfWidth = std::max<int>(1, static_cast<int>((arrowRect.right - arrowRect.left) / 2));
        const int height = std::max<int>(2, static_cast<int>((arrowRect.bottom - arrowRect.top) / 2));

        POINT points[3];
        if (ascending) {
            points[0] = {centerX, centerY - height};
            points[1] = {centerX - halfWidth, centerY + height};
            points[2] = {centerX + halfWidth, centerY + height};
        } else {
            points[0] = {centerX - halfWidth, centerY - height};
            points[1] = {centerX + halfWidth, centerY - height};
            points[2] = {centerX, centerY + height};
        }

        const COLORREF color = GetSysColor(COLOR_BTNTEXT);
        HPEN pen = CreatePen(PS_SOLID, 1, color);
        HBRUSH brush = CreateSolidBrush(color);
        HPEN oldPen = nullptr;
        HBRUSH oldBrush = nullptr;

        if (pen) {
            oldPen = static_cast<HPEN>(SelectObject(custom.hdc, pen));
        }
        if (brush) {
            oldBrush = static_cast<HBRUSH>(SelectObject(custom.hdc, brush));
        }

        Polygon(custom.hdc, points, 3);

        if (oldBrush) {
            SelectObject(custom.hdc, oldBrush);
        }
        if (oldPen) {
            SelectObject(custom.hdc, oldPen);
        }
        if (brush) {
            DeleteObject(brush);
        }
        if (pen) {
            DeleteObject(pen);
        }
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

        const LONG_PTR rebarStyle = GetWindowLongPtrW(hwnd, GWL_STYLE);
        const bool verticalOrientation = ((rebarStyle & CCS_VERT) == CCS_VERT) ||
                                          ((clientRect.bottom - clientRect.top) >
                                           (clientRect.right - clientRect.left));
        const int rowTolerance = ScaleByDpi(2, verticalOrientation ? DpiX() : DpiY());

        const UINT bandCount = static_cast<UINT>(SendMessageW(hwnd, RB_GETBANDCOUNT, 0, 0));
        RECT previousRect{};
        bool havePrevious = false;
        int previousAxis = 0;

        for (UINT index = 0; index < bandCount; ++index) {
            REBARBANDINFOW bandInfo{};
            bandInfo.cbSize = sizeof(bandInfo);
            bandInfo.fMask = RBBIM_STYLE;
            if (!SendMessageW(hwnd, RB_GETBANDINFO, index, reinterpret_cast<LPARAM>(&bandInfo))) {
                continue;
            }
            if ((bandInfo.fStyle & RBBS_HIDDEN) != 0) {
                continue;
            }

            RECT bandRect{};
            if (!SendMessageW(hwnd, RB_GETRECT, index, reinterpret_cast<LPARAM>(&bandRect))) {
                continue;
            }
            if (bandRect.bottom <= bandRect.top || bandRect.right <= bandRect.left) {
                continue;
            }

            const int currentAxis = verticalOrientation ? bandRect.left : bandRect.top;
            const bool startsNewBandLine = (bandInfo.fStyle & RBBS_BREAK) != 0;

            if (!havePrevious || startsNewBandLine ||
                !ApproximatelyEqual(currentAxis, previousAxis, rowTolerance)) {
                previousRect = bandRect;
                previousAxis = currentAxis;
                havePrevious = true;
                continue;
            }

            RECT candidateLine{};
            if (verticalOrientation) {
                const int boundary = previousRect.bottom;
                candidateLine.left = std::min(previousRect.left, bandRect.left);
                candidateLine.right = std::max(previousRect.right, bandRect.right);
                candidateLine.top = boundary - lineThickness / 2;
                candidateLine.bottom = candidateLine.top + lineThickness;
            } else {
                const int boundary = previousRect.right;
                candidateLine.top = std::min(previousRect.top, bandRect.top);
                candidateLine.bottom = std::max(previousRect.bottom, bandRect.bottom);
                candidateLine.left = boundary - lineThickness / 2;
                candidateLine.right = candidateLine.left + lineThickness;
            }

            RECT lineRect{};
            if (!IntersectRect(&lineRect, &candidateLine, &paintRect) || lineRect.right <= lineRect.left ||
                lineRect.bottom <= lineRect.top) {
                previousRect = bandRect;
                previousAxis = currentAxis;
                havePrevious = true;
                continue;
            }

            RECT candidateHalo = lineRect;
            if (verticalOrientation) {
                const int centerY = (lineRect.top + lineRect.bottom) / 2;
                candidateHalo.top = centerY - haloThickness / 2;
                candidateHalo.bottom = candidateHalo.top + haloThickness;
                candidateHalo.left = lineRect.left;
                candidateHalo.right = lineRect.right;
            } else {
                const int centerX = (lineRect.left + lineRect.right) / 2;
                candidateHalo.left = centerX - haloThickness / 2;
                candidateHalo.right = candidateHalo.left + haloThickness;
                candidateHalo.top = lineRect.top;
                candidateHalo.bottom = lineRect.bottom;
            }

            RECT haloRect{};
            if (IntersectRect(&haloRect, &candidateHalo, &paintRect) && haloRect.right > haloRect.left &&
                haloRect.bottom > haloRect.top) {
                FillGradientRect(graphics, colors, RectToGdiplus(haloRect), kHaloAlpha);
            }
            FillGradientRect(graphics, colors, RectToGdiplus(lineRect), kLineAlpha);

            previousRect = bandRect;
            previousAxis = currentAxis;
            havePrevious = true;
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
                *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT;
                return true;
            case CDDS_ITEMPREPAINT:
                if (IsSeparatorButton(*custom)) {
                    *result = CDRF_SKIPDEFAULT | CDRF_NOTIFYPOSTPAINT;
                    return true;
                }
                break;
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
    bool IsSeparatorButton(const NMCUSTOMDRAW& custom) const {
        HWND hwnd = Handle();
        if (!hwnd || !IsWindow(hwnd)) {
            return false;
        }

        const RECT& rect = custom.rc;
        if (rect.right <= rect.left || rect.bottom <= rect.top) {
            return false;
        }

        ToolbarHitTestInfo hit{};
        hit.hwnd = hwnd;
        hit.pt.x = rect.left + (rect.right - rect.left) / 2;
        hit.pt.y = rect.top + (rect.bottom - rect.top) / 2;

        const LRESULT index = SendMessageW(hwnd, TB_HITTEST, 0, reinterpret_cast<LPARAM>(&hit));
        if (index < 0) {
            return false;
        }

        TBBUTTON button{};
        if (!SendMessageW(hwnd, TB_GETBUTTON, index, reinterpret_cast<LPARAM>(&button))) {
            return false;
        }

        return (button.fsStyle & TBSTYLE_SEP) != 0;
    }

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

        const LONG_PTR toolbarStyle = GetWindowLongPtrW(hwnd, GWL_STYLE);
        const bool verticalOrientation = ((toolbarStyle & CCS_VERT) == CCS_VERT) ||
                                          ((clientRect.bottom - clientRect.top) >
                                           (clientRect.right - clientRect.left));

        const int buttonCount = static_cast<int>(SendMessageW(hwnd, TB_BUTTONCOUNT, 0, 0));
        for (int index = 0; index < buttonCount; ++index) {
            TBBUTTON button{};
            if (!SendMessageW(hwnd, TB_GETBUTTON, index, reinterpret_cast<LPARAM>(&button))) {
                continue;
            }
            const bool isHidden = (button.fsState & TBSTATE_HIDDEN) != 0;
#if defined(TBSTATE_INVISIBLE)
            const bool isInvisible = (button.fsState & TBSTATE_INVISIBLE) != 0;
#else
            const bool isInvisible = false;
#endif
            if (isHidden || isInvisible) {
                continue;
            }
            if ((button.fsStyle & TBSTYLE_SEP) == 0) {
                continue;
            }

            RECT separatorRect{};
            if (!SendMessageW(hwnd, TB_GETITEMRECT, index, reinterpret_cast<LPARAM>(&separatorRect))) {
                continue;
            }
            if (separatorRect.bottom <= separatorRect.top || separatorRect.right <= separatorRect.left) {
                continue;
            }

            RECT intersection{};
            if (!IntersectRect(&intersection, &separatorRect, &paintRect)) {
                continue;
            }

            RECT candidateLine = intersection;
            if (verticalOrientation) {
                const int centerY = (intersection.top + intersection.bottom) / 2;
                candidateLine.top = centerY - lineThickness / 2;
                candidateLine.bottom = candidateLine.top + lineThickness;
            } else {
                const int centerX = (intersection.left + intersection.right) / 2;
                candidateLine.left = centerX - lineThickness / 2;
                candidateLine.right = candidateLine.left + lineThickness;
            }

            RECT lineRect{};
            if (!IntersectRect(&lineRect, &candidateLine, &intersection) || lineRect.right <= lineRect.left ||
                lineRect.bottom <= lineRect.top) {
                continue;
            }

            RECT candidateHalo = intersection;
            if (verticalOrientation) {
                const int centerY = (lineRect.top + lineRect.bottom) / 2;
                candidateHalo.top = centerY - haloThickness / 2;
                candidateHalo.bottom = candidateHalo.top + haloThickness;
            } else {
                const int centerX = (lineRect.left + lineRect.right) / 2;
                candidateHalo.left = centerX - haloThickness / 2;
                candidateHalo.right = candidateHalo.left + haloThickness;
            }

            RECT haloRect{};
            if (IntersectRect(&haloRect, &candidateHalo, &paintRect) && haloRect.right > haloRect.left &&
                haloRect.bottom > haloRect.top) {
                FillGradientRect(graphics, colors, RectToGdiplus(haloRect), kHaloAlpha);
            }
            FillGradientRect(graphics, colors, RectToGdiplus(lineRect), kLineAlpha);
        }

        EndBufferedPaint(buffer, TRUE);
    }
};

class EditGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    void OnAttached() override {
        HWND hwnd = Handle();
        if (hwnd && IsWindow(hwnd)) {
            SetWindowTheme(hwnd, L"", L"");
        }
    }

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

        const int frameThicknessX = ScaleByDpi(1, DpiX());
        const int frameThicknessY = ScaleByDpi(1, DpiY());
        const int haloThicknessX = ScaleByDpi(3, DpiX());
        const int haloThicknessY = ScaleByDpi(3, DpiY());

        RECT inner = clientRect;
        InflateRect(&inner, -frameThicknessX, -frameThicknessY);
        RECT frame = clientRect;
        RECT halo = inner;
        InflateRect(&halo, haloThicknessX, haloThicknessY);

        bool mergedPainted = false;
        if (inner.right > inner.left && inner.bottom > inner.top) {
            RECT selfScreenRect{};
            if (GetWindowRect(hwnd, &selfScreenRect)) {
                std::vector<RECT> siblingScreenRects;
                const int verticalTolerance = ScaleByDpi(6, DpiY());
                CollectSiblingEditRects(hwnd, selfScreenRect, verticalTolerance, siblingScreenRects);

                std::vector<RECT> siblingLocalRects;
                siblingLocalRects.reserve(siblingScreenRects.size());
                for (const RECT& rect : siblingScreenRects) {
                    siblingLocalRects.push_back(MapScreenRectToWindow(hwnd, rect));
                }

                if (!siblingLocalRects.empty()) {
                    std::sort(siblingLocalRects.begin(), siblingLocalRects.end(),
                              [](const RECT& lhs, const RECT& rhs) {
                                  if (lhs.left == rhs.left) {
                                      return lhs.top < rhs.top;
                                  }
                                  return lhs.left < rhs.left;
                              });

                    RECT mergedLocal = siblingLocalRects.front();
                    for (size_t index = 1; index < siblingLocalRects.size(); ++index) {
                        const RECT& candidate = siblingLocalRects[index];
                        mergedLocal.left = std::min(mergedLocal.left, candidate.left);
                        mergedLocal.top = std::min(mergedLocal.top, candidate.top);
                        mergedLocal.right = std::max(mergedLocal.right, candidate.right);
                        mergedLocal.bottom = std::max(mergedLocal.bottom, candidate.bottom);
                    }

                    if (mergedLocal.right > mergedLocal.left && mergedLocal.bottom > mergedLocal.top) {
                        auto fillClippedRect = [&](RECT rect, BYTE alpha, float angle) {
                            RECT clipped{};
                            if (!IntersectRect(&clipped, &rect, &paintRect)) {
                                return;
                            }
                            if (clipped.right <= clipped.left || clipped.bottom <= clipped.top) {
                                return;
                            }
                            FillGradientRect(graphics, colors, RectToGdiplus(clipped), alpha, angle);
                        };

                        const float horizontalAngle = 90.0f;
                        const float verticalAngle = 0.0f;

                        RECT topHaloRect = {mergedLocal.left - haloThicknessX, inner.top - haloThicknessY,
                                            mergedLocal.right + haloThicknessX, inner.top};
                        RECT bottomHaloRect = {mergedLocal.left - haloThicknessX, inner.bottom,
                                               mergedLocal.right + haloThicknessX,
                                               inner.bottom + haloThicknessY};
                        fillClippedRect(topHaloRect, kFrameHaloAlpha, horizontalAngle);
                        fillClippedRect(bottomHaloRect, kFrameHaloAlpha, horizontalAngle);

                        RECT topFrameRect = {mergedLocal.left, inner.top - frameThicknessY, mergedLocal.right,
                                             inner.top};
                        RECT bottomFrameRect = {mergedLocal.left, inner.bottom, mergedLocal.right,
                                                inner.bottom + frameThicknessY};
                        fillClippedRect(topFrameRect, kFrameAlpha, horizontalAngle);
                        fillClippedRect(bottomFrameRect, kFrameAlpha, horizontalAngle);

                        for (const RECT& rect : siblingLocalRects) {
                            RECT leftHaloRect = {rect.left - haloThicknessX, inner.top - haloThicknessY,
                                                 rect.left + haloThicknessX,
                                                 inner.bottom + haloThicknessY};
                            RECT rightHaloRect = {rect.right - haloThicknessX, inner.top - haloThicknessY,
                                                  rect.right + haloThicknessX,
                                                  inner.bottom + haloThicknessY};
                            fillClippedRect(leftHaloRect, kFrameHaloAlpha, verticalAngle);
                            fillClippedRect(rightHaloRect, kFrameHaloAlpha, verticalAngle);

                            RECT leftFrameRect = {rect.left, inner.top - frameThicknessY,
                                                  rect.left + frameThicknessX,
                                                  inner.bottom + frameThicknessY};
                            RECT rightFrameRect = {rect.right - frameThicknessX, inner.top - frameThicknessY,
                                                   rect.right, inner.bottom + frameThicknessY};
                            fillClippedRect(leftFrameRect, kFrameAlpha, verticalAngle);
                            fillClippedRect(rightFrameRect, kFrameAlpha, verticalAngle);
                        }

                        mergedPainted = true;
                    }
                }
            }
        }

        if (!mergedPainted) {
            FillFrameRegion(graphics, colors, halo, inner, kFrameHaloAlpha, 0.0f);
            FillFrameRegion(graphics, colors, frame, inner, kFrameAlpha, 0.0f);
        }

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

        std::vector<RECT> highlightRects;
        if (!EnumerateDirectUiRectangles(hwnd, clientRect, highlightRects)) {
            return;
        }

        const int toleranceX = ScaleByDpi(2, DpiX());
        const int toleranceY = ScaleByDpi(2, DpiY());

        std::vector<RECT> filteredRects;
        filteredRects.reserve(highlightRects.size());
        for (const RECT& rect : highlightRects) {
            RECT intersection{};
            if (!IntersectRect(&intersection, &rect, &paintRect)) {
                continue;
            }

            if (ApproximatelyEqual(rect.left, clientRect.left, toleranceX) &&
                ApproximatelyEqual(rect.top, clientRect.top, toleranceY) &&
                ApproximatelyEqual(rect.right, clientRect.right, toleranceX) &&
                ApproximatelyEqual(rect.bottom, clientRect.bottom, toleranceY)) {
                continue;
            }

            filteredRects.push_back(rect);
        }

        if (filteredRects.empty()) {
            return;
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

        const int frameThicknessX = ScaleByDpi(1, DpiX());
        const int frameThicknessY = ScaleByDpi(1, DpiY());
        const int haloThicknessX = std::max(frameThicknessX * 2, ScaleByDpi(3, DpiX()));
        const int haloThicknessY = std::max(frameThicknessY * 2, ScaleByDpi(3, DpiY()));

        for (const RECT& rect : filteredRects) {
            RECT paintCheck{};
            if (!IntersectRect(&paintCheck, &rect, &paintRect)) {
                continue;
            }

            RECT inner = rect;
            if ((rect.right - rect.left) > frameThicknessX * 2) {
                inner.left += frameThicknessX;
                inner.right -= frameThicknessX;
            }
            if ((rect.bottom - rect.top) > frameThicknessY * 2) {
                inner.top += frameThicknessY;
                inner.bottom -= frameThicknessY;
            }

            if (inner.right <= inner.left || inner.bottom <= inner.top) {
                continue;
            }

            RECT frame = rect;
            RECT halo = rect;
            InflateRect(&halo, haloThicknessX, haloThicknessY);

            RECT haloClipped = halo;
            if (!IntersectRect(&haloClipped, &halo, &clientRect)) {
                continue;
            }
            if (!IntersectRect(&haloClipped, &haloClipped, &paintRect)) {
                continue;
            }

            RECT frameClipped = frame;
            if (!IntersectRect(&frameClipped, &frame, &clientRect)) {
                continue;
            }
            if (!IntersectRect(&frameClipped, &frameClipped, &paintRect)) {
                continue;
            }

            RECT innerClipped = inner;
            if (!IntersectRect(&innerClipped, &inner, &clientRect)) {
                continue;
            }
            if (!IntersectRect(&innerClipped, &innerClipped, &paintRect)) {
                continue;
            }

            FillFrameRegion(graphics, colors, haloClipped, innerClipped, kFrameHaloAlpha);
            FillFrameRegion(graphics, colors, frameClipped, innerClipped, kFrameAlpha);
        }

        EndBufferedPaint(buffer, TRUE);
    }
};

class ScrollBarGlowSurface : public ExplorerGlowSurface {
public:
    using ExplorerGlowSurface::ExplorerGlowSurface;

protected:
    void OnPaint(HDC targetDc, const RECT& clipRect, const GlowColorSet& colors) override {
        HWND hwnd = Handle();
        if (!hwnd || !IsWindow(hwnd)) {
            return;
        }

        auto scrollbarDefinition = Coordinator().ResolveScrollbarDefinition();
        if (!scrollbarDefinition.has_value()) {
            return;
        }
        const ScrollbarGlowDefinition& definition = *scrollbarDefinition;

        RECT clientRect = GetClientRectSafe(hwnd);
        if (clientRect.right <= clientRect.left || clientRect.bottom <= clientRect.top) {
            return;
        }

        SCROLLBARINFO info{};
        info.cbSize = sizeof(info);
        if (!GetScrollBarInfo(hwnd, OBJID_CLIENT, &info)) {
            return;
        }

        if (info.rgstate[0] & (STATE_SYSTEM_INVISIBLE | STATE_SYSTEM_OFFSCREEN | STATE_SYSTEM_UNAVAILABLE)) {
            return;
        }

        RECT scrollRect = info.rcScrollBar;
        if (scrollRect.right <= scrollRect.left || scrollRect.bottom <= scrollRect.top) {
            return;
        }

        RECT trackRect = scrollRect;
        RECT thumbRect = scrollRect;
        const bool infoVertical =
            (scrollRect.bottom - scrollRect.top) >= (scrollRect.right - scrollRect.left);
        if (infoVertical) {
            const LONG adjustedTop = std::min(trackRect.bottom, trackRect.top + info.dxyLineButton);
            const LONG adjustedBottom = std::max(adjustedTop, trackRect.bottom - info.dxyLineButton);
            trackRect.top = adjustedTop;
            trackRect.bottom = adjustedBottom;

            const LONG thumbTop = std::clamp<LONG>(info.xyThumbTop, trackRect.top, trackRect.bottom);
            const LONG thumbBottom = std::clamp<LONG>(info.xyThumbBottom, thumbTop, trackRect.bottom);
            thumbRect.top = thumbTop;
            thumbRect.bottom = thumbBottom;
        } else {
            const LONG adjustedLeft = std::min(trackRect.right, trackRect.left + info.dxyLineButton);
            const LONG adjustedRight = std::max(adjustedLeft, trackRect.right - info.dxyLineButton);
            trackRect.left = adjustedLeft;
            trackRect.right = adjustedRight;

            const LONG thumbLeft = std::clamp<LONG>(info.xyThumbTop, trackRect.left, trackRect.right);
            const LONG thumbRight = std::clamp<LONG>(info.xyThumbBottom, thumbLeft, trackRect.right);
            thumbRect.left = thumbLeft;
            thumbRect.right = thumbRight;
        }
        MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&trackRect), 2);
        MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&thumbRect), 2);

        if (trackRect.right <= trackRect.left || trackRect.bottom <= trackRect.top) {
            return;
        }

        RECT effectiveClip{};
        if (!IntersectRect(&effectiveClip, &clipRect, &clientRect)) {
            effectiveClip = clientRect;
        }

        const bool vertical = infoVertical;
        const int crossExtent = vertical ? (trackRect.right - trackRect.left) : (trackRect.bottom - trackRect.top);
        const UINT crossDpi = vertical ? DpiX() : DpiY();
        const UINT alongDpi = vertical ? DpiY() : DpiX();

        const int baseLine = ScaleByDpi(2, crossDpi);
        const int lineThickness = std::clamp(baseLine, 1, std::max(crossExtent, 1));
        const int haloPadding = std::max(lineThickness, ScaleByDpi(3, crossDpi));
        const int thumbAlongPadding = ScaleByDpi(6, alongDpi);

        Gdiplus::Graphics graphics(targetDc);
        graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);

        RECT trackHalo = trackRect;
        if (vertical) {
            InflateRect(&trackHalo, haloPadding, 0);
        } else {
            InflateRect(&trackHalo, 0, haloPadding);
        }
        RECT trackHaloClip{};
        if (IntersectRect(&trackHaloClip, &trackHalo, &effectiveClip) && !IsRectEmpty(&trackHaloClip)) {
            FillGradientRect(graphics, colors, RectToGdiplus(trackHaloClip), definition.trackHaloAlpha);
        }

        RECT lineRect = trackRect;
        if (vertical) {
            const int center = trackRect.left + ((trackRect.right - trackRect.left) - lineThickness) / 2;
            lineRect.left = center;
            lineRect.right = center + lineThickness;
        } else {
            const int center = trackRect.top + ((trackRect.bottom - trackRect.top) - lineThickness) / 2;
            lineRect.top = center;
            lineRect.bottom = center + lineThickness;
        }
        RECT lineClip{};
        if (IntersectRect(&lineClip, &lineRect, &effectiveClip) && !IsRectEmpty(&lineClip)) {
            FillGradientRect(graphics, colors, RectToGdiplus(lineClip), definition.trackLineAlpha);
        }

        const bool thumbVisible =
            !(info.rgstate[3] & (STATE_SYSTEM_INVISIBLE | STATE_SYSTEM_OFFSCREEN | STATE_SYSTEM_UNAVAILABLE));
        if (thumbVisible && thumbRect.right > thumbRect.left && thumbRect.bottom > thumbRect.top) {
            RECT thumbHalo = thumbRect;
            if (vertical) {
                InflateRect(&thumbHalo, haloPadding, thumbAlongPadding);
            } else {
                InflateRect(&thumbHalo, thumbAlongPadding, haloPadding);
            }

            RECT thumbHaloClip{};
            if (IntersectRect(&thumbHaloClip, &thumbHalo, &effectiveClip) && !IsRectEmpty(&thumbHaloClip)) {
                FillGradientRect(graphics, colors, RectToGdiplus(thumbHaloClip), definition.thumbHaloAlpha);
            }

            RECT thumbClip{};
            if (IntersectRect(&thumbClip, &thumbRect, &effectiveClip) && !IsRectEmpty(&thumbClip)) {
                FillGradientRect(graphics, colors, RectToGdiplus(thumbClip), definition.thumbFillAlpha);
            }
        }
    }
};

}  // namespace

ExplorerGlowCoordinator::ExplorerGlowCoordinator() = default;

void ExplorerGlowCoordinator::Configure(const ShellTabsOptions& options) {
    m_glowEnabled = options.enableNeonGlow;
    m_palette = options.glowPalette;
    m_breadcrumbFontGradient.enabled = options.enableBreadcrumbFontGradient;
    m_breadcrumbFontGradient.brightness = std::clamp(options.breadcrumbFontBrightness, 0, 100);
    m_breadcrumbFontGradient.useCustomFontColors = options.useCustomBreadcrumbFontColors;
    m_breadcrumbFontGradient.useCustomGradientColors = options.useCustomBreadcrumbGradientColors;
    m_breadcrumbFontGradient.fontGradientStartColor = options.breadcrumbFontGradientStartColor;
    m_breadcrumbFontGradient.fontGradientEndColor = options.breadcrumbFontGradientEndColor;
    m_breadcrumbFontGradient.gradientStartColor = options.breadcrumbGradientStartColor;
    m_breadcrumbFontGradient.gradientEndColor = options.breadcrumbGradientEndColor;
    m_bitmapInterceptEnabled = options.enableBitmapIntercept;
    RefreshAccessibilityState();
    UpdateAccentColor();
    RefreshDescriptorAccessibility();
}

bool ExplorerGlowCoordinator::HandleThemeChanged() {
    const bool accessibilityChanged = RefreshAccessibilityState();
    const COLORREF previousAccent = m_accentColor;
    UpdateAccentColor();
    if (accessibilityChanged) {
        RefreshDescriptorAccessibility();
    }
    return accessibilityChanged || (previousAccent != m_accentColor);
}

bool ExplorerGlowCoordinator::HandleSettingChanged() {
    const bool accessibilityChanged = RefreshAccessibilityState();
    const COLORREF previousAccent = m_accentColor;
    UpdateAccentColor();
    if (accessibilityChanged) {
        RefreshDescriptorAccessibility();
    }
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

std::optional<ScrollbarGlowDefinition> ExplorerGlowCoordinator::ResolveScrollbarDefinition() const {
    ScrollbarGlowDefinition definition{};
    GlowColorSet colors = ResolveColors(ExplorerSurfaceKind::Scrollbar);
    if (!colors.valid) {
        return std::nullopt;
    }

    definition.colors = colors;
    definition.trackLineAlpha = kLineAlpha;
    definition.trackHaloAlpha = kHaloAlpha;
    definition.thumbFillAlpha = kFrameAlpha;
    definition.thumbHaloAlpha = kFrameHaloAlpha;
    return definition;
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
        case ExplorerSurfaceKind::Scrollbar:
            return &m_palette.scrollbars;
        case ExplorerSurfaceKind::DirectUi:
            return &m_palette.directUi;
        case ExplorerSurfaceKind::PopupMenu:
            return &m_palette.popupMenus;
        case ExplorerSurfaceKind::Tooltip:
            return &m_palette.tooltips;
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

SurfaceColorDescriptor* ExplorerGlowCoordinator::AcquireSurfaceDescriptor(HWND hwnd, ExplorerSurfaceKind kind) {
    if (!hwnd) {
        return nullptr;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    auto& slot = m_surfaceDescriptors[hwnd];
    if (!slot) {
        slot = std::make_unique<SurfaceColorDescriptor>();
    }
    slot->kind = kind;
    slot->role = SurfacePaintRole::Generic;
    slot->accessibilityOptOut = slot->userAccessibilityOptOut || m_highContrastActive;
    if (slot->accessibilityOptOut) {
        slot->forcedHooks = false;
    }
    return slot.get();
}

SurfaceColorDescriptor* ExplorerGlowCoordinator::LookupSurfaceDescriptor(HWND hwnd) const {
    if (!hwnd) {
        return nullptr;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    auto it = m_surfaceDescriptors.find(hwnd);
    if (it == m_surfaceDescriptors.end() || !it->second) {
        return nullptr;
    }
    return it->second.get();
}

void ExplorerGlowCoordinator::ReleaseSurfaceDescriptor(HWND hwnd) {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    m_surfaceDescriptors.erase(hwnd);
}

void ExplorerGlowCoordinator::UpdateSurfaceDescriptor(HWND hwnd, const SurfaceColorDescriptor& descriptor) {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    auto& slot = m_surfaceDescriptors[hwnd];
    if (!slot) {
        slot = std::make_unique<SurfaceColorDescriptor>();
    }
    const bool wasForced = slot->forcedHooks;
    slot->kind = descriptor.kind;
    slot->role = descriptor.role;
    slot->fillColors = descriptor.fillColors;
    slot->fillOverride = descriptor.fillOverride;
    slot->textColor = descriptor.textColor;
    slot->textOverride = descriptor.textOverride;
    slot->backgroundColor = descriptor.backgroundColor;
    slot->backgroundOverride = descriptor.backgroundOverride;
    slot->forceOpaqueBackground = descriptor.forceOpaqueBackground;
    slot->backgroundPaintCallback = descriptor.backgroundPaintCallback;
    slot->backgroundPaintContext = descriptor.backgroundPaintContext;
    slot->userAccessibilityOptOut = descriptor.userAccessibilityOptOut;
    slot->accessibilityOptOut = descriptor.userAccessibilityOptOut || m_highContrastActive;
    slot->forcedHooks = wasForced && !slot->accessibilityOptOut;
}

void ExplorerGlowCoordinator::SetSurfaceForcedHooks(HWND hwnd, bool forced) {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    auto it = m_surfaceDescriptors.find(hwnd);
    if (it == m_surfaceDescriptors.end() || !it->second) {
        return;
    }
    SurfaceColorDescriptor& descriptor = *it->second;
    if (descriptor.accessibilityOptOut) {
        descriptor.forcedHooks = false;
        return;
    }
    descriptor.forcedHooks = forced;
}

void ExplorerGlowCoordinator::SetSurfaceRole(HWND hwnd, SurfacePaintRole role) {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    auto it = m_surfaceDescriptors.find(hwnd);
    if (it == m_surfaceDescriptors.end() || !it->second) {
        return;
    }
    it->second->role = role;
}

void ExplorerGlowCoordinator::SetSurfaceAccessibilityOptOut(HWND hwnd, bool optOut) {
    if (!hwnd) {
        return;
    }
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    auto it = m_surfaceDescriptors.find(hwnd);
    if (it == m_surfaceDescriptors.end() || !it->second) {
        return;
    }
    SurfaceColorDescriptor& descriptor = *it->second;
    descriptor.userAccessibilityOptOut = optOut;
    descriptor.accessibilityOptOut = optOut || m_highContrastActive;
    if (descriptor.accessibilityOptOut) {
        descriptor.forcedHooks = false;
    }
}

void ExplorerGlowCoordinator::RefreshDescriptorAccessibility() {
    std::lock_guard<std::mutex> lock(m_descriptorMutex);
    for (auto& entry : m_surfaceDescriptors) {
        if (!entry.second) {
            continue;
        }
        entry.second->accessibilityOptOut = entry.second->userAccessibilityOptOut || m_highContrastActive;
        if (entry.second->accessibilityOptOut) {
            entry.second->forcedHooks = false;
        }
    }
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
        case WM_NCPAINT:
            if (Kind() == ExplorerSurfaceKind::Edit) {
                HRGN updateRegion = reinterpret_cast<HRGN>(wParam);
                const bool hasValidRegion = updateRegion && updateRegion != reinterpret_cast<HRGN>(1);
                const UINT dcFlags = DCX_WINDOW | DCX_CACHE | DCX_CLIPSIBLINGS | DCX_CLIPCHILDREN |
                                     (hasValidRegion ? DCX_INTERSECTRGN : 0);
                HDC targetDc = GetDCEx(hwnd, hasValidRegion ? updateRegion : nullptr, dcFlags);
                if (targetDc) {
                    if (Coordinator().ShouldRenderSurface(Kind())) {
                        RECT clip{};
                        if (GetClipBox(targetDc, &clip) == ERROR || IsRectEmpty(&clip)) {
                            clip = GetClientRectSafe(hwnd);
                        }
                        if (!IsRectEmpty(&clip)) {
                            PaintInternal(targetDc, clip);
                        }
                    }
                    ReleaseDC(hwnd, targetDc);
                }

                if (hasValidRegion) {
                    ValidateRgn(hwnd, updateRegion);
                } else {
                    ValidateRect(hwnd, nullptr);
                }
                return 0;
            }
            break;
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

    const bool glowActive = Coordinator().ShouldRenderSurface(Kind());
    const bool gradientActive =
        (Kind() == ExplorerSurfaceKind::Edit && Coordinator().BreadcrumbFontGradient().enabled);

    if (!glowActive && !gradientActive) {
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

    if (gradientActive) {
        const auto& gradientConfig = Coordinator().BreadcrumbFontGradient();
        GradientEditRenderOptions options;
        options.hideCaret = true;
        options.requestEraseBackground = false;
        if (!IsRectEmpty(&clipRect)) {
            options.clipRect = clipRect;
        }
        RenderGradientEditContent(hwnd, targetDc, gradientConfig, options);
    }

    if (glowActive && hasClip) {
        PaintInternal(targetDc, clipRect);
    }

    if (releaseDc) {
        ReleaseDC(hwnd, targetDc);
    }

    return defResult;
}

LRESULT ExplorerGlowSurface::HandlePrintClient(HWND hwnd, WPARAM wParam, LPARAM lParam) {
    LRESULT defResult = DefSubclassProc(hwnd, WM_PRINTCLIENT, wParam, lParam);

    HDC targetDc = reinterpret_cast<HDC>(wParam);
    if (!targetDc) {
        return defResult;
    }

    const bool glowActive = Coordinator().ShouldRenderSurface(Kind());
    const bool gradientActive =
        (Kind() == ExplorerSurfaceKind::Edit && Coordinator().BreadcrumbFontGradient().enabled);

    RECT clip{};
    if (GetClipBox(targetDc, &clip) == ERROR || IsRectEmpty(&clip)) {
        clip = GetClientRectSafe(hwnd);
    }

    if (gradientActive) {
        const auto& gradientConfig = Coordinator().BreadcrumbFontGradient();
        GradientEditRenderOptions options;
        options.hideCaret = false;
        options.requestEraseBackground = true;
        if (!IsRectEmpty(&clip)) {
            options.clipRect = clip;
        }
        RenderGradientEditContent(hwnd, targetDc, gradientConfig, options);
    }

    if (glowActive && !IsRectEmpty(&clip)) {
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
        case ExplorerSurfaceKind::Scrollbar:
            return std::make_unique<ScrollBarGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::DirectUi:
            return std::make_unique<DirectUiGlowSurface>(kind, coordinator);
        case ExplorerSurfaceKind::PopupMenu:
        case ExplorerSurfaceKind::Tooltip:
            return nullptr;
        default:
            return nullptr;
    }
}

}  // namespace shelltabs

