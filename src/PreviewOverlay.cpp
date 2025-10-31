#include "PreviewOverlay.h"

#include <algorithm>
#include <mutex>
#include <string>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <CommCtrl.h>

#include "Module.h"

#include <gdiplus.h>

namespace shelltabs {
namespace {

constexpr wchar_t kPreviewWindowClassName[] = L"ShellTabsPreviewWindow";

constexpr int kPreviewFramePadding = 6;
constexpr int kPreviewFrameCornerRadius = 12;
constexpr int kPreviewFrameHighlightThickness = 1;

bool EnsureGdiplusInitialized() {
    static std::once_flag initOnce;
    static bool initialized = false;
    static ULONG_PTR token = 0;
    std::call_once(initOnce, [&]() {
        Gdiplus::GdiplusStartupInput input;
        if (Gdiplus::GdiplusStartup(&token, &input, nullptr) == Gdiplus::Ok) {
            initialized = true;
        }
    });
    return initialized;
}

void BuildRoundedRectPath(Gdiplus::GraphicsPath* path, const Gdiplus::RectF& rect, Gdiplus::REAL radius) {
    if (!path) {
        return;
    }

    path->Reset();

    const Gdiplus::REAL limitedRadius = std::max<Gdiplus::REAL>(
        0.0f,
        std::min(radius, std::min(rect.Width, rect.Height) / 2.0f));
    if (limitedRadius <= 0.0f) {
        path->AddRectangle(rect);
        return;
    }

    const Gdiplus::REAL diameter = limitedRadius * 2.0f;
    Gdiplus::RectF arc(rect.X, rect.Y, diameter, diameter);

    path->StartFigure();
    path->AddArc(arc, 180.0f, 90.0f);
    arc.X = rect.GetRight() - diameter;
    path->AddArc(arc, 270.0f, 90.0f);
    arc.Y = rect.GetBottom() - diameter;
    path->AddArc(arc, 0.0f, 90.0f);
    arc.X = rect.X;
    path->AddArc(arc, 90.0f, 90.0f);
    path->CloseFigure();
}

HBITMAP CreateFramedPreviewBitmap(HBITMAP sourceBitmap, const SIZE& sourceSize, SIZE* outSize) {
    if (!sourceBitmap || sourceSize.cx <= 0 || sourceSize.cy <= 0) {
        return nullptr;
    }

    if (!EnsureGdiplusInitialized()) {
        return nullptr;
    }

    const int width = sourceSize.cx + kPreviewFramePadding * 2;
    const int height = sourceSize.cy + kPreviewFramePadding * 2;

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(info.bmiHeader);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP framedBitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!framedBitmap || !bits) {
        if (framedBitmap) {
            DeleteObject(framedBitmap);
        }
        return nullptr;
    }

    Gdiplus::Bitmap destBitmap(width, height, width * 4, PixelFormat32bppPARGB, static_cast<BYTE*>(bits));
    if (destBitmap.GetLastStatus() != Gdiplus::Ok) {
        DeleteObject(framedBitmap);
        return nullptr;
    }

    Gdiplus::Graphics graphics(&destBitmap);
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        DeleteObject(framedBitmap);
        return nullptr;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
    graphics.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    graphics.Clear(Gdiplus::Color(0, 0, 0, 0));

    const COLORREF frameRgb = GetSysColor(COLOR_HIGHLIGHT);
    const Gdiplus::Color frameColor(230, GetRValue(frameRgb), GetGValue(frameRgb), GetBValue(frameRgb));
    const Gdiplus::Color shadowColor(200, 16, 16, 16);

    const Gdiplus::RectF outerRect(0.0f, 0.0f, static_cast<Gdiplus::REAL>(width), static_cast<Gdiplus::REAL>(height));
    Gdiplus::GraphicsPath outerPath;
    BuildRoundedRectPath(&outerPath, outerRect, static_cast<Gdiplus::REAL>(kPreviewFrameCornerRadius));

    Gdiplus::SolidBrush shadowBrush(shadowColor);
    graphics.FillPath(&shadowBrush, &outerPath);

    Gdiplus::RectF frameRect(static_cast<Gdiplus::REAL>(kPreviewFramePadding),
                             static_cast<Gdiplus::REAL>(kPreviewFramePadding),
                             static_cast<Gdiplus::REAL>(sourceSize.cx),
                             static_cast<Gdiplus::REAL>(sourceSize.cy));
    const Gdiplus::REAL innerRadius = std::max<Gdiplus::REAL>(
        0.0f,
        static_cast<Gdiplus::REAL>(kPreviewFrameCornerRadius - kPreviewFramePadding));
    Gdiplus::GraphicsPath innerPath;
    BuildRoundedRectPath(&innerPath, frameRect, innerRadius);

    Gdiplus::Region frameRegion(&outerPath);
    frameRegion.Exclude(&innerPath);
    Gdiplus::SolidBrush frameBrush(frameColor);
    graphics.FillRegion(&frameBrush, &frameRegion);

    graphics.SetClip(&innerPath, Gdiplus::CombineModeReplace);

    Gdiplus::Bitmap source(sourceBitmap, nullptr);
    if (source.GetLastStatus() == Gdiplus::Ok) {
        graphics.DrawImage(&source, frameRect, 0.0f, 0.0f, static_cast<Gdiplus::REAL>(sourceSize.cx),
                           static_cast<Gdiplus::REAL>(sourceSize.cy), Gdiplus::UnitPixel);
    }

    graphics.ResetClip();

    if constexpr (kPreviewFrameHighlightThickness > 0) {
        Gdiplus::Pen highlightPen(Gdiplus::Color(190, 255, 255, 255),
                                  static_cast<Gdiplus::REAL>(kPreviewFrameHighlightThickness));
        highlightPen.SetAlignment(Gdiplus::PenAlignmentInset);
        graphics.DrawPath(&highlightPen, &innerPath);
    }

    if (outSize) {
        outSize->cx = width;
        outSize->cy = height;
    }

    return framedBitmap;
}

HBITMAP CreatePlaceholderContentBitmap(const SIZE& desiredSize, const wchar_t* text, HICON icon, SIZE* outSize) {
    if (!EnsureGdiplusInitialized()) {
        return nullptr;
    }

    const wchar_t* fallbackText = L"Generating preview…";
    std::wstring label = (text && *text) ? text : fallbackText;

    const LONG minWidth = std::max<LONG>(desiredSize.cx, 128L);
    const LONG minHeight = std::max<LONG>(desiredSize.cy, 96L);
    const int width = minWidth > 0 ? static_cast<int>(minWidth) : 192;
    const int height = minHeight > 0 ? static_cast<int>(minHeight) : 128;

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(info.bmiHeader);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap || !bits) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return nullptr;
    }

    Gdiplus::Bitmap surface(width, height, width * 4, PixelFormat32bppPARGB, static_cast<BYTE*>(bits));
    if (surface.GetLastStatus() != Gdiplus::Ok) {
        DeleteObject(bitmap);
        return nullptr;
    }

    Gdiplus::Graphics graphics(&surface);
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        DeleteObject(bitmap);
        return nullptr;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
    graphics.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

    graphics.Clear(Gdiplus::Color(0, 0, 0, 0));

    const COLORREF windowRgb = GetSysColor(COLOR_WINDOW);
    const COLORREF highlightRgb = GetSysColor(COLOR_HIGHLIGHT);
    const Gdiplus::Color topColor(230, static_cast<BYTE>(GetRValue(windowRgb)),
                                  static_cast<BYTE>(GetGValue(windowRgb)),
                                  static_cast<BYTE>(GetBValue(windowRgb)));
    const Gdiplus::Color bottomColor(255,
                                     static_cast<BYTE>(std::min(255, GetRValue(windowRgb) + 10)),
                                     static_cast<BYTE>(std::min(255, GetGValue(windowRgb) + 10)),
                                     static_cast<BYTE>(std::min(255, GetBValue(windowRgb) + 10)));

    Gdiplus::RectF backgroundRect(0.0f, 0.0f, static_cast<Gdiplus::REAL>(width), static_cast<Gdiplus::REAL>(height));
    Gdiplus::LinearGradientBrush backgroundBrush(backgroundRect, topColor, bottomColor, 90.0f);
    graphics.FillRectangle(&backgroundBrush, backgroundRect);

    const Gdiplus::Color accentColor(70, GetRValue(highlightRgb), GetGValue(highlightRgb), GetBValue(highlightRgb));
    Gdiplus::Pen accentPen(accentColor, 1.0f);
    graphics.DrawRectangle(&accentPen, backgroundRect);

    Gdiplus::REAL currentY = 16.0f;
    const Gdiplus::REAL horizontalPadding = 16.0f;
    const Gdiplus::REAL availableWidth = static_cast<Gdiplus::REAL>(width) - horizontalPadding * 2.0f;

    if (icon) {
        Gdiplus::Bitmap iconBitmap(icon);
        if (iconBitmap.GetLastStatus() == Gdiplus::Ok && iconBitmap.GetWidth() > 0 && iconBitmap.GetHeight() > 0) {
            const Gdiplus::REAL maxIconWidth = availableWidth;
            const Gdiplus::REAL maxIconHeight = static_cast<Gdiplus::REAL>(height) * 0.5f;
            const Gdiplus::REAL scale = std::min({maxIconWidth / static_cast<Gdiplus::REAL>(iconBitmap.GetWidth()),
                                                  maxIconHeight / static_cast<Gdiplus::REAL>(iconBitmap.GetHeight()),
                                                  1.0f});
            const Gdiplus::REAL iconWidth = std::max<Gdiplus::REAL>(24.0f, iconBitmap.GetWidth() * scale);
            const Gdiplus::REAL iconHeight = std::max<Gdiplus::REAL>(24.0f, iconBitmap.GetHeight() * scale);
            const Gdiplus::REAL iconX = (static_cast<Gdiplus::REAL>(width) - iconWidth) / 2.0f;
            Gdiplus::RectF iconRect(iconX, currentY, iconWidth, iconHeight);
            graphics.DrawImage(&iconBitmap, iconRect);
            currentY += iconHeight + 12.0f;
        }
    }

    if (!label.empty()) {
        Gdiplus::FontFamily family(L"Segoe UI");
        Gdiplus::Font primary(&family, 12.0f, Gdiplus::FontStyleRegular, Gdiplus::UnitPoint);
        Gdiplus::Font fallbackFont(L"Segoe UI", 12.0f, Gdiplus::FontStyleRegular, Gdiplus::UnitPoint);
        Gdiplus::Font* font = nullptr;
        if (primary.GetLastStatus() == Gdiplus::Ok) {
            font = &primary;
        } else if (fallbackFont.GetLastStatus() == Gdiplus::Ok) {
            font = &fallbackFont;
        }
        if (font) {
            const COLORREF textRgb = GetSysColor(COLOR_WINDOWTEXT);
            Gdiplus::Color textColor(255, GetRValue(textRgb), GetGValue(textRgb), GetBValue(textRgb));
            Gdiplus::RectF textRect(horizontalPadding, currentY,
                                    availableWidth, static_cast<Gdiplus::REAL>(height) - currentY - 12.0f);
            Gdiplus::StringFormat format;
            format.SetAlignment(Gdiplus::StringAlignmentCenter);
            format.SetLineAlignment(Gdiplus::StringAlignmentNear);
            Gdiplus::SolidBrush labelBrush(textColor);
            graphics.DrawString(label.c_str(), static_cast<INT>(label.size()), font, textRect, &format, &labelBrush);
        }
    }

    if (outSize) {
        outSize->cx = width;
        outSize->cy = height;
    }
    return bitmap;
}

}  // namespace

PreviewOverlay::~PreviewOverlay() { Destroy(); }

ATOM PreviewOverlay::EnsureWindowClass() {
    static ATOM atom = 0;
    if (atom != 0) {
        return atom;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc = DefWindowProcW;
    wc.hInstance = GetModuleHandleInstance();
    wc.lpszClassName = kPreviewWindowClassName;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);

    atom = RegisterClassW(&wc);
    if (atom == 0 && GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
        atom = 1;  // sentinel for already registered
    }

    return atom;
}

HWND PreviewOverlay::EnsureWindow(HWND owner) {
    if (m_window && IsWindow(m_window)) {
        if (owner && GetParent(m_window) != owner) {
            SetWindowLongPtrW(m_window, GWLP_HWNDPARENT, reinterpret_cast<LONG_PTR>(owner));
        }
        m_owner = owner;
        return m_window;
    }

    if (!EnsureWindowClass()) {
        return nullptr;
    }

    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_NOACTIVATE, kPreviewWindowClassName,
                                L"", WS_POPUP, 0, 0, 0, 0, owner, nullptr, GetModuleHandleInstance(), nullptr);
    if (!hwnd) {
        return nullptr;
    }

    m_window = hwnd;
    m_owner = owner;
    return m_window;
}

bool PreviewOverlay::Show(HWND owner, HBITMAP bitmap, const SIZE& size, const POINT& screenPt,
                          const wchar_t* placeholderText, HICON placeholderIcon) {
    HWND window = EnsureWindow(owner);
    if (!window) {
        return false;
    }

    HDC screenDC = GetDC(nullptr);
    if (!screenDC) {
        return false;
    }

    HDC memDC = CreateCompatibleDC(screenDC);
    if (!memDC) {
        ReleaseDC(nullptr, screenDC);
        return false;
    }

    SIZE finalSize = size;
    HBITMAP placeholderBitmap = nullptr;
    HBITMAP sourceBitmap = bitmap;
    if (!sourceBitmap) {
        placeholderBitmap = CreatePlaceholderContentBitmap(size, placeholderText, placeholderIcon, &finalSize);
        sourceBitmap = placeholderBitmap;
    }
    if (!sourceBitmap) {
        DeleteDC(memDC);
        ReleaseDC(nullptr, screenDC);
        return false;
    }

    HBITMAP framedBitmap = CreateFramedPreviewBitmap(sourceBitmap, finalSize, &finalSize);
    HBITMAP bitmapToDraw = framedBitmap ? framedBitmap : sourceBitmap;

    HGDIOBJ oldBitmap = SelectObject(memDC, bitmapToDraw);
    if (!oldBitmap) {
        if (framedBitmap) {
            DeleteObject(framedBitmap);
        }
        if (placeholderBitmap) {
            DeleteObject(placeholderBitmap);
        }
        DeleteDC(memDC);
        ReleaseDC(nullptr, screenDC);
        return false;
    }
    POINT dest = screenPt;
    dest.x += 16;
    dest.y += 16;
    POINT src{0, 0};
    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 255;
    blend.AlphaFormat = AC_SRC_ALPHA;

    SIZE bitmapSize = finalSize;
    const BOOL updated = UpdateLayeredWindow(window, screenDC, &dest, &bitmapSize, memDC, &src, 0, &blend, ULW_ALPHA);

    SelectObject(memDC, oldBitmap);
    DeleteDC(memDC);
    ReleaseDC(nullptr, screenDC);

    if (framedBitmap) {
        DeleteObject(framedBitmap);
    }
    if (placeholderBitmap) {
        DeleteObject(placeholderBitmap);
    }

    if (!updated) {
        return false;
    }

    m_bitmapSize = finalSize;
    m_visible = true;
    ShowWindow(window, SW_SHOWNOACTIVATE);
    return true;
}

void PreviewOverlay::Hide(bool destroyWindow) {
    if (!m_window || !IsWindow(m_window)) {
        m_visible = false;
        return;
    }

    ShowWindow(m_window, SW_HIDE);
    if (destroyWindow) {
        DestroyWindow(m_window);
        m_window = nullptr;
    }
    m_visible = false;
}

void PreviewOverlay::Destroy() {
    Hide(true);
    m_owner = nullptr;
    m_bitmapSize = {};
}

void PreviewOverlay::PositionRelativeToRect(const RECT& anchorRectScreen, const POINT& cursorScreenPt) {
    if (!m_window || !IsWindow(m_window)) {
        return;
    }

    const int width = m_bitmapSize.cx;
    const int height = m_bitmapSize.cy;
    int x = anchorRectScreen.left + ((anchorRectScreen.right - anchorRectScreen.left) - width) / 2;
    int y = anchorRectScreen.bottom + 8;

    HMONITOR monitor = MonitorFromPoint(cursorScreenPt, MONITOR_DEFAULTTONEAREST);
    MONITORINFO monitorInfo{sizeof(monitorInfo)};
    if (GetMonitorInfoW(monitor, &monitorInfo)) {
        x = std::clamp<int>(x, monitorInfo.rcWork.left + 4, monitorInfo.rcWork.right - width - 4);
        if (y + height > monitorInfo.rcWork.bottom) {
            y = anchorRectScreen.top - height - 8;
        }
        if (y < monitorInfo.rcWork.top + 4) {
            y = monitorInfo.rcWork.top + 4;
        }
    }

    SetWindowPos(m_window, HWND_TOPMOST, x, y, width, height, SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_SHOWWINDOW);
}

}  // namespace shelltabs

