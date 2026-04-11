#include "TaskbarTabListPopup.h"

#include "ExplorerThemeUtils.h"
#include "Logging.h"
#include "Module.h"

#include <windowsx.h>
#include <shellapi.h>
#include <dwmapi.h>

#pragma comment(lib, "gdiplus.lib")

namespace shelltabs {

ATOM TaskbarTabListPopup::s_windowClass = 0;

static ULONG_PTR s_gdiplusToken = 0;

static void EnsureGdiplusInitialized() {
    if (s_gdiplusToken) return;
    Gdiplus::GdiplusStartupInput input;
    Gdiplus::GdiplusStartup(&s_gdiplusToken, &input, nullptr);
}

TaskbarTabListPopup::TaskbarTabListPopup() = default;

TaskbarTabListPopup::~TaskbarTabListPopup() {
    Destroy();
}

void TaskbarTabListPopup::Create() {
    if (m_hwnd) return;

    EnsureGdiplusInitialized();

    HINSTANCE hInst = GetModuleHandleInstance();

    if (!s_windowClass) {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = WndProc;
        wc.hInstance = hInst;
        wc.lpszClassName = L"ShellTabsTaskbarPopup";
        wc.cbWndExtra = sizeof(void*);
        s_windowClass = RegisterClassExW(&wc);
        if (!s_windowClass) {
            LogLastError(L"TaskbarTabListPopup::RegisterClass", GetLastError());
            return;
        }
    }

    m_hwnd = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_LAYERED,
        L"ShellTabsTaskbarPopup",
        nullptr,
        WS_POPUP,
        0, 0, 1, 1,
        nullptr, nullptr, hInst, this);

    if (!m_hwnd) {
        LogLastError(L"TaskbarTabListPopup::CreateWindow", GetLastError());
    }
}

void TaskbarTabListPopup::Destroy() {
    if (m_hwnd) {
        DestroyWindow(m_hwnd);
        m_hwnd = nullptr;
    }
}

void TaskbarTabListPopup::SetTabActivatedCallback(TabActivatedCallback cb) {
    m_callback = std::move(cb);
}

void TaskbarTabListPopup::ShowForWindow(
    const TaskbarTabSnapshot& snapshot, const RECT& thumbnailRect, HWND taskbarHwnd) {
    if (!m_hwnd) return;

    m_snapshot = snapshot;
    m_explorerHwnd = snapshot.explorerHwnd;
    m_hoveredIndex = -1;
    m_isDarkMode = IsAppDarkModePreferred();

    // Get DPI from the monitor containing the thumbnail
    HMONITOR mon = MonitorFromRect(&thumbnailRect, MONITOR_DEFAULTTONEAREST);
    UINT dpiX = 96, dpiY = 96;
    // Try GetDpiForMonitor if available (Windows 8.1+)
    using GetDpiForMonitorFn = HRESULT(WINAPI*)(HMONITOR, int, UINT*, UINT*);
    static auto pGetDpiForMonitor = reinterpret_cast<GetDpiForMonitorFn>(
        GetProcAddress(GetModuleHandleW(L"shcore.dll"), "GetDpiForMonitor"));
    if (pGetDpiForMonitor) {
        pGetDpiForMonitor(mon, 0 /*MDT_EFFECTIVE_DPI*/, &dpiX, &dpiY);
    }
    m_dpi = static_cast<int>(dpiX);

    int itemCount = static_cast<int>(m_snapshot.tabs.size());
    if (itemCount > kMaxVisibleItems) itemCount = kMaxVisibleItems;

    // Build hit rects
    m_hitRects.clear();
    float y = static_cast<float>(PopupPadding());
    int popW = PopupWidth();
    for (int i = 0; i < itemCount; i++) {
        TabHitRect hr;
        hr.bounds = Gdiplus::RectF(
            static_cast<float>(PopupPadding()), y,
            static_cast<float>(popW - 2 * PopupPadding()),
            static_cast<float>(ItemHeight()));
        hr.groupIndex = m_snapshot.tabs[i].groupIndex;
        hr.tabIndex = m_snapshot.tabs[i].tabIndex;
        m_hitRects.push_back(hr);
        y += ItemHeight();
    }

    RECT rc = CalculatePopupRect(thumbnailRect, taskbarHwnd, itemCount);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;

    SetWindowPos(m_hwnd, HWND_TOPMOST, rc.left, rc.top, w, h,
                 SWP_NOACTIVATE);
    PaintToLayered();
    ShowWindow(m_hwnd, SW_SHOWNOACTIVATE);
}

void TaskbarTabListPopup::Hide() {
    if (m_hwnd && IsWindowVisible(m_hwnd)) {
        ShowWindow(m_hwnd, SW_HIDE);
    }
    m_hoveredIndex = -1;
    m_snapshot = {};
    m_hitRects.clear();
    m_explorerHwnd = nullptr;
}

bool TaskbarTabListPopup::IsVisible() const {
    return m_hwnd && IsWindowVisible(m_hwnd);
}

LRESULT CALLBACK TaskbarTabListPopup::WndProc(
    HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    TaskbarTabListPopup* self = nullptr;

    if (msg == WM_NCCREATE) {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        self = static_cast<TaskbarTabListPopup*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, 0, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<TaskbarTabListPopup*>(GetWindowLongPtrW(hwnd, 0));
    }

    if (!self) return DefWindowProcW(hwnd, msg, wParam, lParam);

    switch (msg) {
        case WM_MOUSEMOVE: {
            POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            self->OnMouseMove(pt);
            return 0;
        }
        case WM_MOUSELEAVE: {
            self->OnMouseLeave();
            return 0;
        }
        case WM_LBUTTONDOWN: {
            POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            self->OnLButtonDown(pt);
            return 0;
        }
        case WM_NCHITTEST: {
            // Allow mouse messages to reach the window
            return HTCLIENT;
        }
        default:
            break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void TaskbarTabListPopup::OnPaint() {
    // Not used — we paint via UpdateLayeredWindow
}

void TaskbarTabListPopup::OnMouseMove(POINT clientPt) {
    if (!m_trackingMouse) {
        TRACKMOUSEEVENT tme{};
        tme.cbSize = sizeof(tme);
        tme.dwFlags = TME_LEAVE;
        tme.hwndTrack = m_hwnd;
        TrackMouseEvent(&tme);
        m_trackingMouse = true;
    }

    int newHovered = -1;
    Gdiplus::PointF pt(static_cast<float>(clientPt.x), static_cast<float>(clientPt.y));
    for (int i = 0; i < static_cast<int>(m_hitRects.size()); i++) {
        if (m_hitRects[i].bounds.Contains(pt)) {
            newHovered = i;
            break;
        }
    }

    if (newHovered != m_hoveredIndex) {
        m_hoveredIndex = newHovered;
        PaintToLayered();
    }
}

void TaskbarTabListPopup::OnMouseLeave() {
    m_trackingMouse = false;
    if (m_hoveredIndex != -1) {
        m_hoveredIndex = -1;
        PaintToLayered();
    }
}

void TaskbarTabListPopup::OnLButtonDown(POINT clientPt) {
    Gdiplus::PointF pt(static_cast<float>(clientPt.x), static_cast<float>(clientPt.y));
    for (int i = 0; i < static_cast<int>(m_hitRects.size()); i++) {
        if (m_hitRects[i].bounds.Contains(pt)) {
            if (m_callback && m_explorerHwnd) {
                m_callback(m_explorerHwnd, m_hitRects[i].groupIndex, m_hitRects[i].tabIndex);
            }
            return;
        }
    }
}

void TaskbarTabListPopup::PaintToLayered() {
    if (!m_hwnd) return;

    RECT rc;
    GetWindowRect(m_hwnd, &rc);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    if (w <= 0 || h <= 0) return;

    HDC screenDC = GetDC(nullptr);
    HDC memDC = CreateCompatibleDC(screenDC);

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = w;
    bmi.bmiHeader.biHeight = -h;  // top-down
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    void* bits = nullptr;
    HBITMAP dib = CreateDIBSection(memDC, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!dib) {
        DeleteDC(memDC);
        ReleaseDC(nullptr, screenDC);
        return;
    }
    HGDIOBJ oldBmp = SelectObject(memDC, dib);

    // Clear to transparent
    memset(bits, 0, static_cast<size_t>(w) * static_cast<size_t>(h) * 4);

    {
        Gdiplus::Graphics g(memDC);
        g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

        // Background with rounded corners
        Gdiplus::GraphicsPath path;
        float cr = static_cast<float>(CornerRadius());
        float fw = static_cast<float>(w);
        float fh = static_cast<float>(h);

        path.AddArc(0.0f, 0.0f, cr * 2.0f, cr * 2.0f, 180.0f, 90.0f);
        path.AddArc(fw - cr * 2.0f, 0.0f, cr * 2.0f, cr * 2.0f, 270.0f, 90.0f);
        path.AddArc(fw - cr * 2.0f, fh - cr * 2.0f, cr * 2.0f, cr * 2.0f, 0.0f, 90.0f);
        path.AddArc(0.0f, fh - cr * 2.0f, cr * 2.0f, cr * 2.0f, 90.0f, 90.0f);
        path.CloseFigure();

        Gdiplus::SolidBrush bgBrush(GetBackgroundColor());
        g.FillPath(&bgBrush, &path);

        // Draw 1px border
        Gdiplus::Color borderColor = m_isDarkMode
            ? Gdiplus::Color(80, 255, 255, 255)
            : Gdiplus::Color(40, 0, 0, 0);
        Gdiplus::Pen borderPen(borderColor, 1.0f);
        g.DrawPath(&borderPen, &path);

        // Draw each tab entry
        int itemCount = static_cast<int>(m_snapshot.tabs.size());
        if (itemCount > kMaxVisibleItems) itemCount = kMaxVisibleItems;

        for (int i = 0; i < itemCount; i++) {
            DrawTabEntry(g, i, m_snapshot.tabs[i], m_hitRects[i].bounds,
                         i == m_hoveredIndex, m_snapshot.tabs[i].isSelected);
        }
    }

    POINT ptSrc{0, 0};
    SIZE sz{w, h};
    POINT ptDst{rc.left, rc.top};
    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 255;
    blend.AlphaFormat = AC_SRC_ALPHA;
    UpdateLayeredWindow(m_hwnd, screenDC, &ptDst, &sz, memDC, &ptSrc, 0, &blend, ULW_ALPHA);

    SelectObject(memDC, oldBmp);
    DeleteObject(dib);
    DeleteDC(memDC);
    ReleaseDC(nullptr, screenDC);
}

void TaskbarTabListPopup::DrawTabEntry(
    Gdiplus::Graphics& g, int /*index*/, const TaskbarTabEntry& entry,
    const Gdiplus::RectF& bounds, bool isHovered, bool isSelected) {
    float itemCr = static_cast<float>(MulDiv(4, m_dpi, 96));

    // Hover/selected background
    if (isHovered || isSelected) {
        Gdiplus::GraphicsPath itemPath;
        itemPath.AddArc(bounds.X, bounds.Y, itemCr * 2, itemCr * 2, 180, 90);
        itemPath.AddArc(bounds.GetRight() - itemCr * 2, bounds.Y, itemCr * 2, itemCr * 2, 270, 90);
        itemPath.AddArc(bounds.GetRight() - itemCr * 2, bounds.GetBottom() - itemCr * 2,
                        itemCr * 2, itemCr * 2, 0, 90);
        itemPath.AddArc(bounds.X, bounds.GetBottom() - itemCr * 2, itemCr * 2, itemCr * 2, 90, 90);
        itemPath.CloseFigure();

        Gdiplus::SolidBrush fillBrush(isHovered ? GetHoverColor() : GetSelectedColor());
        g.FillPath(&fillBrush, &itemPath);
    }

    // Accent bar for selected tab
    if (isSelected) {
        float barW = static_cast<float>(AccentBarWidth());
        float barH = bounds.Height * 0.5f;
        float barY = bounds.Y + (bounds.Height - barH) / 2.0f;
        Gdiplus::RectF barRect(bounds.X + 2.0f, barY, barW, barH);

        Gdiplus::GraphicsPath barPath;
        float barCr = barW / 2.0f;
        barPath.AddArc(barRect.X, barRect.Y, barCr * 2, barCr * 2, 180, 90);
        barPath.AddArc(barRect.GetRight() - barCr * 2, barRect.Y, barCr * 2, barCr * 2, 270, 90);
        barPath.AddArc(barRect.GetRight() - barCr * 2, barRect.GetBottom() - barCr * 2,
                       barCr * 2, barCr * 2, 0, 90);
        barPath.AddArc(barRect.X, barRect.GetBottom() - barCr * 2, barCr * 2, barCr * 2, 90, 90);
        barPath.CloseFigure();

        Gdiplus::SolidBrush accentBrush(GetAccentColor());
        g.FillPath(&accentBrush, &barPath);
    }

    float padX = static_cast<float>(ItemPaddingX());
    float iconSz = static_cast<float>(IconSize());

    // Icon
    float iconX = bounds.X + padX + (isSelected ? static_cast<float>(AccentBarWidth()) + 4.0f : 0.0f);
    float iconY = bounds.Y + (bounds.Height - iconSz) / 2.0f;

    // Get shell icon for this path
    if (!entry.path.empty()) {
        SHFILEINFOW sfi{};
        DWORD_PTR result = SHGetFileInfoW(
            entry.path.c_str(), 0, &sfi, sizeof(sfi),
            SHGFI_ICON | SHGFI_SMALLICON);
        if (result && sfi.hIcon) {
            Gdiplus::Bitmap iconBmp(sfi.hIcon);
            g.DrawImage(&iconBmp,
                        Gdiplus::RectF(iconX, iconY, iconSz, iconSz),
                        0.0f, 0.0f,
                        static_cast<float>(iconBmp.GetWidth()),
                        static_cast<float>(iconBmp.GetHeight()),
                        Gdiplus::UnitPixel);
            DestroyIcon(sfi.hIcon);
        }
    }

    // Text
    float textX = iconX + iconSz + padX;
    float textW = bounds.GetRight() - textX - padX;
    if (entry.isPinned) {
        textW -= iconSz;  // Reserve space for pin indicator
    }

    Gdiplus::RectF textRect(textX, bounds.Y, textW, bounds.Height);

    Gdiplus::StringFormat fmt;
    fmt.SetAlignment(Gdiplus::StringAlignmentNear);
    fmt.SetLineAlignment(Gdiplus::StringAlignmentCenter);
    fmt.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
    fmt.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);

    int fontSize = MulDiv(12, m_dpi, 96);
    Gdiplus::Font font(L"Segoe UI", static_cast<float>(fontSize),
                       Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
    Gdiplus::SolidBrush textBrush(GetTextColor());

    g.DrawString(entry.name.c_str(), -1, &font, textRect, &fmt, &textBrush);

    // Pinned indicator
    if (entry.isPinned) {
        float pinX = bounds.GetRight() - padX - iconSz;
        float pinY = bounds.Y + (bounds.Height - iconSz) / 2.0f;
        Gdiplus::RectF pinRect(pinX, pinY, iconSz, iconSz);

        Gdiplus::Color pinColor = m_isDarkMode
            ? Gdiplus::Color(120, 255, 255, 255)
            : Gdiplus::Color(120, 0, 0, 0);
        Gdiplus::SolidBrush pinBrush(pinColor);

        int pinFontSize = MulDiv(10, m_dpi, 96);
        Gdiplus::Font pinFont(L"Segoe UI Symbol", static_cast<float>(pinFontSize),
                              Gdiplus::FontStyleRegular, Gdiplus::UnitPixel);
        Gdiplus::StringFormat pinFmt;
        pinFmt.SetAlignment(Gdiplus::StringAlignmentCenter);
        pinFmt.SetLineAlignment(Gdiplus::StringAlignmentCenter);
        g.DrawString(L"\xD83D\xDCCC", -1, &pinFont, pinRect, &pinFmt, &pinBrush);
    }
}

Gdiplus::Color TaskbarTabListPopup::GetBackgroundColor() const {
    return m_isDarkMode
        ? Gdiplus::Color(230, 0x2B, 0x2B, 0x2B)
        : Gdiplus::Color(240, 0xFF, 0xFF, 0xFF);
}

Gdiplus::Color TaskbarTabListPopup::GetTextColor() const {
    return m_isDarkMode
        ? Gdiplus::Color(255, 0xFF, 0xFF, 0xFF)
        : Gdiplus::Color(255, 0x1B, 0x1B, 0x1B);
}

Gdiplus::Color TaskbarTabListPopup::GetSelectedColor() const {
    return m_isDarkMode
        ? Gdiplus::Color(255, 0x38, 0x38, 0x38)
        : Gdiplus::Color(255, 0xEC, 0xF0, 0xF1);
}

Gdiplus::Color TaskbarTabListPopup::GetHoverColor() const {
    return m_isDarkMode
        ? Gdiplus::Color(255, 0x3E, 0x3E, 0x3E)
        : Gdiplus::Color(255, 0xE5, 0xE5, 0xE5);
}

Gdiplus::Color TaskbarTabListPopup::GetAccentColor() const {
    return Gdiplus::Color(255, 0x00, 0x78, 0xD4);
}

RECT TaskbarTabListPopup::CalculatePopupRect(
    const RECT& thumbRect, HWND taskbarHwnd, int itemCount) const {
    int edge = GetTaskbarEdge(taskbarHwnd);
    int popW = PopupWidth();
    int popH = std::min(itemCount, kMaxVisibleItems) * ItemHeight() + 2 * PopupPadding();
    int gap = MulDiv(4, m_dpi, 96);

    RECT rc{};
    if (edge == ABE_BOTTOM) {
        rc.left = thumbRect.right + gap;
        rc.top = thumbRect.bottom - popH;
        rc.right = rc.left + popW;
        rc.bottom = thumbRect.bottom;
    } else if (edge == ABE_TOP) {
        rc.left = thumbRect.right + gap;
        rc.top = thumbRect.top;
        rc.right = rc.left + popW;
        rc.bottom = rc.top + popH;
    } else if (edge == ABE_LEFT) {
        rc.left = thumbRect.right + gap;
        rc.top = thumbRect.top;
        rc.right = rc.left + popW;
        rc.bottom = rc.top + popH;
    } else {
        // ABE_RIGHT
        rc.left = thumbRect.left - gap - popW;
        rc.top = thumbRect.top;
        rc.right = thumbRect.left - gap;
        rc.bottom = rc.top + popH;
    }

    // Clamp to monitor work area
    HMONITOR mon = MonitorFromRect(&thumbRect, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    GetMonitorInfoW(mon, &mi);

    if (rc.right > mi.rcWork.right) OffsetRect(&rc, mi.rcWork.right - rc.right, 0);
    if (rc.left < mi.rcWork.left) OffsetRect(&rc, mi.rcWork.left - rc.left, 0);
    if (rc.bottom > mi.rcWork.bottom) OffsetRect(&rc, 0, mi.rcWork.bottom - rc.bottom);
    if (rc.top < mi.rcWork.top) OffsetRect(&rc, 0, mi.rcWork.top - rc.top);

    return rc;
}

int TaskbarTabListPopup::GetTaskbarEdge(HWND taskbarHwnd) const {
    APPBARDATA abd{};
    abd.cbSize = sizeof(abd);
    if (taskbarHwnd) {
        abd.hWnd = taskbarHwnd;
    } else {
        abd.hWnd = FindWindowW(L"Shell_TrayWnd", nullptr);
    }
    SHAppBarMessage(ABM_GETTASKBARPOS, &abd);
    return static_cast<int>(abd.uEdge);
}

int TaskbarTabListPopup::ItemHeight() const { return MulDiv(kBaseItemHeight, m_dpi, 96); }
int TaskbarTabListPopup::ItemPaddingX() const { return MulDiv(kBaseItemPadX, m_dpi, 96); }
int TaskbarTabListPopup::IconSize() const { return MulDiv(kBaseIconSize, m_dpi, 96); }
int TaskbarTabListPopup::PopupWidth() const { return MulDiv(kBasePopupWidth, m_dpi, 96); }
int TaskbarTabListPopup::PopupPadding() const { return MulDiv(kBasePopupPadding, m_dpi, 96); }
int TaskbarTabListPopup::CornerRadius() const { return MulDiv(kBaseCornerRadius, m_dpi, 96); }
int TaskbarTabListPopup::AccentBarWidth() const { return MulDiv(kBaseAccentBar, m_dpi, 96); }

}  // namespace shelltabs
