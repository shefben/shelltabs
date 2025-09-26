#include "TabBandWindow.h"

#include <algorithm>
#include <cmath>
#include <string>
#include <vector>
#include <limits>

#include <windowsx.h>
#include <shellapi.h>
#include <ShlObj.h>

#include "Module.h"
#include "TabBand.h"

namespace shelltabs {

namespace {
const wchar_t kWindowClassName[] = L"ShellTabsBandWindow";
constexpr int kButtonWidth = 28;
constexpr int kItemMinWidth = 60;
constexpr int kGroupMinWidth = 90;
constexpr int kGroupGap = 16;
constexpr int kTabGap = 6;
constexpr int kPaddingX = 12;
constexpr int kGroupPaddingX = 16;
constexpr int kDragThreshold = 4;
constexpr int kBadgePaddingX = 8;
constexpr int kBadgePaddingY = 2;
constexpr int kBadgeHeight = 18;
constexpr int kSplitIndicatorWidth = 14;
constexpr double kTagLightenFactor = 0.35;
constexpr int kTabCornerRadius = 8;
constexpr int kGroupCornerRadius = 10;
constexpr int kGroupOutlineThickness = 2;
constexpr int kIconGap = 6;
constexpr int kIslandIndicatorWidth = 3;
constexpr int kIndicatorHitPadding = 6;
constexpr int kCollapsedIndicatorPadding = 10;

COLORREF GetGroupColor(bool selected) {
    return selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_BTNFACE);
}

COLORREF GetTabColor(bool selected) {
    return selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_WINDOW);
}

COLORREF GetTabTextColor(bool selected) {
    return selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_WINDOWTEXT);
}

COLORREF LightenColor(COLORREF color, double factor) {
    factor = std::clamp(factor, 0.0, 1.0);
    const int r = static_cast<int>(GetRValue(color) + (255 - GetRValue(color)) * factor);
    const int g = static_cast<int>(GetGValue(color) + (255 - GetGValue(color)) * factor);
    const int b = static_cast<int>(GetBValue(color) + (255 - GetBValue(color)) * factor);
    return RGB(std::clamp(r, 0, 255), std::clamp(g, 0, 255), std::clamp(b, 0, 255));
}

COLORREF DarkenColor(COLORREF color, double factor) {
    factor = std::clamp(factor, 0.0, 1.0);
    const int r = static_cast<int>(GetRValue(color) * (1.0 - factor));
    const int g = static_cast<int>(GetGValue(color) * (1.0 - factor));
    const int b = static_cast<int>(GetBValue(color) * (1.0 - factor));
    return RGB(std::clamp(r, 0, 255), std::clamp(g, 0, 255), std::clamp(b, 0, 255));
}

double ComputeLuminance(COLORREF color) {
    const double r = GetRValue(color) / 255.0;
    const double g = GetGValue(color) / 255.0;
    const double b = GetBValue(color) / 255.0;
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

HFONT GetDefaultFont() {
    return static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
}

}  // namespace

TabBandWindow::TabBandWindow(TabBand* owner) : m_owner(owner) {}

TabBandWindow::~TabBandWindow() { Destroy(); }

HWND TabBandWindow::Create(HWND parent) {
    if (m_hwnd) {
        return m_hwnd;
    }

    WNDCLASSW wc{};
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = TabBandWindow::WndProc;
    wc.hInstance = GetModuleHandleInstance();
    wc.lpszClassName = kWindowClassName;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

    static ATOM atom = RegisterClassW(&wc);
    if (!atom && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        return nullptr;
    }

    m_hwnd = CreateWindowExW(0, kWindowClassName, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP,
                             0, 0, 0, 0, parent, nullptr, GetModuleHandleInstance(), this);

    return m_hwnd;
}

void TabBandWindow::Destroy() {
    CancelDrag();
    ClearExplorerContext();
    ClearVisualItems();

    if (m_newTabButton) {
        DestroyWindow(m_newTabButton);
        m_newTabButton = nullptr;
    }
    if (m_hwnd) {
        DestroyWindow(m_hwnd);
        m_hwnd = nullptr;
    }
    m_tabData.clear();
}

void TabBandWindow::Show(bool show) {
    if (!m_hwnd) {
        return;
    }
    ShowWindow(m_hwnd, show ? SW_SHOW : SW_HIDE);
}

void TabBandWindow::SetTabs(const std::vector<TabViewItem>& items) {
    m_tabData = items;
    m_contextHit = {};
    ClearExplorerContext();
    RebuildLayout();
    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
}

bool TabBandWindow::HasFocus() const {
    if (!m_hwnd) {
        return false;
    }
    const HWND focus = GetFocus();
    return focus == m_hwnd;
}

void TabBandWindow::FocusTab() {
    if (m_hwnd) {
        SetFocus(m_hwnd);
    }
}

void TabBandWindow::Layout(int width, int height) {
    m_clientRect = {0, 0, width, height};

    const int buttonWidth = kButtonWidth;
    const int buttonX = (width > buttonWidth) ? (width - buttonWidth) : 0;
    const int buttonHeight = height;
    if (m_newTabButton) {
        MoveWindow(m_newTabButton, buttonX, 0, buttonWidth, buttonHeight, TRUE);
    }

    m_clientRect.right = std::max(0, buttonX);
    RebuildLayout();
    InvalidateRect(m_hwnd, nullptr, TRUE);
}

void TabBandWindow::RebuildLayout() {
    ClearVisualItems();
    if (!m_hwnd) {
        return;
    }

    RECT bounds = m_clientRect;
    const int availableWidth = bounds.right - bounds.left;
    if (availableWidth <= 0) {
        return;
    }

    HDC dc = GetDC(m_hwnd);
    if (!dc) {
        return;
    }
    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(dc, font));

    const int top = bounds.top + 2;
    const int bottom = bounds.bottom - 2;
    const int baseIconWidth = std::max(GetSystemMetrics(SM_CXSMICON), 16);
    const int baseIconHeight = std::max(GetSystemMetrics(SM_CYSMICON), 16);
    int x = bounds.left + 4;
    int currentGroup = -1;
    TabViewItem currentHeader{};
    bool hasHeader = false;
    bool expectFirstTab = false;

    for (const auto& item : m_tabData) {
        if (item.type == TabViewItemType::kGroupHeader) {
            VisualItem visual;
            visual.data = item;
            visual.firstInGroup = true;
            visual.collapsedPlaceholder = item.collapsed || item.visibleTabs == 0;

            if (currentGroup >= 0) {
                x += kGroupGap;
            }

            SIZE textSize{0, 0};
            if (!item.name.empty()) {
                GetTextExtentPoint32W(dc, item.name.c_str(), static_cast<int>(item.name.size()), &textSize);
            }

            int width = kIslandIndicatorWidth;
            if (visual.collapsedPlaceholder) {
                const int padding = kGroupPaddingX;
                width = textSize.cx + padding * 2;
                width = std::max(width, kGroupMinWidth);
            }

            const int remaining = bounds.right - x;
            if (remaining <= 0) {
                break;
            }
            width = std::min(width, remaining);

            visual.bounds = {x, top, x + width, bottom};
            m_items.emplace_back(std::move(visual));

            currentGroup = item.location.groupIndex;
            currentHeader = item;
            hasHeader = true;
            expectFirstTab = true;
            x += width;
            continue;
        }

        VisualItem visual;
        visual.data = item;

        if (currentGroup != item.location.groupIndex) {
            currentGroup = item.location.groupIndex;
            hasHeader = false;
            expectFirstTab = true;
            if (!m_items.empty()) {
                x += kGroupGap;
            }
        } else if (!expectFirstTab) {
            x += kTabGap;
        }

        if (expectFirstTab) {
            visual.firstInGroup = true;
            expectFirstTab = false;
        }
        visual.hasGroupHeader = hasHeader;
        if (hasHeader) {
            visual.groupHeader = currentHeader;
        }

        SIZE textSize{0, 0};
        if (!item.name.empty()) {
            GetTextExtentPoint32W(dc, item.name.c_str(), static_cast<int>(item.name.size()), &textSize);
        }

        int width = textSize.cx + kPaddingX * 2;
        width = std::max(width, kItemMinWidth);

        visual.badgeWidth = MeasureBadgeWidth(item, dc);
        width += visual.badgeWidth;
        visual.icon = LoadItemIcon(item);
        if (visual.icon) {
            visual.iconWidth = baseIconWidth;
            visual.iconHeight = baseIconHeight;
            ICONINFO iconInfo{};
            if (GetIconInfo(visual.icon, &iconInfo)) {
                BITMAP bitmap{};
                if (iconInfo.hbmColor &&
                    GetObject(iconInfo.hbmColor, sizeof(bitmap), &bitmap) == sizeof(bitmap)) {
                    visual.iconWidth = bitmap.bmWidth;
                    visual.iconHeight = bitmap.bmHeight;
                } else if (iconInfo.hbmMask &&
                           GetObject(iconInfo.hbmMask, sizeof(bitmap), &bitmap) == sizeof(bitmap)) {
                    visual.iconWidth = bitmap.bmWidth;
                    visual.iconHeight = bitmap.bmHeight / 2;
                }
                if (iconInfo.hbmColor) {
                    DeleteObject(iconInfo.hbmColor);
                }
                if (iconInfo.hbmMask) {
                    DeleteObject(iconInfo.hbmMask);
                }
            }
            if (visual.iconWidth <= 0) {
                visual.iconWidth = baseIconWidth;
            }
            if (visual.iconHeight <= 0) {
                visual.iconHeight = baseIconHeight;
            }
            width += visual.iconWidth + kIconGap;
        }

        const int remaining = bounds.right - x;
        if (remaining <= 0) {
            break;
        }

        if (width > remaining) {
            width = remaining;
        }
        if (width < 40) {
            width = std::min(remaining, 40);
        }

        visual.bounds = {x, top, x + width, bottom};
        m_items.emplace_back(std::move(visual));
        x += width;
    }

    if (oldFont) {
        SelectObject(dc, oldFont);
    }
    ReleaseDC(m_hwnd, dc);
}

void TabBandWindow::Draw(HDC dc) const {
    RECT fillRect = m_clientRect;
    HBRUSH background = CreateSolidBrush(GetSysColor(COLOR_WINDOW));
    FillRect(dc, &fillRect, background);
    DeleteObject(background);

    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(dc, font));
    SetBkMode(dc, TRANSPARENT);

    for (const auto& item : m_items) {
        if (item.data.type == TabViewItemType::kGroupHeader) {
            DrawGroupHeader(dc, item);
        } else {
            DrawTab(dc, item);
        }
    }

    DrawDropIndicator(dc);
    DrawDragVisual(dc);

    if (oldFont) {
        SelectObject(dc, oldFont);
    }
}

COLORREF TabBandWindow::ResolveTabBackground(const TabViewItem& item) const {
    if (item.selected) {
        return GetTabColor(true);
    }
    if (item.hasCustomOutline) {
        return LightenColor(item.outlineColor, 0.55);
    }
    if (item.hasTagColor) {
        return LightenColor(item.tagColor, kTagLightenFactor);
    }
    return GetTabColor(false);
}

COLORREF TabBandWindow::ResolveGroupBackground(const TabViewItem& item) const {
    if (item.selected) {
        return GetGroupColor(true);
    }
    if (item.hasCustomOutline) {
        return LightenColor(item.outlineColor, 0.45);
    }
    if (item.hasTagColor) {
        return LightenColor(item.tagColor, kTagLightenFactor);
    }
    return GetGroupColor(false);
}

COLORREF TabBandWindow::ResolveTextColor(COLORREF background) const {
    return ComputeLuminance(background) > 0.6 ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

std::wstring TabBandWindow::BuildGitBadgeText(const TabViewItem& item) const {
    if (!item.hasGitStatus) {
        return {};
    }
    std::wstring text = item.gitStatus.branch.empty() ? L"git" : item.gitStatus.branch;
    if (item.gitStatus.hasChanges) {
        text += L"*";
    }
    if (item.gitStatus.hasUntracked) {
        text += L"+";
    }
    if (item.gitStatus.ahead > 0) {
        text += L" ↑" + std::to_wstring(item.gitStatus.ahead);
    }
    if (item.gitStatus.behind > 0) {
        text += L" ↓" + std::to_wstring(item.gitStatus.behind);
    }
    return text;
}

int TabBandWindow::MeasureBadgeWidth(const TabViewItem& item, HDC dc) const {
    int width = 0;
    if (item.hasGitStatus) {
        const std::wstring badge = BuildGitBadgeText(item);
        if (!badge.empty()) {
            SIZE badgeSize{};
            GetTextExtentPoint32W(dc, badge.c_str(), static_cast<int>(badge.size()), &badgeSize);
            width += badgeSize.cx + kBadgePaddingX * 2;
        }
    }
    if (item.splitSecondary || item.splitPrimary) {
        width += kSplitIndicatorWidth;
    }
    return width;
}

void TabBandWindow::DrawGroupHeader(HDC dc, const VisualItem& item) const {
    RECT rect = item.bounds;
    if (!item.collapsedPlaceholder) {
        RECT indicator = rect;
        indicator.right = indicator.left + kIslandIndicatorWidth;
        COLORREF indicatorColor = item.data.hasCustomOutline
                                      ? item.data.outlineColor
                                      : (item.data.hasTagColor ? item.data.tagColor : GetSysColor(COLOR_HOTLIGHT));
        if (item.data.selected) {
            indicatorColor = DarkenColor(indicatorColor, 0.2);
        }
        HBRUSH brush = CreateSolidBrush(indicatorColor);
        if (brush) {
            FillRect(dc, &indicator, brush);
            DeleteObject(brush);
        }
        return;
    }

    const bool selected = item.data.selected;
    COLORREF backgroundColor = ResolveGroupBackground(item.data);
    COLORREF textColor = selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : ResolveTextColor(backgroundColor);
    COLORREF outlineColor = item.data.hasCustomOutline
                                ? (selected ? DarkenColor(item.data.outlineColor, 0.25) : item.data.outlineColor)
                                : (item.data.hasTagColor ? DarkenColor(item.data.tagColor, 0.25)
                                                         : (selected ? DarkenColor(GetGroupColor(true), 0.2)
                                                                     : GetSysColor(COLOR_HOTLIGHT)));

    HBRUSH brush = CreateSolidBrush(backgroundColor);
    FillRect(dc, &rect, brush);
    DeleteObject(brush);

    RECT outline = rect;
    InflateRect(&outline, -1, -1);
    if (outline.right > outline.left && outline.bottom > outline.top) {
        HPEN pen = CreatePen(PS_SOLID, kGroupOutlineThickness, outlineColor);
        if (pen) {
            HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
            HBRUSH oldBrush = static_cast<HBRUSH>(SelectObject(dc, GetStockObject(HOLLOW_BRUSH)));
            RoundRect(dc, outline.left, outline.top, outline.right, outline.bottom, kGroupCornerRadius,
                      kGroupCornerRadius);
            SelectObject(dc, oldBrush);
            SelectObject(dc, oldPen);
            DeleteObject(pen);
        }
    }

    RECT textRect = rect;
    textRect.left += kCollapsedIndicatorPadding;
    textRect.right -= kCollapsedIndicatorPadding;
    SetTextColor(dc, textColor);
    DrawTextW(dc, item.data.name.c_str(), static_cast<int>(item.data.name.size()), &textRect,
              DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
}

void TabBandWindow::DrawTab(HDC dc, const VisualItem& item) const {
    RECT rect = item.bounds;
    const bool selected = item.data.selected;
    COLORREF backgroundColor = ResolveTabBackground(item.data);
    COLORREF textColor = selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : ResolveTextColor(backgroundColor);
    COLORREF borderColor = selected ? GetSysColor(COLOR_WINDOWFRAME) : GetSysColor(COLOR_3DSHADOW);

    RECT shapeRect = rect;
    if (!selected) {
        shapeRect.bottom -= 1;
    }

    const int radius = kTabCornerRadius;
    POINT points[] = {{shapeRect.left, shapeRect.bottom},
                      {shapeRect.left, shapeRect.top + radius},
                      {shapeRect.left + radius, shapeRect.top},
                      {shapeRect.right - radius, shapeRect.top},
                      {shapeRect.right, shapeRect.top + radius},
                      {shapeRect.right, shapeRect.bottom}};

    HRGN region = CreatePolygonRgn(points, ARRAYSIZE(points), WINDING);
    if (region) {
        HBRUSH brush = CreateSolidBrush(backgroundColor);
        if (brush) {
            FillRgn(dc, region, brush);
            DeleteObject(brush);
        }
        HPEN pen = CreatePen(PS_SOLID, 1, borderColor);
        if (pen) {
            HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
            HBRUSH oldBrush = static_cast<HBRUSH>(SelectObject(dc, GetStockObject(HOLLOW_BRUSH)));
            Polygon(dc, points, ARRAYSIZE(points));
            SelectObject(dc, oldBrush);
            SelectObject(dc, oldPen);
            DeleteObject(pen);
        }
        DeleteObject(region);
    }

    COLORREF bottomLineColor = selected ? backgroundColor : GetSysColor(COLOR_3DLIGHT);
    HPEN bottomPen = CreatePen(PS_SOLID, 1, bottomLineColor);
    if (bottomPen) {
        HPEN oldPen = static_cast<HPEN>(SelectObject(dc, bottomPen));
        MoveToEx(dc, rect.left + 1, rect.bottom - 1, nullptr);
        LineTo(dc, rect.right - 1, rect.bottom - 1);
        SelectObject(dc, oldPen);
        DeleteObject(bottomPen);
    }

    const int indicatorWidth = (item.data.splitSecondary || item.data.splitPrimary) ? kSplitIndicatorWidth : 0;
    const int badgeWidth = item.badgeWidth - indicatorWidth;

    int textLeft = rect.left + kPaddingX;
    if (item.icon) {
        const int availableHeight = rect.bottom - rect.top;
        const int iconHeight = std::min(item.iconHeight, availableHeight - 4);
        const int iconWidth = item.iconWidth;
        const int iconY = rect.top + (availableHeight - iconHeight) / 2;
        DrawIconEx(dc, textLeft, iconY, item.icon, iconWidth, iconHeight, 0, nullptr, DI_NORMAL);
        textLeft += iconWidth + kIconGap;
    }

    RECT textRect = rect;
    textRect.left = textLeft;
    textRect.top += 3;
    textRect.right -= (item.badgeWidth > 0 ? item.badgeWidth : kPaddingX);

    SetTextColor(dc, textColor);
    DrawTextW(dc, item.data.name.c_str(), static_cast<int>(item.data.name.size()), &textRect,
              DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);

    if (badgeWidth > 0) {
        RECT badgeRect = rect;
        badgeRect.left = rect.right - indicatorWidth - badgeWidth + kBadgePaddingX;
        badgeRect.right = rect.right - indicatorWidth - kBadgePaddingX;
        badgeRect.top = rect.top + 4;
        badgeRect.bottom = badgeRect.top + kBadgeHeight;
        COLORREF badgeColor = item.data.gitStatus.hasChanges ? RGB(200, 130, 60) : RGB(90, 150, 90);
        HBRUSH badgeBrush = CreateSolidBrush(LightenColor(badgeColor, 0.15));
        if (badgeBrush) {
            RECT badgeFill = badgeRect;
            badgeFill.left -= kBadgePaddingX;
            badgeFill.right += kBadgePaddingX;
            FillRect(dc, &badgeFill, badgeBrush);
            DeleteObject(badgeBrush);
            HPEN badgePen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_3DSHADOW));
            if (badgePen) {
                HPEN oldPen = static_cast<HPEN>(SelectObject(dc, badgePen));
                MoveToEx(dc, badgeFill.left, badgeFill.top, nullptr);
                LineTo(dc, badgeFill.right, badgeFill.top);
                LineTo(dc, badgeFill.right, badgeFill.bottom);
                LineTo(dc, badgeFill.left, badgeFill.bottom);
                LineTo(dc, badgeFill.left, badgeFill.top);
                SelectObject(dc, oldPen);
                DeleteObject(badgePen);
            }
        }
        SetTextColor(dc, ResolveTextColor(badgeColor));
        const std::wstring badgeText = BuildGitBadgeText(item.data);
        DrawTextW(dc, badgeText.c_str(), static_cast<int>(badgeText.size()), &badgeRect,
                  DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX);
        SetTextColor(dc, textColor);
    }

    if (indicatorWidth > 0) {
        RECT indicator = rect;
        indicator.left = rect.right - indicatorWidth;
        indicator.top += 4;
        indicator.bottom -= 2;
        COLORREF indicatorColor = item.data.splitSecondary ? RGB(90, 140, 220) : RGB(120, 120, 120);
        HBRUSH indicatorBrush = CreateSolidBrush(indicatorColor);
        if (indicatorBrush) {
            FillRect(dc, &indicator, indicatorBrush);
            DeleteObject(indicatorBrush);
        }
        SetTextColor(dc, RGB(255, 255, 255));
        std::wstring marker = item.data.splitSecondary ? L"R" : L"L";
        DrawTextW(dc, marker.c_str(), static_cast<int>(marker.size()), &indicator,
                  DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX);
        SetTextColor(dc, textColor);
    }
}


void TabBandWindow::DrawDropIndicator(HDC dc) const {
    if (!m_drag.dragging || !m_drag.target.active || m_drag.target.outside || m_drag.target.indicatorX < 0) {
        return;
    }

    HPEN pen = CreatePen(PS_SOLID, 2, GetSysColor(COLOR_HIGHLIGHT));
    HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
    const int x = m_drag.target.indicatorX;
    MoveToEx(dc, x, m_clientRect.top + 2, nullptr);
    LineTo(dc, x, m_clientRect.bottom - 2);
    SelectObject(dc, oldPen);
    DeleteObject(pen);
}

void TabBandWindow::DrawDragVisual(HDC dc) const {
    if (!m_drag.dragging || !m_drag.origin.hit || !m_drag.hasCurrent) {
        return;
    }

    const VisualItem* originItem = FindVisualForHit(m_drag.origin);
    if (!originItem) {
        return;
    }

    const int width = originItem->bounds.right - originItem->bounds.left;
    const int height = originItem->bounds.bottom - originItem->bounds.top;
    if (width <= 0 || height <= 0) {
        return;
    }

    HDC memDC = CreateCompatibleDC(dc);
    if (!memDC) {
        return;
    }

    HBITMAP bitmap = CreateCompatibleBitmap(dc, width, height);
    if (!bitmap) {
        DeleteDC(memDC);
        return;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, bitmap);
    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(memDC, font));
    SetBkMode(memDC, TRANSPARENT);

    VisualItem ghost = *originItem;
    ghost.bounds = {0, 0, width, height};

    if (ghost.data.type == TabViewItemType::kGroupHeader) {
        DrawGroupHeader(memDC, ghost);
    } else {
        DrawTab(memDC, ghost);
    }

    if (oldFont) {
        SelectObject(memDC, oldFont);
    }

    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 160;
    blend.AlphaFormat = 0;

    const int left = m_drag.current.x - width / 2;
    const int top = m_drag.current.y - height / 2;
    AlphaBlend(dc, left, top, width, height, memDC, 0, 0, width, height, blend);

    SelectObject(memDC, oldBitmap);
    DeleteObject(bitmap);
    DeleteDC(memDC);
}

void TabBandWindow::ClearVisualItems() {
    for (auto& visual : m_items) {
        if (visual.icon) {
            DestroyIcon(visual.icon);
            visual.icon = nullptr;
        }
    }
    m_items.clear();
}

void TabBandWindow::ClearExplorerContext() {
    m_explorerContext = {};
}

HICON TabBandWindow::LoadItemIcon(const TabViewItem& item) const {
    if (item.type != TabViewItemType::kTab) {
        return nullptr;
    }

    SHFILEINFOW info{};
    UINT flags = SHGFI_ICON | SHGFI_SMALLICON | SHGFI_ADDOVERLAYS;
    if (item.pidl) {
        if (SHGetFileInfoW(reinterpret_cast<PCWSTR>(item.pidl), 0, &info, sizeof(info), flags | SHGFI_PIDL)) {
            return info.hIcon;
        }
    }
    if (!item.path.empty()) {
        if (SHGetFileInfoW(item.path.c_str(), 0, &info, sizeof(info), flags)) {
            return info.hIcon;
        }
    }
    return nullptr;
}

bool TabBandWindow::HandleExplorerMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, LRESULT* result) {
    if (m_explorerContext.menu3) {
        return SUCCEEDED(m_explorerContext.menu3->HandleMenuMsg2(message, wParam, lParam, result));
    }
    if (m_explorerContext.menu2) {
        HRESULT hr = m_explorerContext.menu2->HandleMenuMsg(message, wParam, lParam);
        if (SUCCEEDED(hr)) {
            if (result) {
                *result = 0;
            }
            return true;
        }
    }
    return false;
}

void TabBandWindow::HandleCommand(WPARAM wParam, LPARAM) {
    if (!m_owner) {
        return;
    }

    const UINT id = LOWORD(wParam);
    if (id == IDC_NEW_TAB) {
        m_owner->OnNewTabRequested();
        return;
    }

    if (id == IDM_CREATE_SAVED_GROUP) {
        const int insertAfter = ResolveInsertGroupIndex();
        m_owner->OnCreateSavedGroup(insertAfter);
        ClearExplorerContext();
        return;
    }

    if (id >= IDM_LOAD_SAVED_GROUP_BASE && id <= IDM_LOAD_SAVED_GROUP_LAST) {
        for (const auto& entry : m_savedGroupCommands) {
            if (entry.first == id) {
                const int insertAfter = ResolveInsertGroupIndex();
                m_owner->OnLoadSavedGroup(entry.second, insertAfter);
                break;
            }
        }
        ClearExplorerContext();
        return;
    }

    if (!m_contextHit.hit) {
        ClearExplorerContext();
        return;
    }

    if (id == IDM_CLOSE_TAB && m_contextHit.location.IsValid()) {
        m_owner->OnCloseTabRequested(m_contextHit.location);
    } else if (id == IDM_HIDE_TAB && m_contextHit.location.IsValid()) {
        m_owner->OnHideTabRequested(m_contextHit.location);
    } else if (id == IDM_DETACH_TAB && m_contextHit.location.IsValid()) {
        m_owner->OnDetachTabRequested(m_contextHit.location);
    } else if (id == IDM_OPEN_TERMINAL && m_contextHit.location.IsValid()) {
        m_owner->OnOpenTerminal(m_contextHit.location);
    } else if (id == IDM_OPEN_VSCODE && m_contextHit.location.IsValid()) {
        m_owner->OnOpenVSCode(m_contextHit.location);
    } else if (id == IDM_COPY_PATH && m_contextHit.location.IsValid()) {
        m_owner->OnCopyPath(m_contextHit.location);
    } else if (id == IDM_TOGGLE_ISLAND_HEADER && m_contextHit.location.groupIndex >= 0) {
        const bool visible = m_owner->IsGroupHeaderVisible(m_contextHit.location.groupIndex);
        m_owner->OnSetGroupHeaderVisible(m_contextHit.location.groupIndex, !visible);
    } else if (id == IDM_SET_SPLIT_SECONDARY && m_contextHit.location.IsValid()) {
        m_owner->OnPromoteSplitSecondary(m_contextHit.location);
    } else if (id == IDM_TOGGLE_ISLAND) {
        m_owner->OnToggleGroupCollapsed(m_contextHit.location.groupIndex);
    } else if (id == IDM_UNHIDE_ALL) {
        m_owner->OnUnhideAllInGroup(m_contextHit.location.groupIndex);
    } else if (id == IDM_NEW_ISLAND) {
        m_owner->OnCreateIslandAfter(m_contextHit.location.groupIndex);
    } else if (id == IDM_DETACH_ISLAND) {
        m_owner->OnDetachGroupRequested(m_contextHit.location.groupIndex);
    } else if (id == IDM_TOGGLE_SPLIT) {
        m_owner->OnToggleSplitView(m_contextHit.location.groupIndex);
    } else if (id == IDM_CLEAR_SPLIT_SECONDARY) {
        m_owner->OnClearSplitSecondary(m_contextHit.location.groupIndex);
    } else if (id == IDM_SWAP_SPLIT) {
        m_owner->OnSwapSplitPanes(m_contextHit.location.groupIndex);
    } else if (id >= IDM_HIDDEN_TAB_BASE) {
        for (const auto& entry : m_hiddenTabCommands) {
            if (entry.first == id) {
                m_owner->OnUnhideTabRequested(entry.second);
                break;
            }
        }
    } else if (m_explorerContext.menu && id >= m_explorerContext.idFirst &&
               id <= m_explorerContext.idLast) {
        m_owner->InvokeExplorerContextCommand(m_explorerContext.location,
                                              m_explorerContext.menu.Get(), id,
                                              m_explorerContext.idFirst, m_lastContextPoint);
    }

    ClearExplorerContext();
}

void TabBandWindow::HandleMouseDown(const POINT& pt) {
    SetFocus(m_hwnd);
    HitInfo hit = HitTest(pt);
    if (!hit.hit) {
        return;
    }
    m_drag = {};
    m_drag.tracking = true;
    m_drag.origin = hit;
    m_drag.start = pt;
    m_drag.current = pt;
    m_drag.hasCurrent = true;
}

void TabBandWindow::HandleMouseUp(const POINT& pt) {
    if (m_drag.dragging) {
        m_drag.current = pt;
        m_drag.hasCurrent = true;
        UpdateDropTarget(pt);
        CompleteDrop();
    } else if (m_drag.tracking) {
        HitInfo hit = HitTest(pt);
        if (hit.hit) {
            RequestSelection(hit);
        }
    }
    CancelDrag();
}

void TabBandWindow::HandleMouseMove(const POINT& pt) {
    if (!m_drag.tracking) {
        return;
    }

    m_drag.current = pt;
    m_drag.hasCurrent = true;

    if (!m_drag.dragging) {
        if (std::abs(pt.x - m_drag.start.x) > kDragThreshold || std::abs(pt.y - m_drag.start.y) > kDragThreshold) {
            m_drag.dragging = true;
            SetCapture(m_hwnd);
        }
    }

    if (m_drag.dragging) {
        UpdateDropTarget(pt);
    }
}

void TabBandWindow::HandleDoubleClick(const POINT& pt) {
    HitInfo hit = HitTest(pt);
    if (!hit.hit || !m_owner) {
        return;
    }

    if (hit.type == TabViewItemType::kGroupHeader) {
        m_owner->OnToggleGroupCollapsed(hit.location.groupIndex);
    } else if (hit.location.IsValid()) {
        m_owner->OnDetachTabRequested(hit.location);
    }
}

void TabBandWindow::HandleFileDrop(HDROP drop) {
    if (!drop || !m_owner) {
        return;
    }
    POINT pt{};
    BOOL inside = DragQueryPoint(drop, &pt);
    if (!inside) {
        DragFinish(drop);
        return;
    }

    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
    if (count == 0) {
        DragFinish(drop);
        return;
    }

    std::vector<std::wstring> paths;
    paths.reserve(count);
    wchar_t buffer[MAX_PATH];
    for (UINT i = 0; i < count; ++i) {
        const UINT length = DragQueryFileW(drop, i, buffer, ARRAYSIZE(buffer));
        if (length == 0) {
            continue;
        }
        paths.emplace_back(buffer, buffer + length);
    }

    HitInfo hit = HitTest(pt);
    if (hit.hit && hit.type == TabViewItemType::kTab && hit.location.IsValid() && !paths.empty()) {
        const bool move = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        m_owner->OnFilesDropped(hit.location, paths, move);
    }

    DragFinish(drop);
}

void TabBandWindow::CancelDrag() {
    if (m_drag.dragging) {
        ReleaseCapture();
    }
    m_drag = {};
    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
}

void TabBandWindow::UpdateDropTarget(const POINT& pt) {
    DropTarget target{};
    target.active = true;

    if (pt.x < m_clientRect.left || pt.x > m_clientRect.right || pt.y < m_clientRect.top || pt.y > m_clientRect.bottom) {
        target.outside = true;
        m_drag.target = target;
        InvalidateRect(m_hwnd, nullptr, TRUE);
        return;
    }

    HitInfo hit = HitTest(pt);
    if (!hit.hit) {
        if (m_drag.origin.type == TabViewItemType::kTab && m_owner) {
            target.group = false;
            target.newGroup = true;
            target.floating = true;
            target.groupIndex = m_owner->GetGroupCount();
            target.tabIndex = 0;
            target.indicatorX = m_clientRect.right - 10;
        } else if (!m_items.empty()) {
            const VisualItem* lastHeader = FindLastGroupHeader();
            if (lastHeader) {
                if (m_drag.origin.type == TabViewItemType::kGroupHeader) {
                    target.group = true;
                    target.groupIndex = lastHeader->data.location.groupIndex + 1;
                    target.indicatorX = lastHeader->bounds.right;
                } else {
                    target.group = false;
                    target.groupIndex = lastHeader->data.location.groupIndex;
                    target.tabIndex = static_cast<int>(lastHeader->data.totalTabs);
                    target.indicatorX = lastHeader->bounds.right;
                }
            } else {
                const auto& tail = m_items.back();
                target.group = false;
                target.groupIndex = tail.data.location.groupIndex;
                target.tabIndex = tail.data.location.tabIndex + 1;
                target.indicatorX = tail.bounds.right;
            }
        }
        m_drag.target = target;
        InvalidateRect(m_hwnd, nullptr, TRUE);
        return;
    }

    const VisualItem& visual = m_items[hit.itemIndex];
    const int midX = (visual.bounds.left + visual.bounds.right) / 2;
    const bool leftSide = pt.x < midX;

    if (m_drag.origin.type == TabViewItemType::kGroupHeader) {
        target.group = true;
        target.groupIndex = visual.data.location.groupIndex + (leftSide ? 0 : 1);
        target.indicatorX = leftSide ? visual.bounds.left : visual.bounds.right;
    } else {
        target.group = false;
        target.groupIndex = visual.data.location.groupIndex;
        if (visual.data.type == TabViewItemType::kTab) {
            target.tabIndex = visual.data.location.tabIndex + (leftSide ? 0 : 1);
            target.indicatorX = leftSide ? visual.bounds.left : visual.bounds.right;
        } else {
            target.tabIndex = leftSide ? 0 : static_cast<int>(visual.data.totalTabs);
            target.indicatorX = leftSide ? visual.bounds.left : visual.bounds.right;
        }
    }

    m_drag.target = target;
    InvalidateRect(m_hwnd, nullptr, TRUE);
}

void TabBandWindow::CompleteDrop() {
    if (!m_owner || !m_drag.dragging) {
        return;
    }

    const auto origin = m_drag.origin;
    const auto target = m_drag.target;

    if (!target.active) {
        return;
    }

    if (target.outside) {
        if (origin.type == TabViewItemType::kGroupHeader) {
            m_owner->OnDetachGroupRequested(origin.location.groupIndex);
        } else if (origin.location.IsValid()) {
            m_owner->OnDetachTabRequested(origin.location);
        }
        return;
    }

    if (target.newGroup && origin.location.IsValid()) {
        m_owner->OnMoveTabToNewGroup(origin.location, target.groupIndex, !target.floating);
        return;
    }

    if (origin.type == TabViewItemType::kGroupHeader) {
        int fromGroup = origin.location.groupIndex;
        int toGroup = target.groupIndex;
        const int groupCount = GroupCount();
        if (toGroup < 0) {
            toGroup = 0;
        }
        if (toGroup > groupCount) {
            toGroup = groupCount;
        }
        if (fromGroup != toGroup && fromGroup + 1 != toGroup) {
            m_owner->OnMoveGroupRequested(fromGroup, toGroup);
        }
    } else if (origin.location.IsValid()) {
        TabLocation to{target.groupIndex, target.tabIndex};
        if (to.groupIndex < 0) {
            to.groupIndex = origin.location.groupIndex;
        }
        if (to.tabIndex < 0) {
            to.tabIndex = origin.location.tabIndex;
        }
        if (origin.location.groupIndex == to.groupIndex) {
            if (origin.location.tabIndex == to.tabIndex || origin.location.tabIndex + 1 == to.tabIndex) {
                return;
            }
        }
        if (to.tabIndex < 0) {
            to.tabIndex = 0;
        }
        if (to.groupIndex < 0) {
            to.groupIndex = 0;
        }
        if (!(origin.location.groupIndex == to.groupIndex && origin.location.tabIndex == to.tabIndex)) {
            m_owner->OnMoveTabRequested(origin.location, to);
        }
    }
}

void TabBandWindow::RequestSelection(const HitInfo& hit) {
    if (!m_owner) {
        return;
    }

    if (hit.type == TabViewItemType::kTab && hit.location.IsValid()) {
        m_owner->OnTabSelected(hit.location);
    } else if (hit.type == TabViewItemType::kGroupHeader) {
        m_owner->OnToggleGroupCollapsed(hit.location.groupIndex);
    }
}

TabBandWindow::HitInfo TabBandWindow::HitTest(const POINT& pt) const {
    HitInfo info;
    if (pt.x < m_clientRect.left || pt.x > m_clientRect.right || pt.y < m_clientRect.top || pt.y > m_clientRect.bottom) {
        return info;
    }

    for (size_t i = 0; i < m_items.size(); ++i) {
        const auto& item = m_items[i];
        if (PtInRect(&item.bounds, pt)) {
            info.hit = true;
            info.itemIndex = i;
            info.type = item.data.type;
            info.location = item.data.location;
            const int midX = (item.bounds.left + item.bounds.right) / 2;
            info.before = pt.x < midX;
            info.after = !info.before;
            return info;
        }
    }

    return info;
}

void TabBandWindow::ShowContextMenu(const POINT& screenPt) {
    if (!m_owner) {
        return;
    }

    POINT clientPt = screenPt;
    if (screenPt.x == -1 && screenPt.y == -1) {
        clientPt.x = m_clientRect.left + 10;
        clientPt.y = m_clientRect.top + 10;
        ClientToScreen(m_hwnd, &clientPt);
    }

    POINT pt = clientPt;
    ScreenToClient(m_hwnd, &pt);
    HitInfo hit = HitTest(pt);
    ClearExplorerContext();
    m_savedGroupCommands.clear();
    m_contextHit = hit;
    m_lastContextPoint = clientPt;

    HMENU menu = CreatePopupMenu();
    if (!menu) {
        return;
    }

    m_hiddenTabCommands.clear();

    bool hasItemCommands = false;

    if (hit.hit) {
        if (hit.type == TabViewItemType::kTab) {
            AppendMenuW(menu, MF_STRING, IDM_CLOSE_TAB, L"Close Tab");
            AppendMenuW(menu, MF_STRING, IDM_HIDE_TAB, L"Hide Tab");
            AppendMenuW(menu, MF_STRING, IDM_DETACH_TAB, L"Move to New Window");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

            const bool headerVisible = m_owner->IsGroupHeaderVisible(hit.location.groupIndex);
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND_HEADER,
                        headerVisible ? L"Hide Island Indicator" : L"Show Island Indicator");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

            AppendMenuW(menu, MF_STRING, IDM_OPEN_TERMINAL, L"Open Terminal Here");
            AppendMenuW(menu, MF_STRING, IDM_OPEN_VSCODE, L"Open in VS Code");
            AppendMenuW(menu, MF_STRING, IDM_COPY_PATH, L"Copy Path");

            HMENU explorerMenu = CreatePopupMenu();
            bool explorerInserted = false;
            if (explorerMenu) {
                Microsoft::WRL::ComPtr<IContextMenu> cmenu;
                Microsoft::WRL::ComPtr<IContextMenu2> cmenu2;
                Microsoft::WRL::ComPtr<IContextMenu3> cmenu3;
                UINT usedLast = 0;
                if (m_owner->BuildExplorerContextMenu(hit.location, explorerMenu, IDM_EXPLORER_CONTEXT_BASE,
                                                      IDM_EXPLORER_CONTEXT_LAST, &cmenu, &cmenu2, &cmenu3,
                                                      &usedLast)) {
                    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                    AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(explorerMenu), L"Explorer Context");
                    m_explorerContext.menu = std::move(cmenu);
                    m_explorerContext.menu2 = std::move(cmenu2);
                    m_explorerContext.menu3 = std::move(cmenu3);
                    m_explorerContext.idFirst = IDM_EXPLORER_CONTEXT_BASE;
                    m_explorerContext.idLast = usedLast;
                    m_explorerContext.location = hit.location;
                    explorerInserted = true;
                } else {
                    DestroyMenu(explorerMenu);
                }
            }
            if (!explorerInserted) {
                AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, L"Explorer Context");
            }

            const auto& data = m_items[hit.itemIndex].data;
            if (data.splitAvailable) {
                AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
                if (data.splitSecondary) {
                    AppendMenuW(menu, MF_STRING, IDM_CLEAR_SPLIT_SECONDARY, L"Remove Split Companion");
                } else {
                    AppendMenuW(menu, MF_STRING, IDM_SET_SPLIT_SECONDARY,
                                data.splitEnabled ? L"Set as Split Companion" : L"Enable Split View with This Tab");
                }
            }
            hasItemCommands = true;
        } else {
            const auto& item = m_items[hit.itemIndex];
            const bool collapsed = item.data.collapsed;
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND, collapsed ? L"Show Island" : L"Hide Island");
            const bool headerVisible = m_owner->IsGroupHeaderVisible(item.data.location.groupIndex);
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND_HEADER,
                        headerVisible ? L"Hide Island Indicator" : L"Show Island Indicator");

            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_SPLIT,
                        item.data.splitEnabled ? L"Disable Split View" : L"Enable Split View");
            if (item.data.splitEnabled) {
                AppendMenuW(menu, MF_STRING, IDM_SWAP_SPLIT, L"Swap Split Panes");
                AppendMenuW(menu, MF_STRING, IDM_CLEAR_SPLIT_SECONDARY, L"Clear Split Companion");
            }

            if (item.data.hiddenTabs > 0) {
                HMENU hiddenMenu = CreatePopupMenu();
                PopulateHiddenTabsMenu(hiddenMenu, item.data.location.groupIndex);
                AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(hiddenMenu), L"Hidden Tabs");
                AppendMenuW(menu, MF_STRING, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            } else {
                AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            }

            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, IDM_NEW_ISLAND, L"New Island After");
            AppendMenuW(menu, MF_STRING, IDM_DETACH_ISLAND, L"Move Island to New Window");
            hasItemCommands = true;
        }
    }

    PopulateSavedGroupsMenu(menu, hasItemCommands);

    TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON, screenPt.x, screenPt.y, 0, m_hwnd, nullptr);
    DestroyMenu(menu);
}

void TabBandWindow::PopulateHiddenTabsMenu(HMENU menu, int groupIndex) {
    if (!menu) {
        return;
    }
    m_hiddenTabCommands.clear();

    if (!m_owner) {
        AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_HIDDEN_TAB_BASE, L"No hidden tabs");
        return;
    }

    const auto hiddenTabs = m_owner->GetHiddenTabs(groupIndex);
    if (hiddenTabs.empty()) {
        AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_HIDDEN_TAB_BASE, L"No hidden tabs");
        return;
    }

    UINT command = IDM_HIDDEN_TAB_BASE;
    for (const auto& entry : hiddenTabs) {
        AppendMenuW(menu, MF_STRING, command, entry.second.c_str());
        m_hiddenTabCommands.emplace_back(command, entry.first);
        ++command;
    }
}

void TabBandWindow::PopulateSavedGroupsMenu(HMENU parent, bool addSeparator) {
    if (!parent || !m_owner) {
        return;
    }

    HMENU groupsMenu = CreatePopupMenu();
    if (!groupsMenu) {
        return;
    }

    const auto names = m_owner->GetSavedGroupNames();
    if (names.empty()) {
        AppendMenuW(groupsMenu, MF_STRING | MF_GRAYED, 0, L"No Saved Groups");
    } else {
        UINT command = IDM_LOAD_SAVED_GROUP_BASE;
        for (const auto& name : names) {
            if (command > IDM_LOAD_SAVED_GROUP_LAST) {
                break;
            }
            AppendMenuW(groupsMenu, MF_STRING, command, name.c_str());
            m_savedGroupCommands.emplace_back(command, name);
            ++command;
        }
    }

    if (addSeparator) {
        AppendMenuW(parent, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(parent, MF_POPUP, reinterpret_cast<UINT_PTR>(groupsMenu), L"Groups");
    AppendMenuW(parent, MF_STRING, IDM_CREATE_SAVED_GROUP, L"Create Saved Group...");
}

int TabBandWindow::ResolveInsertGroupIndex() const {
    if (!m_owner) {
        return -1;
    }
    if (m_contextHit.hit && m_contextHit.location.groupIndex >= 0) {
        return m_contextHit.location.groupIndex;
    }
    return m_owner->GetGroupCount() - 1;
}

int TabBandWindow::GroupCount() const {
    int count = 0;
    int lastGroup = std::numeric_limits<int>::min();
    for (const auto& item : m_tabData) {
        if (item.location.groupIndex < 0) {
            continue;
        }
        if (item.location.groupIndex != lastGroup) {
            ++count;
            lastGroup = item.location.groupIndex;
        }
    }
    return count;
}

const TabBandWindow::VisualItem* TabBandWindow::FindLastGroupHeader() const {
    for (auto it = m_items.rbegin(); it != m_items.rend(); ++it) {
        if (it->data.type == TabViewItemType::kGroupHeader) {
            return &(*it);
        }
    }
    return nullptr;
}

const TabBandWindow::VisualItem* TabBandWindow::FindVisualForHit(const HitInfo& hit) const {
    if (!hit.hit) {
        return nullptr;
    }

    for (const auto& item : m_items) {
        if (item.data.type != hit.type) {
            continue;
        }
        if (hit.type == TabViewItemType::kGroupHeader) {
            if (item.data.location.groupIndex == hit.location.groupIndex) {
                return &item;
            }
        } else if (hit.location.IsValid() && item.data.location.IsValid()) {
            if (item.data.location.groupIndex == hit.location.groupIndex &&
                item.data.location.tabIndex == hit.location.tabIndex) {
                return &item;
            }
        }
    }
    return nullptr;
}

LRESULT CALLBACK TabBandWindow::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    TabBandWindow* self = reinterpret_cast<TabBandWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (message == WM_NCCREATE) {
        auto create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
        self = static_cast<TabBandWindow*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        if (self) {
            self->m_hwnd = hwnd;
        }
    }

    if (!self) {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    switch (message) {
        case WM_CREATE: {
            self->m_newTabButton = CreateWindowExW(0, L"BUTTON", L"+", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                                   0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDC_NEW_TAB),
                                                   GetModuleHandleInstance(), nullptr);
            DragAcceptFiles(hwnd, TRUE);
            break;
        }
        case WM_SIZE: {
            const int width = LOWORD(lParam);
            const int height = HIWORD(lParam);
            self->Layout(width, height);
            break;
        }
        case WM_INITMENUPOPUP:
        case WM_DRAWITEM:
        case WM_MEASUREITEM: {
            LRESULT handled = 0;
            if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                return handled;
            }
            break;
        }
        case WM_MENUCHAR: {
            LRESULT handled = 0;
            if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                return handled;
            }
            break;
        }
        case WM_COMMAND: {
            self->HandleCommand(wParam, lParam);
            break;
        }
        case WM_LBUTTONDOWN: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            self->HandleMouseDown(pt);
            break;
        }
        case WM_LBUTTONUP: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            self->HandleMouseUp(pt);
            break;
        }
        case WM_MOUSEMOVE: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            self->HandleMouseMove(pt);
            break;
        }
        case WM_RBUTTONUP: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            ClientToScreen(hwnd, &pt);
            self->ShowContextMenu(pt);
            return 0;
        }
        case WM_LBUTTONDBLCLK: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            self->HandleDoubleClick(pt);
            break;
        }
        case WM_DROPFILES: {
            self->HandleFileDrop(reinterpret_cast<HDROP>(wParam));
            return 0;
        }
        case WM_CONTEXTMENU: {
            POINT screenPt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            if (screenPt.x == -1 && screenPt.y == -1) {
                screenPt.x = 0;
                screenPt.y = 0;
                ClientToScreen(hwnd, &screenPt);
            }
            self->ShowContextMenu(screenPt);
            return 0;
        }
        case WM_SHELLTABS_DEFER_NAVIGATE: {
            if (self->m_owner) {
                self->m_owner->OnDeferredNavigate();
            }
            return 0;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC dc = BeginPaint(hwnd, &ps);
            self->Draw(dc);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_ERASEBKGND: {
            return 1;
        }
        case WM_CAPTURECHANGED: {
            self->CancelDrag();
            break;
        }
        case WM_DESTROY: {
            DragAcceptFiles(hwnd, FALSE);
            self->ClearExplorerContext();
            self->ClearVisualItems();
            self->m_hwnd = nullptr;
            self->m_newTabButton = nullptr;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            break;
        }
        default:
            break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

}  // namespace shelltabs
