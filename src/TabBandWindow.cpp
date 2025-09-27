#include "TabBandWindow.h"

#include <algorithm>
#include <cmath>
#include <memory>
#include <mutex>
#include <optional>
#include <limits>
#include <string>
#include <unordered_map>
#include <vector>

#include <windowsx.h>
#include <shellapi.h>
#include <ShlObj.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <winreg.h>
#include <dwmapi.h>

#include "Module.h"
#include "TabBand.h"
#include "Utilities.h"

namespace shelltabs {

namespace {
const wchar_t kWindowClassName[] = L"ShellTabsBandWindow";
constexpr int kButtonWidth = 26;
constexpr int kItemMinWidth = 60;
constexpr int kGroupMinWidth = 90;
constexpr int kGroupGap = 16;
constexpr int kTabGap = 6;
constexpr int kPaddingX = 12;
constexpr int kGroupPaddingX = 16;
constexpr int kToolbarGripWidth = 14;
constexpr int kToolbarGripDotSize = 2;
constexpr int kToolbarGripDotSpacing = 5;
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
constexpr int kIslandIndicatorWidth = 2;
constexpr int kCollapsedIndicatorPadding = 10;
constexpr int kIslandOutlineThickness = 1;
const wchar_t kThemePreferenceKey[] =
    L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
const wchar_t kThemePreferenceValue[] = L"AppsUseLightTheme";
constexpr UINT WM_SHELLTABS_EXTERNAL_DRAG = WM_APP + 60;
constexpr UINT WM_SHELLTABS_EXTERNAL_DRAG_LEAVE = WM_APP + 61;
constexpr UINT WM_SHELLTABS_EXTERNAL_DROP = WM_APP + 62;

struct WindowRegistry {
    std::mutex mutex;
    std::unordered_map<HWND, TabBandWindow*> windows;
};

WindowRegistry& GetWindowRegistry() {
    static WindowRegistry registry;
    return registry;
}

void RegisterWindow(HWND hwnd, TabBandWindow* window) {
    if (!hwnd || !window) {
        return;
    }
    auto& registry = GetWindowRegistry();
    std::scoped_lock lock(registry.mutex);
    registry.windows[hwnd] = window;
}

void UnregisterWindow(HWND hwnd, TabBandWindow* window) {
    if (!hwnd) {
        return;
    }
    auto& registry = GetWindowRegistry();
    std::scoped_lock lock(registry.mutex);
    auto it = registry.windows.find(hwnd);
    if (it != registry.windows.end() && (!window || it->second == window)) {
        registry.windows.erase(it);
    }
}

TabBandWindow* LookupWindow(HWND hwnd) {
    if (!hwnd) {
        return nullptr;
    }
    auto& registry = GetWindowRegistry();
    std::scoped_lock lock(registry.mutex);
    auto it = registry.windows.find(hwnd);
    if (it == registry.windows.end()) {
        return nullptr;
    }
    return it->second;
}

TabBandWindow* FindWindowFromPoint(const POINT& screenPt) {
    HWND target = WindowFromPoint(screenPt);
    while (target) {
        if (auto* window = LookupWindow(target)) {
            return window;
        }
        target = GetParent(target);
    }
    return nullptr;
}

void DispatchExternalMessage(HWND hwnd, UINT message) {
    if (!hwnd) {
        return;
    }
    SendMessageTimeoutW(hwnd, message, 0, 0, SMTO_ABORTIFHUNG | SMTO_BLOCK, 50, nullptr);
}

struct TransferPayload {
    enum class Type {
        None,
        Tab,
        Group,
    } type = Type::None;
    TabBand* source = nullptr;
    TabBand* target = nullptr;
    bool select = false;
    int targetGroupIndex = -1;
    int targetTabIndex = -1;
    bool createGroup = false;
    bool headerVisible = true;
    TabInfo tab;
    TabGroup group;
};

struct SharedDragState {
    std::mutex mutex;
    TabBandWindow* source = nullptr;
    TabBandWindow* hover = nullptr;
    POINT screen{};
    TabBandWindow::HitInfo origin;
    bool targetValid = false;
    TabBandWindow::DropTarget target;
    std::unique_ptr<TransferPayload> payload;
};

SharedDragState& GetSharedDragState() {
    static SharedDragState state;
    return state;
}

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

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

COLORREF BlendColors(COLORREF base, COLORREF accent, double ratio) {
    ratio = std::clamp(ratio, 0.0, 1.0);
    const double inverse = 1.0 - ratio;
    const int r = static_cast<int>(GetRValue(base) * inverse + GetRValue(accent) * ratio);
    const int g = static_cast<int>(GetGValue(base) * inverse + GetGValue(accent) * ratio);
    const int b = static_cast<int>(GetBValue(base) * inverse + GetBValue(accent) * ratio);
    return RGB(std::clamp(r, 0, 255), std::clamp(g, 0, 255), std::clamp(b, 0, 255));
}

double ComputeLuminance(COLORREF color) {
    const double r = GetRValue(color) / 255.0;
    const double g = GetGValue(color) / 255.0;
    const double b = GetBValue(color) / 255.0;
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

COLORREF AdjustForDarkTone(COLORREF color, double baseFactor, bool darkMode) {
    if (!darkMode) {
        return color;
    }
    double factor = std::clamp(baseFactor, 0.0, 1.0);
    const double luminance = ComputeLuminance(color);
    if (luminance > 0.3) {
        factor = std::clamp(factor + (luminance - 0.3) * 1.1, factor, 0.8);
    }
    return BlendColors(color, RGB(0, 0, 0), factor);
}

COLORREF ResolveIndicatorColor(const TabViewItem* header, const TabViewItem& tab) {
    if (header) {
        if (header->hasCustomOutline) {
            return header->outlineColor;
        }
        if (header->hasTagColor) {
            return header->tagColor;
        }
    }
    if (tab.hasCustomOutline) {
        return tab.outlineColor;
    }
    if (tab.hasTagColor) {
        return tab.tagColor;
    }
    return GetSysColor(COLOR_HOTLIGHT);
}

HFONT GetDefaultFont() {
    return static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
}

void ApplyImmersiveDarkMode(HWND hwnd, bool enabled) {
    if (!hwnd) {
        return;
    }

    using DwmSetWindowAttributeFn = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFn setWindowAttribute = []() -> DwmSetWindowAttributeFn {
        HMODULE module = GetModuleHandleW(L"dwmapi.dll");
        if (!module) {
            module = LoadLibraryW(L"dwmapi.dll");
        }
        if (!module) {
            return nullptr;
        }
        return reinterpret_cast<DwmSetWindowAttributeFn>(
            GetProcAddress(module, "DwmSetWindowAttribute"));
    }();

    if (!setWindowAttribute) {
        return;
    }

    const BOOL value = enabled ? TRUE : FALSE;
    setWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &value, sizeof(value));
}

}  // namespace

TabBandWindow::TabBandWindow(TabBand* owner) : m_owner(owner) { ResetThemePalette(); }

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

    if (m_hwnd) {
        RegisterWindow(m_hwnd, this);
    }

    return m_hwnd;
}

void TabBandWindow::Destroy() {
    CancelDrag();
    ClearExplorerContext();
    ClearVisualItems();
    CloseThemeHandles();
    m_darkMode = false;
    m_refreshingTheme = false;
    m_windowDarkModeInitialized = false;
    m_windowDarkModeValue = false;
    m_buttonDarkModeInitialized = false;
    m_buttonDarkModeValue = false;
    ResetThemePalette();

    if (m_newTabButton) {
        DestroyWindow(m_newTabButton);
        m_newTabButton = nullptr;
    }
    if (m_hwnd) {
        UnregisterWindow(m_hwnd, this);
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
    const int bandWidth = static_cast<int>(bounds.right - bounds.left);
    const int gripWidth = std::clamp(m_toolbarGripWidth, 0, std::max(0, bandWidth));
    int x = bounds.left + gripWidth + 4;
    int currentGroup = -1;
    TabViewItem currentHeader{};
    bool headerMetadata = false;
    bool expectFirstTab = false;
    bool pendingIndicator = false;
    TabViewItem indicatorHeader{};

    for (const auto& item : m_tabData) {
        if (item.type == TabViewItemType::kGroupHeader) {
            pendingIndicator = false;
            currentGroup = item.location.groupIndex;
            currentHeader = item;
            headerMetadata = true;
            expectFirstTab = true;

            const bool collapsed = item.collapsed;
            const bool hasVisibleTabs = item.visibleTabs > 0;
            if (!item.headerVisible && !collapsed && hasVisibleTabs) {
                indicatorHeader = item;
                pendingIndicator = true;
                continue;
            }

            VisualItem visual;
            visual.data = item;
            visual.firstInGroup = true;
            visual.collapsedPlaceholder = collapsed;

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

            x += width;
            continue;
        }

        VisualItem visual;
        visual.data = item;

        if (currentGroup != item.location.groupIndex) {
            currentGroup = item.location.groupIndex;
            headerMetadata = false;
            expectFirstTab = true;
            if (!m_items.empty()) {
                x += kGroupGap;
            }
            pendingIndicator = false;
        } else if (!expectFirstTab) {
            x += kTabGap;
        }

        if (expectFirstTab) {
            visual.firstInGroup = true;
            expectFirstTab = false;
        }
        visual.hasGroupHeader = headerMetadata;
        if (visual.hasGroupHeader) {
            visual.groupHeader = currentHeader;
        }
        if (pendingIndicator && visual.firstInGroup) {
            visual.hasGroupHeader = true;
            visual.groupHeader = indicatorHeader;
            visual.indicatorHandle = true;
            pendingIndicator = false;
            headerMetadata = true;
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

void TabBandWindow::DrawBackground(HDC dc, const RECT& bounds) const {
    if (!dc) {
        return;
    }

    if (bounds.right <= bounds.left || bounds.bottom <= bounds.top) {
        return;
    }

    const int bandWidth = static_cast<int>(bounds.right - bounds.left);
    const int gripWidth = std::clamp(m_toolbarGripWidth, 0, std::max(0, bandWidth));

    auto drawManualGrip = [&](const RECT& rect) {
        if (gripWidth <= 0) {
            return;
        }
        const int gripLeft = static_cast<int>(rect.left);
        const int gripRight = std::min(static_cast<int>(rect.right), gripLeft + gripWidth);
        if (gripRight <= gripLeft) {
            return;
        }

        const int centerX = (gripLeft + gripRight) / 2;
        const int centerY = (rect.top + rect.bottom) / 2;
        const COLORREF gripColor = BlendColors(m_themePalette.borderTop, m_themePalette.rebarBackground,
                                              m_darkMode ? 0.25 : 0.5);
        HBRUSH gripBrush = CreateSolidBrush(gripColor);
        if (gripBrush) {
            for (int i = -1; i <= 1; ++i) {
                const int offset = i * kToolbarGripDotSpacing;
                RECT dot{centerX - kToolbarGripDotSize, centerY + offset - kToolbarGripDotSize,
                         centerX + kToolbarGripDotSize + 1, centerY + offset + kToolbarGripDotSize + 1};
                FillRect(dc, &dot, gripBrush);
            }
            DeleteObject(gripBrush);
        }

        const COLORREF separatorColor = BlendColors(gripColor, m_themePalette.rebarBackground, 0.4);
        HPEN separatorPen = CreatePen(PS_SOLID, 1, separatorColor);
        if (separatorPen) {
            HPEN oldPen = static_cast<HPEN>(SelectObject(dc, separatorPen));
            MoveToEx(dc, gripRight, rect.top + 1, nullptr);
            LineTo(dc, gripRight, rect.bottom - 1);
            SelectObject(dc, oldPen);
            DeleteObject(separatorPen);
        }
    };

    auto drawManualBand = [&]() {
        RECT fillRect = bounds;
        HBRUSH background = CreateSolidBrush(m_themePalette.rebarBackground);
        if (background) {
            FillRect(dc, &fillRect, background);
            DeleteObject(background);
        }

        HPEN pen = CreatePen(PS_SOLID, 1, m_themePalette.borderTop);
        if (pen) {
            HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
            MoveToEx(dc, fillRect.left, fillRect.top, nullptr);
            LineTo(dc, fillRect.right, fillRect.top);
            SelectObject(dc, oldPen);
            DeleteObject(pen);
        }

        pen = CreatePen(PS_SOLID, 1, m_themePalette.borderBottom);
        if (pen) {
            HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
            const int bottom = fillRect.bottom - 1;
            MoveToEx(dc, fillRect.left, bottom, nullptr);
            LineTo(dc, fillRect.right, bottom);
            SelectObject(dc, oldPen);
            DeleteObject(pen);
        }

        drawManualGrip(fillRect);
    };

    if (m_rebarTheme) {
        RECT fillRect = bounds;
        if (SUCCEEDED(DrawThemeBackground(m_rebarTheme, dc, RP_BAND, 0, &fillRect, nullptr))) {
            if (gripWidth > 0) {
                RECT gripRect{fillRect.left, fillRect.top, fillRect.left + gripWidth, fillRect.bottom};
                if (gripRect.right > gripRect.left) {
                    HRESULT gripResult = DrawThemeBackground(m_rebarTheme, dc, RP_GRIPPER, 0, &gripRect, nullptr);
                    if (FAILED(gripResult)) {
                        gripResult = DrawThemeBackground(m_rebarTheme, dc, RP_GRIPPERVERT, 0, &gripRect, nullptr);
                    }
                    if (FAILED(gripResult)) {
                        drawManualGrip(fillRect);
                    } else {
                        RECT separatorRect{gripRect.right, gripRect.top + 2,
                                           std::min(gripRect.right + 1, fillRect.right), fillRect.bottom - 2};
                        if (separatorRect.bottom <= separatorRect.top) {
                            separatorRect.top = fillRect.top;
                            separatorRect.bottom = fillRect.bottom;
                        }
                        if (separatorRect.right > separatorRect.left) {
                            if (FAILED(DrawThemeEdge(m_rebarTheme, dc, RP_BAND, 0, &separatorRect, EDGE_ETCHED,
                                                     BF_LEFT, nullptr))) {
                                const COLORREF separatorColor =
                                    BlendColors(m_themePalette.borderTop, m_themePalette.rebarBackground, 0.4);
                                HPEN separatorPen = CreatePen(PS_SOLID, 1, separatorColor);
                                if (separatorPen) {
                                    HPEN oldPen = static_cast<HPEN>(SelectObject(dc, separatorPen));
                                    MoveToEx(dc, gripRect.right, fillRect.top + 1, nullptr);
                                    LineTo(dc, gripRect.right, fillRect.bottom - 1);
                                    SelectObject(dc, oldPen);
                                    DeleteObject(separatorPen);
                                }
                            }
                        }
                    }
                }
            }
            return;
        }
    }

    drawManualBand();
}

void TabBandWindow::Draw(HDC dc) const {
    if (!dc) {
        return;
    }

    RECT windowRect = m_clientRect;
    if (m_hwnd) {
        GetClientRect(m_hwnd, &windowRect);
    }

    const int width = windowRect.right - windowRect.left;
    const int height = windowRect.bottom - windowRect.top;
    if (width <= 0 || height <= 0) {
        return;
    }

    HDC memDC = CreateCompatibleDC(dc);
    if (!memDC) {
        PaintSurface(dc, windowRect);
        return;
    }

    HBITMAP buffer = CreateCompatibleBitmap(dc, width, height);
    if (!buffer) {
        DeleteDC(memDC);
        PaintSurface(dc, windowRect);
        return;
    }

    HGDIOBJ oldBitmap = SelectObject(memDC, buffer);
    RECT localRect{0, 0, width, height};
    PaintSurface(memDC, localRect);
    BitBlt(dc, windowRect.left, windowRect.top, width, height, memDC, 0, 0, SRCCOPY);
    SelectObject(memDC, oldBitmap);
    DeleteObject(buffer);
    DeleteDC(memDC);
}

void TabBandWindow::PaintSurface(HDC dc, const RECT& windowRect) const {
    if (!dc) {
        return;
    }

    DrawBackground(dc, windowRect);

    HFONT font = GetDefaultFont();
    HFONT oldFont = static_cast<HFONT>(SelectObject(dc, font));
    SetBkMode(dc, TRANSPARENT);

    const auto outlines = BuildGroupOutlines();

    for (const auto& item : m_items) {
        if (item.data.type == TabViewItemType::kGroupHeader) {
            DrawGroupHeader(dc, item);
        } else {
            DrawTab(dc, item);
        }
    }

    DrawGroupOutlines(dc, outlines);
    DrawDropIndicator(dc);
    DrawDragVisual(dc);

    if (oldFont) {
        SelectObject(dc, oldFont);
    }
}

COLORREF TabBandWindow::ResolveTabBackground(const TabViewItem& item) const {
    COLORREF base = item.selected ? m_themePalette.tabSelectedBase : m_themePalette.tabBase;
    if (item.selected) {
        base = BlendColors(base, m_accentColor, m_darkMode ? 0.45 : 0.35);
    }
    if (item.hasCustomOutline) {
        base = BlendColors(base, item.outlineColor, m_darkMode ? 0.35 : 0.25);
    } else if (item.hasTagColor) {
        base = BlendColors(base, item.tagColor, m_darkMode ? 0.3 : 0.2);
    }
    return base;
}

COLORREF TabBandWindow::ResolveGroupBackground(const TabViewItem& item) const {
    COLORREF base = m_themePalette.groupBase;
    if (item.selected) {
        base = BlendColors(base, m_accentColor, m_darkMode ? 0.4 : 0.25);
    }
    if (item.hasCustomOutline) {
        base = BlendColors(base, item.outlineColor, m_darkMode ? 0.35 : 0.25);
    } else if (item.hasTagColor) {
        base = BlendColors(base, item.tagColor, m_darkMode ? 0.3 : 0.2);
    }
    return base;
}

COLORREF TabBandWindow::ResolveTextColor(COLORREF background) const {
    return ComputeLuminance(background) > 0.55 ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

COLORREF TabBandWindow::ResolveTabTextColor(bool selected, COLORREF background) const {
    if (selected) {
        if (m_themePalette.tabSelectedTextValid) {
            return m_themePalette.tabSelectedText;
        }
        return GetSysColor(COLOR_HIGHLIGHTTEXT);
    }
    if (m_themePalette.tabTextValid) {
        return m_themePalette.tabText;
    }
    return ResolveTextColor(background);
}

COLORREF TabBandWindow::ResolveGroupTextColor(const TabViewItem& item, COLORREF background) const {
    if (item.selected && m_themePalette.tabSelectedTextValid) {
        return m_themePalette.tabSelectedText;
    }
    if (m_themePalette.groupTextValid) {
        return m_themePalette.groupText;
    }
    if (item.selected) {
        return GetSysColor(COLOR_HIGHLIGHTTEXT);
    }
    return ResolveTextColor(background);
}

std::vector<TabBandWindow::GroupOutline> TabBandWindow::BuildGroupOutlines() const {
    std::unordered_map<int, GroupOutline> outlines;
    for (const auto& item : m_items) {
        if (item.data.type != TabViewItemType::kTab) {
            continue;
        }
        if (item.data.location.groupIndex < 0) {
            continue;
        }

        RECT rect = item.bounds;
        if (item.indicatorHandle) {
            rect.left = std::max(m_clientRect.left, rect.left - kIslandIndicatorWidth);
        }

        auto& outline = outlines[item.data.location.groupIndex];
        if (!outline.initialized) {
            outline.groupIndex = item.data.location.groupIndex;
            outline.bounds = rect;
            outline.color = ResolveIndicatorColor(item.hasGroupHeader ? &item.groupHeader : nullptr, item.data);
            outline.initialized = true;
        } else {
            outline.bounds.left = std::min(outline.bounds.left, rect.left);
            outline.bounds.top = std::min(outline.bounds.top, rect.top);
            outline.bounds.right = std::max(outline.bounds.right, rect.right);
            outline.bounds.bottom = std::max(outline.bounds.bottom, rect.bottom);
            if (item.data.hasCustomOutline || item.data.hasTagColor) {
                outline.color = ResolveIndicatorColor(item.hasGroupHeader ? &item.groupHeader : nullptr, item.data);
            }
        }
    }

    std::vector<GroupOutline> result;
    result.reserve(outlines.size());
    for (auto& entry : outlines) {
        if (entry.second.initialized) {
            result.emplace_back(entry.second);
        }
    }
    std::sort(result.begin(), result.end(), [](const GroupOutline& a, const GroupOutline& b) {
        return a.bounds.left < b.bounds.left;
    });
    return result;
}

void TabBandWindow::DrawGroupOutlines(HDC dc, const std::vector<GroupOutline>& outlines) const {
    for (const auto& outline : outlines) {
        if (!outline.initialized) {
            continue;
        }
        RECT rect = outline.bounds;
        rect.left = std::max(rect.left, m_clientRect.left);
        rect.top = std::max(rect.top, m_clientRect.top);
        rect.right = std::min(rect.right, m_clientRect.right);
        rect.bottom = std::min(rect.bottom, m_clientRect.bottom);
        if (rect.right <= rect.left || rect.bottom <= rect.top) {
            continue;
        }

        COLORREF outlineColor = outline.color;
        if (m_darkMode) {
            outlineColor = LightenColor(outlineColor, 0.4);
        }
        HPEN pen = CreatePen(PS_SOLID, kIslandOutlineThickness, outlineColor);
        if (!pen) {
            continue;
        }
        HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));

        const int left = rect.left;
        const int right = rect.right - 1;
        const int top = rect.top;
        const int bottom = rect.bottom - 1;

        MoveToEx(dc, left, top, nullptr);
        LineTo(dc, right, top);
        MoveToEx(dc, left, top, nullptr);
        LineTo(dc, left, bottom);
        MoveToEx(dc, left, bottom, nullptr);
        LineTo(dc, right, bottom);
        MoveToEx(dc, right, top, nullptr);
        LineTo(dc, right, bottom);

        SelectObject(dc, oldPen);
        DeleteObject(pen);
    }
}

void TabBandWindow::RefreshTheme() {
    if (m_refreshingTheme) {
        return;
    }

    m_refreshingTheme = true;
    struct ThemeRefreshGuard {
        TabBandWindow* window;
        ~ThemeRefreshGuard() {
            if (window) {
                window->m_refreshingTheme = false;
            }
        }
    } guard{this};

    CloseThemeHandles();
    m_toolbarGripWidth = kToolbarGripWidth;
    if (!m_hwnd) {
        m_darkMode = false;
        m_windowDarkModeInitialized = false;
        m_windowDarkModeValue = false;
        m_buttonDarkModeInitialized = false;
        m_buttonDarkModeValue = false;
        ResetThemePalette();
        return;
    }

    SetWindowTheme(m_hwnd, L"Explorer", nullptr);
    const bool darkMode = IsSystemDarkMode();
    if (!m_windowDarkModeInitialized || darkMode != m_windowDarkModeValue) {
        ApplyImmersiveDarkMode(m_hwnd, darkMode);
        m_windowDarkModeInitialized = true;
        m_windowDarkModeValue = darkMode;
    }
    m_darkMode = darkMode;
    UpdateAccentColor();
    ResetThemePalette();
    m_tabTheme = OpenThemeData(m_hwnd, L"Tab");
    m_rebarTheme = OpenThemeData(m_hwnd, L"Rebar");
    UpdateThemePalette();
    UpdateToolbarMetrics();
    UpdateNewTabButtonTheme();
    RebuildLayout();
}

void TabBandWindow::UpdateAccentColor() {
    DWORD color = 0;
    BOOL opaque = FALSE;
    if (SUCCEEDED(DwmGetColorizationColor(&color, &opaque))) {
        m_accentColor = RGB((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
    } else {
        m_accentColor = GetSysColor(COLOR_HOTLIGHT);
    }
}

void TabBandWindow::ResetThemePalette() {
    m_themePalette.tabTextValid = false;
    m_themePalette.tabSelectedTextValid = false;
    m_themePalette.groupTextValid = false;

    const COLORREF windowColor = GetSysColor(COLOR_WINDOW);
    const COLORREF buttonColor = GetSysColor(COLOR_BTNFACE);
    const COLORREF windowBase = AdjustForDarkTone(windowColor, 0.55, m_darkMode);
    const COLORREF buttonBase = AdjustForDarkTone(buttonColor, 0.4, m_darkMode);

    m_themePalette.rebarBackground = m_darkMode ? BlendColors(buttonBase, windowBase, 0.55) : buttonBase;

    m_themePalette.borderTop = m_darkMode ? BlendColors(m_themePalette.rebarBackground, RGB(0, 0, 0), 0.6)
                                          : GetSysColor(COLOR_3DSHADOW);
    m_themePalette.borderBottom = m_darkMode ? BlendColors(m_themePalette.rebarBackground, RGB(255, 255, 255), 0.2)
                                             : GetSysColor(COLOR_3DLIGHT);

    m_themePalette.tabBase = windowBase;
    m_themePalette.tabSelectedBase = BlendColors(m_themePalette.tabBase, m_accentColor, m_darkMode ? 0.5 : 0.4);
    m_themePalette.tabText = GetSysColor(COLOR_WINDOWTEXT);
    m_themePalette.tabSelectedText = GetSysColor(COLOR_HIGHLIGHTTEXT);

    const double groupBlend = m_darkMode ? 0.6 : 0.25;
    m_themePalette.groupBase = BlendColors(buttonBase, windowBase, groupBlend);
    m_themePalette.groupText = GetSysColor(COLOR_WINDOWTEXT);
}

void TabBandWindow::UpdateThemePalette() {
    if (m_darkMode) {
        return;
    }

    if (m_rebarTheme) {
        COLORREF color = 0;
        if (SUCCEEDED(GetThemeColor(m_rebarTheme, RP_BAND, 0, TMT_FILLCOLORHINT, &color))) {
            m_themePalette.rebarBackground = color;
        }
        if (SUCCEEDED(GetThemeColor(m_rebarTheme, RP_BAND, 0, TMT_BORDERCOLORHINT, &color))) {
            m_themePalette.borderTop = color;
        }
        if (SUCCEEDED(GetThemeColor(m_rebarTheme, RP_BAND, 0, TMT_EDGEHIGHLIGHTCOLOR, &color))) {
            m_themePalette.borderBottom = color;
        }
    }

    if (m_tabTheme) {
        COLORREF color = 0;
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_BODY, 0, TMT_FILLCOLORHINT, &color))) {
            m_themePalette.tabBase = color;
            m_themePalette.groupBase = color;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, TIS_SELECTED, TMT_FILLCOLORHINT, &color))) {
            m_themePalette.tabSelectedBase = BlendColors(color, m_accentColor, 0.25);
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, TIS_SELECTED, TMT_TEXTCOLOR, &color))) {
            m_themePalette.tabSelectedText = color;
            m_themePalette.tabSelectedTextValid = true;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, TIS_NORMAL, TMT_TEXTCOLOR, &color))) {
            m_themePalette.tabText = color;
            m_themePalette.tabTextValid = true;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_BODY, 0, TMT_TEXTCOLOR, &color))) {
            m_themePalette.groupText = color;
            m_themePalette.groupTextValid = true;
        }
        if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_BODY, 0, TMT_BORDERCOLORHINT, &color))) {
            m_themePalette.borderBottom = color;
        }
    }

    if (m_darkMode) {
        m_themePalette.borderTop = BlendColors(m_themePalette.borderTop, RGB(0, 0, 0), 0.3);
        m_themePalette.borderBottom = BlendColors(m_themePalette.borderBottom, RGB(255, 255, 255), 0.15);
    }
}

void TabBandWindow::UpdateToolbarMetrics() {
    m_toolbarGripWidth = kToolbarGripWidth;
    if (!m_hwnd || !m_rebarTheme) {
        return;
    }

    HDC dc = GetDC(m_hwnd);
    if (!dc) {
        return;
    }

    int part = RP_GRIPPER;
    SIZE gripSize{0, 0};
    HRESULT hr = GetThemePartSize(m_rebarTheme, dc, part, 0, nullptr, TS_TRUE, &gripSize);
    if (FAILED(hr) || gripSize.cx <= 0) {
        part = RP_GRIPPERVERT;
        gripSize = {0, 0};
        hr = GetThemePartSize(m_rebarTheme, dc, part, 0, nullptr, TS_TRUE, &gripSize);
    }

    if (SUCCEEDED(hr) && gripSize.cx > 0) {
        int width = gripSize.cx;
        MARGINS margins{0, 0, 0, 0};
        if (SUCCEEDED(GetThemeMargins(m_rebarTheme, dc, part, 0, TMT_CONTENTMARGINS, nullptr, &margins))) {
            width += margins.cxLeftWidth + margins.cxRightWidth;
        }
        if (width > 0) {
            m_toolbarGripWidth = std::max(width, 8);
        }
    }

    ReleaseDC(m_hwnd, dc);
}

void TabBandWindow::CloseThemeHandles() {
    if (m_tabTheme) {
        CloseThemeData(m_tabTheme);
        m_tabTheme = nullptr;
    }
    if (m_rebarTheme) {
        CloseThemeData(m_rebarTheme);
        m_rebarTheme = nullptr;
    }
}

void TabBandWindow::UpdateNewTabButtonTheme() {
    if (!m_newTabButton) {
        m_buttonDarkModeInitialized = false;
        m_buttonDarkModeValue = false;
        return;
    }
    SetWindowTheme(m_newTabButton, L"Explorer", nullptr);
    if (!m_buttonDarkModeInitialized || m_buttonDarkModeValue != m_darkMode) {
        ApplyImmersiveDarkMode(m_newTabButton, m_darkMode);
        m_buttonDarkModeInitialized = true;
        m_buttonDarkModeValue = m_darkMode;
    }
    SendMessageW(m_newTabButton, WM_SETFONT, reinterpret_cast<WPARAM>(GetDefaultFont()), FALSE);
    InvalidateRect(m_newTabButton, nullptr, TRUE);
}

bool TabBandWindow::IsSystemDarkMode() const {
    DWORD value = 1;
    DWORD size = sizeof(value);
    const LSTATUS status = RegGetValueW(HKEY_CURRENT_USER, kThemePreferenceKey, kThemePreferenceValue, RRF_RT_DWORD,
                                        nullptr, &value, &size);
    if (status != ERROR_SUCCESS) {
        return false;
    }
    return value == 0;
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
                                      : (item.data.hasTagColor ? item.data.tagColor : m_accentColor);
        if (item.data.selected) {
            indicatorColor = BlendColors(indicatorColor, RGB(0, 0, 0), 0.2);
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
    COLORREF textColor = ResolveGroupTextColor(item.data, backgroundColor);
    COLORREF defaultOutline = m_accentColor;
    COLORREF outlineColor = item.data.hasCustomOutline
                                ? (selected ? DarkenColor(item.data.outlineColor, 0.25) : item.data.outlineColor)
                                : (item.data.hasTagColor ? DarkenColor(item.data.tagColor, 0.25)
                                                         : (selected ? BlendColors(defaultOutline, RGB(0, 0, 0), 0.2)
                                                                     : defaultOutline));

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
    const TabViewItem* indicatorSource = item.hasGroupHeader ? &item.groupHeader : nullptr;
    const bool hasAccent = item.data.hasCustomOutline || item.data.hasTagColor ||
                           (indicatorSource && (indicatorSource->hasCustomOutline || indicatorSource->hasTagColor));
    COLORREF accentColor = hasAccent ? ResolveIndicatorColor(indicatorSource, item.data) : m_accentColor;

    const int islandIndicator = item.indicatorHandle ? kIslandIndicatorWidth : 0;
    RECT tabRect = rect;
    tabRect.left += islandIndicator;

    int state = selected ? TIS_SELECTED : TIS_NORMAL;
    COLORREF computedBackground = ResolveTabBackground(item.data);
    COLORREF textColor = ResolveTabTextColor(selected, computedBackground);
    bool usedTheme = false;
    if (m_tabTheme && !m_darkMode) {
        if (SUCCEEDED(DrawThemeBackground(m_tabTheme, dc, TABP_TABITEM, state, &tabRect, nullptr))) {
            usedTheme = true;
            COLORREF themeText = 0;
            if (SUCCEEDED(GetThemeColor(m_tabTheme, TABP_TABITEM, state, TMT_TEXTCOLOR, &themeText))) {
                textColor = themeText;
            } else {
                textColor = ResolveTabTextColor(selected, computedBackground);
            }
        }
    }

    if (!usedTheme) {
        COLORREF backgroundColor = computedBackground;
        textColor = ResolveTabTextColor(selected, backgroundColor);
        COLORREF baseBorder = m_darkMode ? BlendColors(backgroundColor, RGB(255, 255, 255), selected ? 0.1 : 0.05)
                                         : BlendColors(backgroundColor, RGB(0, 0, 0), selected ? 0.15 : 0.1);
        COLORREF borderColor = hasAccent ? BlendColors(accentColor, RGB(0, 0, 0), selected ? 0.25 : 0.15)
                                         : baseBorder;

        RECT shapeRect = tabRect;
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

        COLORREF bottomLineColor = selected ? backgroundColor
                                            : (m_darkMode ? BlendColors(backgroundColor, RGB(0, 0, 0), 0.25)
                                                          : GetSysColor(COLOR_3DLIGHT));
        HPEN bottomPen = CreatePen(PS_SOLID, 1, bottomLineColor);
        if (bottomPen) {
            HPEN oldPen = static_cast<HPEN>(SelectObject(dc, bottomPen));
            MoveToEx(dc, tabRect.left + 1, rect.bottom - 1, nullptr);
            LineTo(dc, rect.right - 1, rect.bottom - 1);
            SelectObject(dc, oldPen);
            DeleteObject(bottomPen);
        }
        computedBackground = backgroundColor;
    }

    if (item.indicatorHandle) {
        RECT indicatorRect = rect;
        indicatorRect.right = indicatorRect.left + kIslandIndicatorWidth;
        COLORREF indicatorColor = hasAccent ? accentColor
                                            : (m_darkMode ? RGB(120, 120, 180) : GetSysColor(COLOR_HOTLIGHT));
        if (selected) {
            indicatorColor = DarkenColor(indicatorColor, 0.2);
        }
        HBRUSH indicatorBrush = CreateSolidBrush(indicatorColor);
        if (indicatorBrush) {
            FillRect(dc, &indicatorRect, indicatorBrush);
            DeleteObject(indicatorBrush);
        }
    }

    const int indicatorWidth = (item.data.splitSecondary || item.data.splitPrimary) ? kSplitIndicatorWidth : 0;
    const int badgeWidth = item.badgeWidth - indicatorWidth;

    int textLeft = rect.left + islandIndicator + kPaddingX;
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
        COLORREF badgeFillColor = m_darkMode ? BlendColors(computedBackground, badgeColor, 0.55)
                                             : LightenColor(badgeColor, 0.15);
        HBRUSH badgeBrush = CreateSolidBrush(badgeFillColor);
        if (badgeBrush) {
            RECT badgeFill = badgeRect;
            badgeFill.left -= kBadgePaddingX;
            badgeFill.right += kBadgePaddingX;
            FillRect(dc, &badgeFill, badgeBrush);
            DeleteObject(badgeBrush);
            COLORREF badgeOutline = m_darkMode ? RGB(70, 70, 74) : GetSysColor(COLOR_3DSHADOW);
            HPEN badgePen = CreatePen(PS_SOLID, 1, badgeOutline);
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
        SetTextColor(dc, ResolveTextColor(badgeFillColor));
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
    const DropTarget* indicator = nullptr;
    if (m_drag.dragging && m_drag.target.active && !m_drag.target.outside && m_drag.target.indicatorX >= 0) {
        indicator = &m_drag.target;
    } else if (m_externalDrop.active && m_externalDrop.target.active && !m_externalDrop.target.outside &&
               m_externalDrop.target.indicatorX >= 0) {
        indicator = &m_externalDrop.target;
    }

    if (!indicator) {
        return;
    }

    HPEN pen = CreatePen(PS_SOLID, 2, m_accentColor);
    HPEN oldPen = static_cast<HPEN>(SelectObject(dc, pen));
    const int x = indicator->indicatorX;
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
    if (hit.itemIndex < m_items.size()) {
        const auto& item = m_items[hit.itemIndex];
        m_drag.originSelected = item.data.selected;
    } else {
        m_drag.originSelected = false;
    }
    m_drag.start = pt;
    m_drag.current = pt;
    m_drag.hasCurrent = true;
}

void TabBandWindow::HandleMouseUp(const POINT& pt) {
    if (m_drag.dragging) {
        m_drag.current = pt;
        m_drag.hasCurrent = true;
        POINT screen = pt;
        ClientToScreen(m_hwnd, &screen);
        UpdateExternalDrag(screen);
        TabBandWindow* targetWindow = FindWindowFromPoint(screen);
        if (!targetWindow || targetWindow == this) {
            UpdateDropTarget(pt);
        } else {
            m_drag.target = {};
            m_drag.target.active = true;
            m_drag.target.outside = true;
        }
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
            auto& state = GetSharedDragState();
            std::scoped_lock lock(state.mutex);
            state.source = this;
            state.origin = m_drag.origin;
            state.screen = POINT{};
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
            state.payload.reset();
        }
    }

    if (m_drag.dragging) {
        POINT screen = pt;
        ClientToScreen(m_hwnd, &screen);
        UpdateExternalDrag(screen);
        TabBandWindow* targetWindow = FindWindowFromPoint(screen);
        if (!targetWindow || targetWindow == this) {
            UpdateDropTarget(pt);
        } else {
            m_drag.target = {};
            m_drag.target.active = true;
            m_drag.target.outside = true;
            RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
        }
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
    {
        auto& state = GetSharedDragState();
        TabBandWindow* hovered = nullptr;
        std::scoped_lock lock(state.mutex);
        if (state.source == this) {
            hovered = state.hover;
            state.source = nullptr;
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
            state.payload.reset();
        } else if (state.hover == this) {
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
        }
        if (hovered && hovered != this) {
            DispatchExternalMessage(hovered->GetHwnd(), WM_SHELLTABS_EXTERNAL_DRAG_LEAVE);
        }
    }
    m_externalDrop = {};
    m_drag = {};
    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
}

TabBandWindow::DropTarget TabBandWindow::ComputeDropTarget(const POINT& pt, const HitInfo& origin) const {
    DropTarget target{};
    target.active = true;

    if (pt.x < m_clientRect.left || pt.x > m_clientRect.right || pt.y < m_clientRect.top || pt.y > m_clientRect.bottom) {
        target.outside = true;
        return target;
    }

    HitInfo hit = HitTest(pt);
    if (!hit.hit) {
        if (origin.type == TabViewItemType::kTab && m_owner) {
            target.group = false;
            target.newGroup = true;
            target.floating = true;
            target.groupIndex = m_owner->GetGroupCount();
            target.tabIndex = 0;
            target.indicatorX = m_clientRect.right - 10;
        } else if (!m_items.empty()) {
            const VisualItem* lastHeader = FindLastGroupHeader();
            if (lastHeader) {
                if (origin.type == TabViewItemType::kGroupHeader) {
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
        return target;
    }

    const VisualItem& visual = m_items[hit.itemIndex];
    const int midX = (visual.bounds.left + visual.bounds.right) / 2;
    const bool leftSide = pt.x < midX;

    if (origin.type == TabViewItemType::kGroupHeader) {
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

    return target;
}

void TabBandWindow::UpdateDropTarget(const POINT& pt) {
    m_drag.target = ComputeDropTarget(pt, m_drag.origin);
    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::UpdateExternalDrag(const POINT& screenPt) {
    auto& state = GetSharedDragState();
    TabBandWindow* targetWindow = FindWindowFromPoint(screenPt);
    TabBandWindow* previousHover = nullptr;

    {
        std::scoped_lock lock(state.mutex);
        state.source = this;
        state.screen = screenPt;
        state.origin = m_drag.origin;
        previousHover = state.hover;
        state.targetValid = false;
        if (targetWindow == this) {
            state.hover = nullptr;
        }
    }

    if (previousHover && previousHover != targetWindow && previousHover != this) {
        DispatchExternalMessage(previousHover->GetHwnd(), WM_SHELLTABS_EXTERNAL_DRAG_LEAVE);
    }

    if (!targetWindow || targetWindow == this) {
        return;
    }

    {
        std::scoped_lock lock(state.mutex);
        if (state.source == this) {
            state.hover = targetWindow;
            state.targetValid = false;
        }
    }

    DispatchExternalMessage(targetWindow->GetHwnd(), WM_SHELLTABS_EXTERNAL_DRAG);
}

bool TabBandWindow::TryCompleteExternalDrop() {
    auto& state = GetSharedDragState();
    TabBandWindow* targetWindow = nullptr;
    DropTarget target{};

    {
        std::scoped_lock lock(state.mutex);
        if (state.source != this || !state.hover || state.hover == this || !state.targetValid) {
            return false;
        }
        targetWindow = state.hover;
        target = state.target;
    }

    if (!targetWindow || !targetWindow->m_owner || !m_owner || target.outside) {
        return false;
    }

    std::unique_ptr<TransferPayload> payload = std::make_unique<TransferPayload>();
    payload->target = targetWindow->m_owner;
    payload->targetGroupIndex = target.groupIndex;
    payload->targetTabIndex = target.tabIndex;
    payload->createGroup = target.newGroup;
    payload->headerVisible = !target.floating;
    payload->select = m_drag.originSelected;
    payload->source = m_owner;

    if (m_drag.origin.type == TabViewItemType::kGroupHeader) {
        auto detachedGroup = m_owner->DetachGroupForTransfer(m_drag.origin.location.groupIndex, nullptr);
        if (!detachedGroup) {
            return false;
        }
        payload->type = TransferPayload::Type::Group;
        payload->group = std::move(*detachedGroup);
    } else if (m_drag.origin.location.IsValid()) {
        auto detachedTab = m_owner->DetachTabForTransfer(m_drag.origin.location, nullptr);
        if (!detachedTab) {
            return false;
        }
        payload->type = TransferPayload::Type::Tab;
        payload->tab = std::move(*detachedTab);
    } else {
        return false;
    }

    {
        std::scoped_lock lock(state.mutex);
        state.payload = std::move(payload);
        state.source = nullptr;
        state.hover = nullptr;
        state.targetValid = false;
        state.target = {};
    }

    DispatchExternalMessage(targetWindow->GetHwnd(), WM_SHELLTABS_EXTERNAL_DROP);
    return true;
}

void TabBandWindow::HandleExternalDragUpdate() {
    auto& state = GetSharedDragState();
    POINT screen{};
    TabBandWindow* sourceWindow = nullptr;
    TabBandWindow::HitInfo origin;

    {
        std::scoped_lock lock(state.mutex);
        if (state.hover != this) {
            return;
        }
        screen = state.screen;
        sourceWindow = state.source;
        origin = state.origin;
    }

    if (!sourceWindow) {
        HandleExternalDragLeave();
        return;
    }

    POINT client = screen;
    ScreenToClient(m_hwnd, &client);
    DropTarget target = ComputeDropTarget(client, origin);

    {
        std::scoped_lock lock(state.mutex);
        if (state.hover == this) {
            state.target = target;
            state.targetValid = !target.outside;
        }
    }

    if (!target.outside) {
        m_externalDrop.active = true;
        m_externalDrop.target = target;
        m_externalDrop.source = sourceWindow;
    } else {
        m_externalDrop = {};
    }

    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::HandleExternalDragLeave() {
    {
        auto& state = GetSharedDragState();
        std::scoped_lock lock(state.mutex);
        if (state.hover == this) {
            state.hover = nullptr;
            state.targetValid = false;
            state.target = {};
        }
    }
    m_externalDrop = {};
    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void TabBandWindow::HandleExternalDropExecute() {
    std::unique_ptr<TransferPayload> payload;
    {
        auto& state = GetSharedDragState();
        std::scoped_lock lock(state.mutex);
        if (!state.payload || !m_owner || state.payload->target != m_owner) {
            return;
        }
        payload = std::move(state.payload);
    }

    if (!payload || !m_owner) {
        return;
    }

    if (payload->type == TransferPayload::Type::Tab) {
        m_owner->InsertTransferredTab(std::move(payload->tab), payload->targetGroupIndex, payload->targetTabIndex,
                                      payload->createGroup, payload->headerVisible, payload->select);
    } else if (payload->type == TransferPayload::Type::Group) {
        m_owner->InsertTransferredGroup(std::move(payload->group), payload->targetGroupIndex, payload->select);
    }

    m_externalDrop = {};
    RedrawWindow(m_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
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

    if (TryCompleteExternalDrop()) {
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
    const VisualItem* hitVisual = FindVisualForHit(hit);
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
        } else if (hit.type == TabViewItemType::kGroupHeader && hit.itemIndex >= 0) {
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
        } else if (hit.type == TabViewItemType::kGroupHeader && hit.location.groupIndex >= 0) {
            const bool headerVisible = m_owner->IsGroupHeaderVisible(hit.location.groupIndex);
            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND_HEADER,
                        headerVisible ? L"Hide Island Indicator" : L"Show Island Indicator");

            bool collapsed = false;
            size_t hiddenCount = 0;
            const TabViewItem* headerInfo = nullptr;
            if (hitVisual) {
                if (hitVisual->data.type == TabViewItemType::kGroupHeader) {
                    collapsed = hitVisual->data.collapsed;
                    hiddenCount = hitVisual->data.hiddenTabs;
                    headerInfo = &hitVisual->data;
                } else if (hitVisual->hasGroupHeader) {
                    collapsed = hitVisual->groupHeader.collapsed;
                    hiddenCount = hitVisual->groupHeader.hiddenTabs;
                    headerInfo = &hitVisual->groupHeader;
                }
            }

            AppendMenuW(menu, MF_STRING, IDM_TOGGLE_ISLAND,
                        collapsed ? L"Show Island" : L"Hide Island");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, IDM_NEW_ISLAND, L"New Island After");
            AppendMenuW(menu, MF_STRING, IDM_DETACH_ISLAND, L"Move Island to New Window");
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

            const bool splitAvailable = headerInfo && headerInfo->splitAvailable;
            const bool splitEnabled = headerInfo && headerInfo->splitEnabled;
            AppendMenuW(menu, MF_STRING | (splitAvailable ? 0 : MF_GRAYED), IDM_TOGGLE_SPLIT, L"Toggle Split View");
            AppendMenuW(menu, MF_STRING | (splitEnabled ? 0 : MF_GRAYED), IDM_CLEAR_SPLIT_SECONDARY,
                        L"Clear Split Companion");
            AppendMenuW(menu, MF_STRING | (splitEnabled ? 0 : MF_GRAYED), IDM_SWAP_SPLIT, L"Swap Split Panes");

            if (hiddenCount > 0) {
                AppendMenuW(menu, MF_STRING, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            } else {
                AppendMenuW(menu, MF_STRING | MF_GRAYED, IDM_UNHIDE_ALL, L"Unhide All Tabs");
            }
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

    if (hit.type == TabViewItemType::kGroupHeader) {
        for (const auto& item : m_items) {
            if (item.data.type == TabViewItemType::kTab && item.indicatorHandle &&
                item.data.location.groupIndex == hit.location.groupIndex) {
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

    auto fallback = [&]() -> LRESULT { return DefWindowProcW(hwnd, message, wParam, lParam); };

    if (!self) {
        return fallback();
    }

    auto dispatch = [&]() -> LRESULT {
        switch (message) {
            case WM_CREATE: {
                self->m_newTabButton = CreateWindowExW(0, L"BUTTON", L"+",
                                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                                       0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDC_NEW_TAB),
                                                       GetModuleHandleInstance(), nullptr);
                self->RefreshTheme();
                DragAcceptFiles(hwnd, TRUE);
                return 0;
            }
            case WM_SIZE: {
                const int width = LOWORD(lParam);
                const int height = HIWORD(lParam);
                self->Layout(width, height);
                return 0;
            }
            case WM_DRAWITEM: {
                LRESULT handled = 0;
                if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                    return handled;
                }
                return fallback();
            }
            case WM_INITMENUPOPUP:
            case WM_MEASUREITEM: {
                LRESULT handled = 0;
                if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                    return handled;
                }
                return fallback();
            }
            case WM_MENUCHAR: {
                LRESULT handled = 0;
                if (self->HandleExplorerMenuMessage(message, wParam, lParam, &handled)) {
                    return handled;
                }
                return fallback();
            }
            case WM_COMMAND: {
                self->HandleCommand(wParam, lParam);
                return 0;
            }
            case WM_LBUTTONDOWN: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                self->HandleMouseDown(pt);
                return 0;
            }
            case WM_LBUTTONUP: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                self->HandleMouseUp(pt);
                return 0;
            }
            case WM_MOUSEMOVE: {
                POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
                self->HandleMouseMove(pt);
                return 0;
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
                return 0;
            }
            case WM_DROPFILES: {
                self->HandleFileDrop(reinterpret_cast<HDROP>(wParam));
                return 0;
            }
            case WM_THEMECHANGED:
            case WM_SETTINGCHANGE:
            case WM_SYSCOLORCHANGE: {
                self->RefreshTheme();
                InvalidateRect(hwnd, nullptr, TRUE);
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
            case WM_SHELLTABS_REFRESH_COLORIZER: {
                if (self->m_owner) {
                    self->m_owner->OnColorizerRefresh();
                }
                return 0;
            }
            case WM_SHELLTABS_ENABLE_GIT_STATUS: {
                if (self->m_owner) {
                    self->m_owner->OnEnableGitStatus();
                }
                return 0;
            }
            case WM_SHELLTABS_REFRESH_GIT_STATUS: {
                if (self->m_owner) {
                    self->m_owner->OnGitStatusUpdated();
                }
                return 0;
            }
            case WM_SHELLTABS_EXTERNAL_DRAG: {
                self->HandleExternalDragUpdate();
                return 0;
            }
            case WM_SHELLTABS_EXTERNAL_DRAG_LEAVE: {
                self->HandleExternalDragLeave();
                return 0;
            }
            case WM_SHELLTABS_EXTERNAL_DROP: {
                self->HandleExternalDropExecute();
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
                HDC eraseDc = reinterpret_cast<HDC>(wParam);
                bool release = false;
                if (!eraseDc) {
                    eraseDc = GetDC(hwnd);
                    release = eraseDc != nullptr;
                }
                if (eraseDc) {
                    RECT client{};
                    GetClientRect(hwnd, &client);
                    self->DrawBackground(eraseDc, client);
                    if (release) {
                        ReleaseDC(hwnd, eraseDc);
                    }
                }
                return 1;
            }
            case WM_CAPTURECHANGED: {
                self->CancelDrag();
                return fallback();
            }
            case WM_DESTROY: {
                DragAcceptFiles(hwnd, FALSE);
                self->ClearExplorerContext();
                self->ClearVisualItems();
                self->CloseThemeHandles();
                UnregisterWindow(hwnd, self);
                self->m_hwnd = nullptr;
                self->m_newTabButton = nullptr;
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                return fallback();
            }
            default:
                return fallback();
        }
    };

    return GuardExplorerCall(L"TabBandWindow::WndProc", dispatch, fallback);
}

}  // namespace shelltabs
