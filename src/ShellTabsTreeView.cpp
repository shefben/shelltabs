#include "ShellTabsTreeView.h"

#include "Logging.h"
#include "Utilities.h"
#include "ColorUtils.h"
#include "ExplorerThemeUtils.h"

#include <mutex>
#include <unordered_map>
#include <algorithm>

namespace shelltabs {
namespace {
std::mutex g_treeRegistryMutex;
std::unordered_map<HWND, ShellTabsTreeView*> g_treeRegistry;

// Mini hook: Override tree selection color to red
constexpr COLORREF kTreeSelectionAccent = RGB(255, 0, 0);
constexpr double kTreeSelectionBaseBlend = 0.38;
constexpr double kTreeSelectionFocusBoost = 0.14;
constexpr double kTreeSelectionBorderBlend = 0.55;
constexpr double kTreeSelectionBorderFocusBlend = 0.7;

double ClampRatio(double value) {
    return std::clamp(value, 0.0, 1.0);
}

BYTE BlendComponent(int base, int overlay, double ratio) {
    const double clamped = ClampRatio(ratio);
    const double blended = static_cast<double>(base) +
                           (static_cast<double>(overlay) - static_cast<double>(base)) * clamped;
    int result = static_cast<int>(blended + 0.5);
    result = std::clamp(result, 0, 255);
    return static_cast<BYTE>(result);
}

COLORREF BlendColor(COLORREF base, COLORREF overlay, double ratio) {
    return RGB(BlendComponent(GetRValue(base), GetRValue(overlay), ratio),
               BlendComponent(GetGValue(base), GetGValue(overlay), ratio),
               BlendComponent(GetBValue(base), GetBValue(overlay), ratio));
}

COLORREF ResolveTreeViewBackground(HWND treeView) {
    COLORREF background = TreeView_GetBkColor(treeView);
    if (background == CLR_NONE || background == CLR_DEFAULT) {
        background = GetSysColor(COLOR_WINDOW);
    }
    return background;
}

COLORREF ChooseSelectionTextColor(COLORREF background) {
    const double luminance = ComputeColorLuminance(background);
    return luminance > 0.6 ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

// Compute gradient text color for tree view items
// Gradient goes from blue (top) to purple (bottom)
COLORREF ComputeGradientTextColor(HWND treeView, HTREEITEM item, const RECT& itemRect) {
    if (!treeView || !item) {
        return RGB(0, 0, 0);  // Default black
    }

    RECT clientRect{};
    if (!GetClientRect(treeView, &clientRect) || clientRect.bottom <= clientRect.top) {
        return RGB(0, 0, 128);  // Default dark blue
    }

    // Define gradient colors: start (blue) and end (purple/magenta)
    constexpr COLORREF gradientStart = RGB(0, 64, 255);    // Bright blue
    constexpr COLORREF gradientEnd = RGB(192, 0, 192);     // Purple/Magenta

    // Calculate item's relative position in the visible area (0.0 to 1.0)
    const int itemCenter = (itemRect.top + itemRect.bottom) / 2;
    const int clientHeight = clientRect.bottom - clientRect.top;

    if (clientHeight <= 0) {
        return gradientStart;
    }

    double position = static_cast<double>(itemCenter - clientRect.top) / static_cast<double>(clientHeight);
    position = std::clamp(position, 0.0, 1.0);

    // Interpolate between gradient colors
    return BlendColor(gradientStart, gradientEnd, position);
}

bool PaintTreeSelection(NMTVCUSTOMDRAW* draw, HWND treeView) {
    if (!draw || !treeView || !IsWindow(treeView)) {
        return false;
    }

    if (IsSystemHighContrastActive()) {
        return false;
    }

    RECT fillRect = draw->nmcd.rc;
    if (fillRect.right <= fillRect.left || fillRect.bottom <= fillRect.top) {
        return false;
    }

    RECT clientRect{};
    if (GetClientRect(treeView, &clientRect)) {
        fillRect.left = clientRect.left;
        fillRect.right = clientRect.right;
    }

    const bool focused = (draw->nmcd.uItemState & CDIS_FOCUS) != 0;
    const COLORREF baseBackground = ResolveTreeViewBackground(treeView);
    const double selectionBlend = ClampRatio(kTreeSelectionBaseBlend + (focused ? kTreeSelectionFocusBoost : 0.0));
    const COLORREF selectionColor = BlendColor(baseBackground, kTreeSelectionAccent, selectionBlend);

    HBRUSH fillBrush = CreateSolidBrush(selectionColor);
    if (!fillBrush) {
        return false;
    }

    FillRect(draw->nmcd.hdc, &fillRect, fillBrush);
    DeleteObject(fillBrush);

    if (fillRect.bottom - fillRect.top >= 2) {
        const COLORREF topHighlight = BlendColor(selectionColor, RGB(255, 255, 255), 0.25);
        const COLORREF bottomShadow = BlendColor(selectionColor, RGB(0, 0, 0), 0.2);

        RECT topEdge = fillRect;
        topEdge.bottom = topEdge.top + 1;
        if (topEdge.bottom > topEdge.top) {
            if (HBRUSH topBrush = CreateSolidBrush(topHighlight)) {
                FillRect(draw->nmcd.hdc, &topEdge, topBrush);
                DeleteObject(topBrush);
            }
        }

        RECT bottomEdge = fillRect;
        bottomEdge.top = bottomEdge.bottom - 1;
        if (bottomEdge.bottom > bottomEdge.top) {
            if (HBRUSH bottomBrush = CreateSolidBrush(bottomShadow)) {
                FillRect(draw->nmcd.hdc, &bottomEdge, bottomBrush);
                DeleteObject(bottomBrush);
            }
        }
    }

    const double borderBlend = ClampRatio(focused ? kTreeSelectionBorderFocusBlend : kTreeSelectionBorderBlend);
    const COLORREF borderColor = BlendColor(selectionColor, kTreeSelectionAccent, borderBlend);

    RECT frameRect = fillRect;
    if (focused) {
        InflateRect(&frameRect, -1, -1);
    }
    if (frameRect.right > frameRect.left && frameRect.bottom > frameRect.top) {
        if (HBRUSH borderBrush = CreateSolidBrush(borderColor)) {
            FrameRect(draw->nmcd.hdc, &frameRect, borderBrush);
            DeleteObject(borderBrush);
        }
    }

    draw->clrTextBk = selectionColor;
    draw->clrText = ChooseSelectionTextColor(selectionColor);
    return true;
}

BOOL GetTreeViewItemWide(HWND treeView, TVITEMEXW* item) {
    if (!treeView || !item) {
        return FALSE;
    }
    return static_cast<BOOL>(
        SendMessageW(treeView, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(item)));
}

bool TryGetTreeItemRectInternal(HWND treeView, HTREEITEM item, RECT* rect) {
    if (!treeView || !item || !rect) {
        return false;
    }

    RECT itemRect{};
    if (!TreeView_GetItemRect(treeView, item, &itemRect, FALSE)) {
        return false;
    }

    RECT clientRect{};
    if (GetClientRect(treeView, &clientRect)) {
        itemRect.left = clientRect.left;
        itemRect.right = clientRect.right;
    }

    *rect = itemRect;
    return true;
}

bool IsPidlPointerValid(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return false;
    }

    __try {
        constexpr UINT kMaxTreeViewPidlSize = 64 * 1024;
        const UINT size = ILGetSize(pidl);
        if (size < sizeof(USHORT) || size > kMaxTreeViewPidlSize) {
            return false;
        }

        const auto* tail = reinterpret_cast<const USHORT*>(
            reinterpret_cast<const BYTE*>(pidl) + size - sizeof(USHORT));
        return *tail == 0;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

}  // namespace

std::unique_ptr<ShellTabsTreeView> ShellTabsTreeView::Create(
    HWND treeView,
    HighlightResolver resolver,
    const Microsoft::WRL::ComPtr<INameSpaceTreeControl>& namespaceControl) {
    if (!treeView) {
        return nullptr;
    }

    return std::unique_ptr<ShellTabsTreeView>(
        new (std::nothrow) ShellTabsTreeView(treeView, std::move(resolver), namespaceControl));
}

ShellTabsTreeView::ShellTabsTreeView(
    HWND treeView,
    HighlightResolver resolver,
    Microsoft::WRL::ComPtr<INameSpaceTreeControl> namespaceControl)
    : m_treeView(treeView), m_resolver(std::move(resolver)),
      m_namespaceTreeControl(std::move(namespaceControl)) {
    if (m_treeView) {
        SubscribeTreeViewForHighlights(m_treeView);
        std::scoped_lock lock(g_treeRegistryMutex);
        g_treeRegistry[m_treeView] = this;
    }
}

ShellTabsTreeView::~ShellTabsTreeView() {
    if (m_treeView) {
        {
            std::scoped_lock lock(g_treeRegistryMutex);
            g_treeRegistry.erase(m_treeView);
        }
        UnsubscribeTreeViewForHighlights(m_treeView);
    }
}

bool ShellTabsTreeView::HandleNotify(const NMHDR* header, LRESULT* result) {
    if (!header || !result || header->hwndFrom != m_treeView) {
        return false;
    }

    switch (header->code) {
        case NM_CUSTOMDRAW: {
            auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(const_cast<NMHDR*>(header));
            return HandleCustomDraw(draw, result);
        }
        case TVN_ITEMCHANGING: {
            const auto* change = reinterpret_cast<const NMTREEVIEWW*>(header);
            if (HandleSelectionChange(*change)) {
                return true;
            }
            break;
        }
        case TVN_ITEMEXPANDED: {
            const auto* expanded = reinterpret_cast<const NMTREEVIEWW*>(header);
            if (HandleItemExpanded(*expanded)) {
                return true;
            }
            break;
        }
        case TVN_SELCHANGEDW:
        case TVN_SELCHANGEDA: {
            if (m_treeView && IsWindow(m_treeView)) {
                InvalidateRect(m_treeView, nullptr, FALSE);
            }
            break;
        }
        default:
            break;
    }

    return false;
}

void ShellTabsTreeView::HandleInvalidationTargets(
    const PaneHighlightInvalidationTargets& targets) const {
    if (!m_treeView || !IsWindow(m_treeView)) {
        return;
    }

    if (targets.invalidateAll) {
        InvalidateRect(m_treeView, nullptr, FALSE);
        return;
    }

    if (targets.items.empty()) {
        return;
    }

    RECT invalidRect{};
    bool hasRect = false;

    for (const auto& target : targets.items) {
        HTREEITEM item = ResolveTargetItem(target);
        if (!item) {
            continue;
        }

        RECT itemRect{};
        if (target.includeTreeBranch) {
            if (TryGetTreeItemRect(item, &itemRect)) {
                RECT clientRect{};
                if (GetClientRect(m_treeView, &clientRect)) {
                    RECT branchRect{clientRect.left, itemRect.top, clientRect.right, clientRect.bottom};
                    if (hasRect) {
                        UnionRect(&invalidRect, &invalidRect, &branchRect);
                    } else {
                        invalidRect = branchRect;
                        hasRect = true;
                    }
                    continue;
                }
            }
            continue;
        }

        if (!TryGetTreeItemRect(item, &itemRect)) {
            continue;
        }

        if (hasRect) {
            UnionRect(&invalidRect, &invalidRect, &itemRect);
        } else {
            invalidRect = itemRect;
            hasRect = true;
        }
    }

    if (hasRect) {
        InvalidateRect(m_treeView, &invalidRect, FALSE);
    } else {
        InvalidateRect(m_treeView, nullptr, FALSE);
    }
}

bool ShellTabsTreeView::HandleCustomDraw(NMTVCUSTOMDRAW* draw, LRESULT* result) {
    if (!draw || !result) {
        return false;
    }

    switch (draw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT: {
            *result = CDRF_NOTIFYITEMDRAW;
            return true;
        }
        case CDDS_ITEMPREPAINT: {
            if (!m_treeView) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            auto* item = reinterpret_cast<HTREEITEM>(draw->nmcd.dwItemSpec);
            if (!item) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            if ((draw->nmcd.uItemState & CDIS_SELECTED) != 0) {
                if (PaintTreeSelection(draw, m_treeView)) {
                    *result = CDRF_NEWFONT;
                    return true;
                }
            }

            PaneHighlight highlight{};
            bool hasHighlight = ResolveHighlight(item, &highlight);

            bool applied = false;

            // Apply gradient text color to all items (unless they have a custom highlight color)
            if (!hasHighlight || !highlight.hasTextColor) {
                COLORREF gradientColor = ComputeGradientTextColor(m_treeView, item, draw->nmcd.rc);
                draw->clrText = gradientColor;
                applied = true;
            } else if (highlight.hasTextColor) {
                draw->clrText = highlight.textColor;
                applied = true;
            }

            if (hasHighlight && highlight.hasBackgroundColor) {
                draw->clrTextBk = highlight.backgroundColor;
                applied = true;
            }

            *result = applied ? CDRF_NEWFONT : CDRF_DODEFAULT;
            return true;
        }
        default:
            break;
    }

    return false;
}

bool ShellTabsTreeView::HandleSelectionChange(const NMTREEVIEWW& change) {
    if (!m_treeView) {
        return false;
    }

    InvalidateSelectionChange(change.itemOld.hItem, change.itemNew.hItem);
    return true;
}

bool ShellTabsTreeView::HandleItemExpanded(const NMTREEVIEWW& expanded) {
    if (!m_treeView) {
        return false;
    }

    InvalidateItemBranch(expanded.itemNew.hItem);
    return true;
}

bool ShellTabsTreeView::ResolveHighlight(HTREEITEM item, PaneHighlight* highlight) {
    if (!item || !m_treeView || !m_resolver) {
        return false;
    }

    TVITEMEXW treeItem{};
    treeItem.mask = TVIF_PARAM;
    treeItem.hItem = item;
    if (!GetTreeViewItemWide(m_treeView, &treeItem)) {
        return false;
    }

    PCIDLIST_ABSOLUTE pidl = reinterpret_cast<PCIDLIST_ABSOLUTE>(treeItem.lParam);
    if (IsPidlPointerValid(pidl)) {
        if (m_resolver(pidl, highlight)) {
            return true;
        }
    }

    if (pidl && !IsPidlPointerValid(pidl)) {
        LogMessage(LogLevel::Info, L"TreeView PIDL pointer invalid; attempting namespace fallback");
    }

    if (!m_namespaceTreeControl) {
        return false;
    }

    RECT itemBounds{};
    if (!TreeView_GetItemRect(m_treeView, item, &itemBounds, TRUE)) {
        return false;
    }

    const LONG centerX = itemBounds.left + (itemBounds.right - itemBounds.left) / 2;
    const LONG centerY = itemBounds.top + (itemBounds.bottom - itemBounds.top) / 2;
    POINT queryPoint{centerX, centerY};

    Microsoft::WRL::ComPtr<IShellItem> shellItem;
    HRESULT hr = m_namespaceTreeControl->HitTest(&queryPoint, &shellItem);
    if (FAILED(hr) || !shellItem) {
        return false;
    }

    PIDLIST_ABSOLUTE resolved = nullptr;
    hr = SHGetIDListFromObject(shellItem.Get(), &resolved);
    if (FAILED(hr) || !resolved) {
        return false;
    }

    UniquePidl owned(resolved);
    return m_resolver(owned.get(), highlight);
}

bool ShellTabsTreeView::TryGetTreeItemRect(HTREEITEM item, RECT* rect) const {
    return TryGetTreeItemRectInternal(m_treeView, item, rect);
}

void ShellTabsTreeView::InvalidateTreeRegion(const RECT* rect) const {
    if (!m_treeView || !IsWindow(m_treeView)) {
        return;
    }

    InvalidateRect(m_treeView, rect, FALSE);
}

void ShellTabsTreeView::InvalidateSelectionChange(HTREEITEM oldItem, HTREEITEM newItem) const {
    if (!m_treeView || !IsWindow(m_treeView)) {
        return;
    }

    RECT invalidRect{};
    bool hasRect = false;
    RECT itemRect{};

    if (TryGetTreeItemRect(oldItem, &itemRect)) {
        invalidRect = itemRect;
        hasRect = true;
    }

    if (TryGetTreeItemRect(newItem, &itemRect)) {
        if (hasRect) {
            UnionRect(&invalidRect, &invalidRect, &itemRect);
        } else {
            invalidRect = itemRect;
            hasRect = true;
        }
    }

    if (hasRect) {
        InvalidateTreeRegion(&invalidRect);
    } else {
        InvalidateTreeRegion(nullptr);
    }
}

void ShellTabsTreeView::InvalidateItemBranch(HTREEITEM item) const {
    if (!m_treeView || !IsWindow(m_treeView)) {
        return;
    }

    RECT itemRect{};
    if (TryGetTreeItemRect(item, &itemRect)) {
        RECT clientRect{};
        if (GetClientRect(m_treeView, &clientRect)) {
            RECT invalidRect{clientRect.left, itemRect.top, clientRect.right, clientRect.bottom};
            InvalidateTreeRegion(&invalidRect);
            return;
        }
    }

    InvalidateTreeRegion(nullptr);
}

HTREEITEM ShellTabsTreeView::ResolveTargetItem(const PaneHighlightInvalidationItem& target) const {
    if (target.treeItem) {
        return target.treeItem;
    }

    if (!m_treeView || !IsWindow(m_treeView) || !target.pidl) {
        return nullptr;
    }

    TVITEMEXW item{};
    item.mask = TVIF_PARAM;

    for (HTREEITEM current = TreeView_GetRoot(m_treeView); current;) {
        item.hItem = current;
        if (GetTreeViewItemWide(m_treeView, &item)) {
            if (reinterpret_cast<PCIDLIST_ABSOLUTE>(item.lParam) == target.pidl) {
                return current;
            }
        }

        HTREEITEM child = TreeView_GetChild(m_treeView, current);
        if (child) {
            current = child;
            continue;
        }

        while (current && !TreeView_GetNextSibling(m_treeView, current)) {
            current = TreeView_GetParent(m_treeView, current);
        }

        if (current) {
            current = TreeView_GetNextSibling(m_treeView, current);
        }
    }

    return nullptr;
}

ShellTabsTreeView* ShellTabsTreeView::FromHandle(HWND treeView) {
    if (!treeView) {
        return nullptr;
    }

    std::scoped_lock lock(g_treeRegistryMutex);
    auto it = g_treeRegistry.find(treeView);
    if (it == g_treeRegistry.end()) {
        return nullptr;
    }
    return it->second;
}

NamespaceTreeHost::NamespaceTreeHost(Microsoft::WRL::ComPtr<INameSpaceTreeControl> control,
                                     ShellTabsTreeView::HighlightResolver resolver)
    : m_control(std::move(control)), m_resolver(std::move(resolver)) {}

NamespaceTreeHost::~NamespaceTreeHost() = default;

bool NamespaceTreeHost::Initialize() {
    if (!m_control) {
        return false;
    }

    if (!EnsureTreeView()) {
        return false;
    }

    return m_treeView != nullptr;
}

bool NamespaceTreeHost::EnsureTreeView() {
    if (m_treeView) {
        return true;
    }

    if (!m_control) {
        return false;
    }

    Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
    if (FAILED(m_control.As(&oleWindow)) || !oleWindow) {
        return false;
    }

    HWND window = nullptr;
    if (FAILED(oleWindow->GetWindow(&window)) || !window) {
        return false;
    }

    m_window = window;
    m_treeView = ShellTabsTreeView::Create(window, m_resolver, m_control);
    return m_treeView != nullptr;
}

bool NamespaceTreeHost::HandleNotify(const NMHDR* header, LRESULT* result) {
    if (!EnsureTreeView() || !m_treeView) {
        return false;
    }

    return m_treeView->HandleNotify(header, result);
}

void NamespaceTreeHost::InvalidateAll() const {
    if (m_window && IsWindow(m_window)) {
        InvalidateRect(m_window, nullptr, FALSE);
    }
}

}  // namespace shelltabs

