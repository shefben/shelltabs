#include "ShellTabsTreeView.h"

#include "Logging.h"
#include "Utilities.h"

#include <mutex>
#include <unordered_map>

namespace shelltabs {
namespace {
std::mutex g_treeRegistryMutex;
std::unordered_map<HWND, ShellTabsTreeView*> g_treeRegistry;

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

            PaneHighlight highlight{};
            if (!ResolveHighlight(item, &highlight)) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            bool applied = false;
            if (highlight.hasTextColor) {
                draw->clrText = highlight.textColor;
                applied = true;
            }
            if (highlight.hasBackgroundColor) {
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
    if (!TreeView_GetItemW(m_treeView, &treeItem)) {
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
        if (TreeView_GetItemW(m_treeView, &item)) {
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

