#include "PaneHooks.h"

#include "Utilities.h"
#include "Logging.h"
#include <shobjidl_core.h>
#include <atomic>
#include <mutex>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <algorithm>
#include <cwctype>
#include <limits>

#ifndef ListView_FindItemW
static int ListView_FindItemW(HWND hwnd, int start, const LVFINDINFOW* findInfo) {
    return static_cast<int>(SendMessageW(
        hwnd, LVM_FINDITEMW, static_cast<WPARAM>(start),
        reinterpret_cast<LPARAM>(findInfo)));
}
#endif

#ifndef TreeView_GetItemW
static BOOL TreeView_GetItemW(HWND hwnd, TVITEMEXW* item) {
    return static_cast<BOOL>(
        SendMessageW(hwnd, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(item)));
}
#endif

namespace shelltabs {

std::wstring NormalizePaneHighlightKey(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    std::wstring normalized = NormalizeFileSystemPath(path);
    std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return normalized;
}

namespace {
std::mutex g_highlightMutex;
std::unordered_map<std::wstring, PaneHighlight> g_highlights;
std::unordered_set<HWND> g_listViewSubscribers;
std::unordered_set<HWND> g_treeViewSubscribers;
std::atomic<PaneHighlightInvalidationCallback> g_invalidationCallback = nullptr;

bool TryGetTreeItemRect(HWND treeView, HTREEITEM item, RECT* rect);

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

bool ResolveTreeItemHighlight(
    HWND treeView,
    HTREEITEM item,
    const std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)>& resolver,
    INameSpaceTreeControl* namespaceTree,
    PaneHighlight* highlight) {
    if (!treeView || !item || !resolver) {
        return false;
    }

    TVITEMEXW treeItem{};
    treeItem.mask = TVIF_PARAM;
    treeItem.hItem = item;
    if (TreeView_GetItemW(treeView, &treeItem)) {
        PCIDLIST_ABSOLUTE pidl = reinterpret_cast<PCIDLIST_ABSOLUTE>(treeItem.lParam);
        if (IsPidlPointerValid(pidl)) {
            if (resolver(pidl, highlight)) {
                return true;
            }
        } else if (pidl) {
            LogMessage(LogLevel::Info, L"TreeView PIDL pointer invalid; attempting namespace fallback");
        }
    }

    if (!namespaceTree) {
        return false;
    }

    RECT bounds{};
    if (!TreeView_GetItemRect(treeView, item, &bounds, TRUE)) {
        return false;
    }

    POINT queryPoint{};
    queryPoint.x = bounds.left + (bounds.right - bounds.left) / 2;
    queryPoint.y = bounds.top + (bounds.bottom - bounds.top) / 2;

    Microsoft::WRL::ComPtr<IShellItem> shellItem;
    HRESULT hr = namespaceTree->HitTest(&queryPoint, &shellItem);
    if (FAILED(hr) || !shellItem) {
        return false;
    }

    PIDLIST_ABSOLUTE resolved = nullptr;
    hr = SHGetIDListFromObject(shellItem.Get(), &resolved);
    if (FAILED(hr) || !resolved) {
        return false;
    }

    UniquePidl owned(resolved);
    return resolver(owned.get(), highlight);
}

void CollectSubscribers(std::vector<HWND>& listViews, std::vector<HWND>& treeViews) {
    auto pruneAndCollect = [&](std::unordered_set<HWND>& subscribers, std::vector<HWND>& collected) {
        for (auto it = subscribers.begin(); it != subscribers.end();) {
            HWND hwnd = *it;
            if (!IsWindow(hwnd)) {
                it = subscribers.erase(it);
                continue;
            }
            collected.push_back(hwnd);
            ++it;
        }
    };

    pruneAndCollect(g_listViewSubscribers, listViews);
    pruneAndCollect(g_treeViewSubscribers, treeViews);
}

bool InvalidateListViewTargets(HWND listView, const PaneHighlightInvalidationTargets& targets) {
    if (!listView || !IsWindow(listView)) {
        return false;
    }

    if (targets.invalidateAll) {
        InvalidateRect(listView, nullptr, FALSE);
        return true;
    }

    if (targets.items.empty()) {
        return false;
    }

    int minIndex = std::numeric_limits<int>::max();
    int maxIndex = std::numeric_limits<int>::min();
    RECT invalidRect{};
    bool hasInvalidRect = false;

    for (const auto& target : targets.items) {
        if (!target.pidl) {
            return false;
        }

        LVFINDINFOW find{};
        find.flags = LVFI_PARAM;
        find.lParam = reinterpret_cast<LPARAM>(target.pidl);

        const int index = ListView_FindItemW(listView, -1, &find);
        if (index < 0) {
            return false;
        }

        minIndex = std::min(minIndex, index);
        maxIndex = std::max(maxIndex, index);

        RECT itemRect{};
        if (ListView_GetItemRect(listView, index, &itemRect, LVIR_BOUNDS)) {
            if (hasInvalidRect) {
                UnionRect(&invalidRect, &invalidRect, &itemRect);
            } else {
                invalidRect = itemRect;
                hasInvalidRect = true;
            }
        }
    }

    if (minIndex == std::numeric_limits<int>::max() || maxIndex == std::numeric_limits<int>::min()) {
        return false;
    }

    ListView_RedrawItems(listView, minIndex, maxIndex);

    if (hasInvalidRect) {
        InvalidateRect(listView, &invalidRect, FALSE);
    } else {
        InvalidateRect(listView, nullptr, FALSE);
    }
    return true;
}

HTREEITEM ResolveTreeTarget(HWND treeView, const PaneHighlightInvalidationItem& target) {
    if (target.treeItem) {
        return target.treeItem;
    }

    if (!treeView || !IsWindow(treeView) || !target.pidl) {
        return nullptr;
    }

    TVITEMEXW item{};
    item.mask = TVIF_PARAM;

    for (HTREEITEM current = TreeView_GetRoot(treeView); current;) {
        item.hItem = current;
        if (TreeView_GetItemW(treeView, &item)) {
            if (reinterpret_cast<PCIDLIST_ABSOLUTE>(item.lParam) == target.pidl) {
                return current;
            }
        }

        HTREEITEM child = TreeView_GetChild(treeView, current);
        if (child) {
            current = child;
            continue;
        }

        while (current && !TreeView_GetNextSibling(treeView, current)) {
            current = TreeView_GetParent(treeView, current);
        }

        if (current) {
            current = TreeView_GetNextSibling(treeView, current);
        }
    }

    return nullptr;
}

bool InvalidateTreeViewTargets(HWND treeView, const PaneHighlightInvalidationTargets& targets) {
    if (!treeView || !IsWindow(treeView)) {
        return false;
    }

    if (targets.invalidateAll) {
        InvalidateRect(treeView, nullptr, FALSE);
        return true;
    }

    if (targets.items.empty()) {
        return false;
    }

    RECT invalidRect{};
    bool hasRect = false;

    for (const auto& target : targets.items) {
        HTREEITEM item = ResolveTreeTarget(treeView, target);
        if (!item) {
            return false;
        }

        RECT itemRect{};
        if (target.includeTreeBranch) {
            if (TryGetTreeItemRect(treeView, item, &itemRect)) {
                RECT clientRect{};
                if (GetClientRect(treeView, &clientRect)) {
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
            return false;
        }

        if (!TryGetTreeItemRect(treeView, item, &itemRect)) {
            return false;
        }

        if (hasRect) {
            UnionRect(&invalidRect, &invalidRect, &itemRect);
        } else {
            invalidRect = itemRect;
            hasRect = true;
        }
    }

    if (!hasRect) {
        return false;
    }

    InvalidateRect(treeView, &invalidRect, FALSE);
    return true;
}

void DispatchInvalidations(const std::vector<HWND>& handles, HighlightPaneType paneType,
                           const PaneHighlightInvalidationTargets& targets) {
    const PaneHighlightInvalidationCallback callback =
        g_invalidationCallback.load(std::memory_order_acquire);

    for (HWND hwnd : handles) {
        if (callback) {
            callback(hwnd, paneType, targets);
            continue;
        }

        if (!IsWindow(hwnd)) {
            continue;
        }

        bool handled = false;
        switch (paneType) {
            case HighlightPaneType::ListView:
                handled = InvalidateListViewTargets(hwnd, targets);
                break;
            case HighlightPaneType::TreeView:
                handled = InvalidateTreeViewTargets(hwnd, targets);
                break;
        }

        if (!handled) {
            InvalidateRect(hwnd, nullptr, FALSE);
        }
    }
}

void NotifyHighlightObservers(const std::vector<HWND>& listViews, const std::vector<HWND>& treeViews,
                              const PaneHighlightInvalidationTargets& listTargets,
                              const PaneHighlightInvalidationTargets& treeTargets) {
    DispatchInvalidations(listViews, HighlightPaneType::ListView, listTargets);
    DispatchInvalidations(treeViews, HighlightPaneType::TreeView, treeTargets);
}

bool TryGetTreeItemRect(HWND treeView, HTREEITEM item, RECT* rect) {
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

void InvalidateTreeRegion(HWND treeView, const RECT* rect) {
    if (!treeView || !IsWindow(treeView)) {
        return;
    }

    InvalidateRect(treeView, rect, FALSE);
}

void InvalidateTreeSelectionChange(HWND treeView, HTREEITEM oldItem, HTREEITEM newItem) {
    if (!treeView || !IsWindow(treeView)) {
        return;
    }

    RECT invalidRect{};
    bool hasRect = false;

    RECT itemRect{};
    if (TryGetTreeItemRect(treeView, oldItem, &itemRect)) {
        invalidRect = itemRect;
        hasRect = true;
    }

    if (TryGetTreeItemRect(treeView, newItem, &itemRect)) {
        if (hasRect) {
            UnionRect(&invalidRect, &invalidRect, &itemRect);
        } else {
            invalidRect = itemRect;
            hasRect = true;
        }
    }

    if (hasRect) {
        InvalidateTreeRegion(treeView, &invalidRect);
    } else {
        InvalidateTreeRegion(treeView, nullptr);
    }
}

void InvalidateTreeItemBranch(HWND treeView, HTREEITEM item) {
    if (!treeView || !IsWindow(treeView)) {
        return;
    }

    RECT itemRect{};
    if (TryGetTreeItemRect(treeView, item, &itemRect)) {
        RECT clientRect{};
        if (GetClientRect(treeView, &clientRect)) {
            RECT invalidRect{clientRect.left, itemRect.top, clientRect.right, clientRect.bottom};
            InvalidateTreeRegion(treeView, &invalidRect);
            return;
        }
    }

    InvalidateTreeRegion(treeView, nullptr);
}

}  // namespace

PaneHookRouter::PaneHookRouter() = default;

PaneHookRouter::~PaneHookRouter() {
    Reset();
}

void PaneHookRouter::SetTreeView(
    HWND treeView,
    std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)> resolver,
    INameSpaceTreeControl* namespaceTree) {
    if (m_treeView && m_treeView != treeView) {
        UnsubscribeTreeViewForHighlights(m_treeView);
    }

    m_treeView = treeView;
    m_resolver = std::move(resolver);
    m_namespaceTreeControl = namespaceTree;

    if (m_treeView) {
        SubscribeTreeViewForHighlights(m_treeView);
    } else {
        m_resolver = {};
        m_namespaceTreeControl.Reset();
    }
}

void PaneHookRouter::Reset() {
    SetTreeView(nullptr);
}

bool PaneHookRouter::HandleNotify(const NMHDR* header, LRESULT* result) {
    if (!header || !result) {
        return false;
    }

    if (!m_treeView || header->hwndFrom != m_treeView) {
        return false;
    }

    switch (header->code) {
        case NM_CUSTOMDRAW: {
            auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(const_cast<NMHDR*>(header));
            if (!draw) {
                return false;
            }

            switch (draw->nmcd.dwDrawStage) {
                case CDDS_PREPAINT:
                    *result = CDRF_NOTIFYITEMDRAW;
                    return true;
                case CDDS_ITEMPREPAINT: {
                    auto item = reinterpret_cast<HTREEITEM>(draw->nmcd.dwItemSpec);
                    PaneHighlight highlight{};
                    const bool resolved = ResolveTreeItemHighlight(
                        m_treeView, item, m_resolver, m_namespaceTreeControl.Get(), &highlight);
                    bool applied = false;
                    if (resolved) {
                        if (highlight.hasTextColor) {
                            draw->clrText = highlight.textColor;
                            applied = true;
                        }
                        if (highlight.hasBackgroundColor) {
                            draw->clrTextBk = highlight.backgroundColor;
                            applied = true;
                        }
                    }
                    *result = applied ? CDRF_NEWFONT : CDRF_DODEFAULT;
                    return true;
                }
                default:
                    break;
            }
            break;
        }
        case TVN_ITEMCHANGING: {
            const auto* change = reinterpret_cast<const NMTREEVIEWW*>(header);
            if (!change) {
                return false;
            }
            InvalidateTreeSelectionChange(m_treeView, change->itemOld.hItem, change->itemNew.hItem);
            *result = 0;
            return true;
        }
        case TVN_ITEMEXPANDED: {
            const auto* expanded = reinterpret_cast<const NMTREEVIEWW*>(header);
            if (!expanded) {
                return false;
            }
            InvalidateTreeItemBranch(m_treeView, expanded->itemNew.hItem);
            *result = 0;
            return true;
        }
        case TVN_SELCHANGEDW:
        case TVN_SELCHANGEDA:
            InvalidateRect(m_treeView, nullptr, FALSE);
            break;
        default:
            break;
    }

    return false;
}

void RegisterPaneHighlight(const std::wstring& path, const PaneHighlight& highlight,
                           const PaneHighlightInvalidationTargets& listViewTargets,
                           const PaneHighlightInvalidationTargets& treeViewTargets) {
    std::wstring normalized = NormalizePaneHighlightKey(path);
    if (normalized.empty()) {
        return;
    }

    std::vector<HWND> listViews;
    std::vector<HWND> treeViews;

    {
        std::scoped_lock lock(g_highlightMutex);
        g_highlights[normalized] = highlight;
        CollectSubscribers(listViews, treeViews);
    }

    NotifyHighlightObservers(listViews, treeViews, listViewTargets, treeViewTargets);
}

void UnregisterPaneHighlight(const std::wstring& path,
                             const PaneHighlightInvalidationTargets& listViewTargets,
                             const PaneHighlightInvalidationTargets& treeViewTargets) {
    std::wstring normalized = NormalizePaneHighlightKey(path);
    if (normalized.empty()) {
        return;
    }

    std::vector<HWND> listViews;
    std::vector<HWND> treeViews;

    {
        std::scoped_lock lock(g_highlightMutex);
        if (g_highlights.erase(normalized) == 0) {
            return;
        }
        CollectSubscribers(listViews, treeViews);
    }

    NotifyHighlightObservers(listViews, treeViews, listViewTargets, treeViewTargets);
}

void ClearPaneHighlights() {
    std::vector<HWND> listViews;
    std::vector<HWND> treeViews;

    PaneHighlightInvalidationTargets listTargets;
    listTargets.invalidateAll = true;
    PaneHighlightInvalidationTargets treeTargets;
    treeTargets.invalidateAll = true;

    {
        std::scoped_lock lock(g_highlightMutex);
        if (g_highlights.empty()) {
            return;
        }

        g_highlights.clear();
        CollectSubscribers(listViews, treeViews);
    }

    NotifyHighlightObservers(listViews, treeViews, listTargets, treeTargets);
}

bool TryGetPaneHighlight(const std::wstring& path, PaneHighlight* highlight) {
    std::wstring normalized = NormalizePaneHighlightKey(path);
    if (normalized.empty()) {
        return false;
    }

    std::scoped_lock lock(g_highlightMutex);
    auto it = g_highlights.find(normalized);
    if (it == g_highlights.end()) {
        return false;
    }

    if (highlight) {
        *highlight = it->second;
    }
    return true;
}

void SubscribeListViewForHighlights(HWND listView) {
    if (!listView) {
        return;
    }

    std::scoped_lock lock(g_highlightMutex);
    g_listViewSubscribers.insert(listView);
}

void SubscribeTreeViewForHighlights(HWND treeView) {
    if (!treeView) {
        return;
    }

    std::scoped_lock lock(g_highlightMutex);
    g_treeViewSubscribers.insert(treeView);
}

void UnsubscribeListViewForHighlights(HWND listView) {
    if (!listView) {
        return;
    }

    std::scoped_lock lock(g_highlightMutex);
    g_listViewSubscribers.erase(listView);
}

void UnsubscribeTreeViewForHighlights(HWND treeView) {
    if (!treeView) {
        return;
    }

    std::scoped_lock lock(g_highlightMutex);
    g_treeViewSubscribers.erase(treeView);
}

void SetPaneHighlightInvalidationCallback(PaneHighlightInvalidationCallback callback) {
    g_invalidationCallback.store(callback, std::memory_order_release);
}

}  // namespace shelltabs
