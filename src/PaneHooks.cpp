#include "PaneHooks.h"

#include <atomic>
#include <mutex>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace shelltabs {

namespace {
std::mutex g_highlightMutex;
std::unordered_map<std::wstring, PaneHighlight> g_highlights;
std::unordered_set<HWND> g_listViewSubscribers;
std::unordered_set<HWND> g_treeViewSubscribers;
std::atomic<PaneHighlightInvalidationCallback> g_invalidationCallback = nullptr;

void CollectSubscribers(std::vector<HWND>& listViews, std::vector<HWND>& treeViews) {
    const PaneHighlightInvalidationCallback callback =
        g_invalidationCallback.load(std::memory_order_acquire);

    auto pruneAndCollect = [&](std::unordered_set<HWND>& subscribers, std::vector<HWND>& collected) {
        for (auto it = subscribers.begin(); it != subscribers.end();) {
            HWND hwnd = *it;
            if (!callback && !IsWindow(hwnd)) {
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

void DispatchInvalidations(const std::vector<HWND>& handles, HighlightPaneType paneType) {
    const PaneHighlightInvalidationCallback callback =
        g_invalidationCallback.load(std::memory_order_acquire);

    for (HWND hwnd : handles) {
        if (callback) {
            callback(hwnd, paneType);
            continue;
        }

        if (!IsWindow(hwnd)) {
            continue;
        }

        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void NotifyHighlightObservers(const std::vector<HWND>& listViews, const std::vector<HWND>& treeViews) {
    DispatchInvalidations(listViews, HighlightPaneType::ListView);
    DispatchInvalidations(treeViews, HighlightPaneType::TreeView);
}

}  // namespace

PaneHookRouter::PaneHookRouter(PaneHighlightProvider* provider) : m_provider(provider) {}

PaneHookRouter::~PaneHookRouter() {
    Reset();
}

void PaneHookRouter::SetHighlightProvider(PaneHighlightProvider* provider) {
    m_provider = provider;
}

void PaneHookRouter::SetListView(HWND listView) {
    if (m_listView == listView) {
        return;
    }

    if (m_listView) {
        UnsubscribeListViewForHighlights(m_listView);
    }

    m_listView = listView;

    if (m_listView) {
        SubscribeListViewForHighlights(m_listView);
    }
}

void PaneHookRouter::SetTreeView(HWND treeView) {
    if (m_treeView == treeView) {
        return;
    }

    if (m_treeView) {
        UnsubscribeTreeViewForHighlights(m_treeView);
    }

    m_treeView = treeView;

    if (m_treeView) {
        SubscribeTreeViewForHighlights(m_treeView);
    }
}

void PaneHookRouter::Reset() {
    SetListView(nullptr);
    SetTreeView(nullptr);
}

bool PaneHookRouter::HandleNotify(const NMHDR* header, LRESULT* result) {
    if (!header || !result) {
        return false;
    }

    if (header->hwndFrom == m_listView && header->code == NM_CUSTOMDRAW) {
        auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(const_cast<NMHDR*>(header));
        return HandleListCustomDraw(draw, result);
    }

    if (header->hwndFrom == m_treeView) {
        switch (header->code) {
            case NM_CUSTOMDRAW: {
                auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(const_cast<NMHDR*>(header));
                return HandleTreeCustomDraw(draw, result);
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
    }

    return false;
}

bool PaneHookRouter::HandleListCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result) {
    if (!draw || !result) {
        return false;
    }

    switch (draw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT: {
            *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW;
            return true;
        }
        case CDDS_ITEMPREPAINT:
        case CDDS_ITEMPREPAINT | CDDS_SUBITEM: {
            if ((draw->nmcd.dwDrawStage & CDDS_SUBITEM) != 0 && draw->iSubItem != 0) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            if (!m_provider || !m_listView) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            const int index = static_cast<int>(draw->nmcd.dwItemSpec);
            if (index < 0) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            PaneHighlight highlight{};
            if (!m_provider->TryGetListViewHighlight(m_listView, index, &highlight)) {
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

bool PaneHookRouter::HandleTreeCustomDraw(NMTVCUSTOMDRAW* draw, LRESULT* result) {
    if (!draw || !result) {
        return false;
    }

    switch (draw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT: {
            *result = CDRF_NOTIFYITEMDRAW;
            return true;
        }
        case CDDS_ITEMPREPAINT: {
            if (!m_provider || !m_treeView) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            auto* item = reinterpret_cast<HTREEITEM>(draw->nmcd.dwItemSpec);
            if (!item) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            PaneHighlight highlight{};
            if (!m_provider->TryGetTreeViewHighlight(m_treeView, item, &highlight)) {
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

void RegisterPaneHighlight(const std::wstring& path, const PaneHighlight& highlight) {
    if (path.empty()) {
        return;
    }

    std::vector<HWND> listViews;
    std::vector<HWND> treeViews;

    {
        std::scoped_lock lock(g_highlightMutex);
        g_highlights[path] = highlight;
        CollectSubscribers(listViews, treeViews);
    }

    NotifyHighlightObservers(listViews, treeViews);
}

void UnregisterPaneHighlight(const std::wstring& path) {
    if (path.empty()) {
        return;
    }

    std::vector<HWND> listViews;
    std::vector<HWND> treeViews;

    {
        std::scoped_lock lock(g_highlightMutex);
        if (g_highlights.erase(path) == 0) {
            return;
        }
        CollectSubscribers(listViews, treeViews);
    }

    NotifyHighlightObservers(listViews, treeViews);
}

void ClearPaneHighlights() {
    std::vector<HWND> listViews;
    std::vector<HWND> treeViews;

    {
        std::scoped_lock lock(g_highlightMutex);
        if (g_highlights.empty()) {
            return;
        }

        g_highlights.clear();
        CollectSubscribers(listViews, treeViews);
    }

    NotifyHighlightObservers(listViews, treeViews);
}

bool TryGetPaneHighlight(const std::wstring& path, PaneHighlight* highlight) {
    if (path.empty()) {
        return false;
    }

    std::scoped_lock lock(g_highlightMutex);
    auto it = g_highlights.find(path);
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
