#include "PaneHooks.h"

#include "Utilities.h"

#include <mutex>
#include <unordered_map>
#include <algorithm>
#include <cwctype>

namespace shelltabs {

namespace {
std::mutex g_highlightMutex;
std::unordered_map<std::wstring, PaneHighlight> g_highlights;

std::wstring NormalizeHighlightKey(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    std::wstring normalized = NormalizeFileSystemPath(path);
    std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return normalized;
}

}  // namespace

PaneHookRouter::PaneHookRouter(PaneHighlightProvider* provider) : m_provider(provider) {}

void PaneHookRouter::SetHighlightProvider(PaneHighlightProvider* provider) {
    m_provider = provider;
}

void PaneHookRouter::SetListView(HWND listView) {
    m_listView = listView;
}

void PaneHookRouter::SetTreeView(HWND treeView) {
    m_treeView = treeView;
}

void PaneHookRouter::Reset() {
    m_listView = nullptr;
    m_treeView = nullptr;
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
    std::wstring normalized = NormalizeHighlightKey(path);
    if (normalized.empty()) {
        return;
    }

    std::scoped_lock lock(g_highlightMutex);
    g_highlights[normalized] = highlight;
}

void UnregisterPaneHighlight(const std::wstring& path) {
    std::wstring normalized = NormalizeHighlightKey(path);
    if (normalized.empty()) {
        return;
    }

    std::scoped_lock lock(g_highlightMutex);
    g_highlights.erase(normalized);
}

void ClearPaneHighlights() {
    std::scoped_lock lock(g_highlightMutex);
    g_highlights.clear();
}

bool TryGetPaneHighlight(const std::wstring& path, PaneHighlight* highlight) {
    std::wstring normalized = NormalizeHighlightKey(path);
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

}  // namespace shelltabs
