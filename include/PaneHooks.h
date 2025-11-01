#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef _WIN32_IE
#define _WIN32_IE 0x0700
#endif

#include <windows.h>
#include <CommCtrl.h>
#include <ShlObj.h>

#include <functional>
#include <string>
#include <vector>
#include <memory>
#include <wrl/client.h>

namespace shelltabs {

struct PaneHighlight {
    bool hasTextColor = false;
    COLORREF textColor = 0;
    bool hasBackgroundColor = false;
    COLORREF backgroundColor = 0;
};

struct PaneHighlightInvalidationItem {
    PCIDLIST_ABSOLUTE pidl = nullptr;
    HTREEITEM treeItem = nullptr;
    bool includeTreeBranch = false;
};

struct PaneHighlightInvalidationTargets {
    std::vector<PaneHighlightInvalidationItem> items;
    bool invalidateAll = false;
};

enum class HighlightPaneType {
    ListView,
    TreeView,
};

using PaneHighlightInvalidationCallback = void (*)(HWND hwnd, HighlightPaneType paneType,
                                                   const PaneHighlightInvalidationTargets& targets);

class PaneHookRouter {
public:
    PaneHookRouter();
    ~PaneHookRouter();

    void SetTreeView(HWND treeView,
                     std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)> resolver = {},
                     INameSpaceTreeControl* namespaceTree = nullptr);
    void Reset();

    bool HandleNotify(const NMHDR* header, LRESULT* result);

private:
    HWND m_treeView = nullptr;
    std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)> m_resolver;
    Microsoft::WRL::ComPtr<INameSpaceTreeControl> m_namespaceTreeControl;
};

void RegisterPaneHighlight(const std::wstring& path, const PaneHighlight& highlight,
                           const PaneHighlightInvalidationTargets& listViewTargets = {},
                           const PaneHighlightInvalidationTargets& treeViewTargets = {});
void UnregisterPaneHighlight(const std::wstring& path,
                             const PaneHighlightInvalidationTargets& listViewTargets = {},
                             const PaneHighlightInvalidationTargets& treeViewTargets = {});
void ClearPaneHighlights();
bool TryGetPaneHighlight(const std::wstring& path, PaneHighlight* highlight);
std::wstring NormalizePaneHighlightKey(const std::wstring& path);
void SubscribeListViewForHighlights(HWND listView);
void SubscribeTreeViewForHighlights(HWND treeView);
void UnsubscribeListViewForHighlights(HWND listView);
void UnsubscribeTreeViewForHighlights(HWND treeView);
void SetPaneHighlightInvalidationCallback(PaneHighlightInvalidationCallback callback);

}  // namespace shelltabs
