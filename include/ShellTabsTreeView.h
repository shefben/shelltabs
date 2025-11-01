#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>
#include <CommCtrl.h>
#include <ShlObj.h>

#include <functional>
#include <memory>

#include <wrl/client.h>

#include "PaneHooks.h"

namespace shelltabs {

class ShellTabsTreeView {
public:
    using HighlightResolver = std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)>;

    static std::unique_ptr<ShellTabsTreeView> Create(
        HWND treeView,
        HighlightResolver resolver,
        const Microsoft::WRL::ComPtr<INameSpaceTreeControl>& namespaceControl = nullptr);

    ~ShellTabsTreeView();

    bool HandleNotify(const NMHDR* header, LRESULT* result);
    void HandleInvalidationTargets(const PaneHighlightInvalidationTargets& targets) const;

    HWND GetHandle() const noexcept { return m_treeView; }
    static ShellTabsTreeView* FromHandle(HWND treeView);

private:
    ShellTabsTreeView(HWND treeView,
                      HighlightResolver resolver,
                      Microsoft::WRL::ComPtr<INameSpaceTreeControl> namespaceControl);

    bool HandleCustomDraw(NMTVCUSTOMDRAW* draw, LRESULT* result);
    bool HandleSelectionChange(const NMTREEVIEWW& change);
    bool HandleItemExpanded(const NMTREEVIEWW& expanded);

    bool ResolveHighlight(HTREEITEM item, PaneHighlight* highlight);
    bool TryGetTreeItemRect(HTREEITEM item, RECT* rect) const;
    void InvalidateTreeRegion(const RECT* rect) const;
    void InvalidateSelectionChange(HTREEITEM oldItem, HTREEITEM newItem) const;
    void InvalidateItemBranch(HTREEITEM item) const;
    HTREEITEM ResolveTargetItem(const PaneHighlightInvalidationItem& target) const;

    HWND m_treeView = nullptr;
    HighlightResolver m_resolver;
    Microsoft::WRL::ComPtr<INameSpaceTreeControl> m_namespaceTreeControl;
};

class NamespaceTreeHost {
public:
    NamespaceTreeHost(Microsoft::WRL::ComPtr<INameSpaceTreeControl> control,
                      ShellTabsTreeView::HighlightResolver resolver);
    ~NamespaceTreeHost();

    bool Initialize();
    bool HandleNotify(const NMHDR* header, LRESULT* result);
    void InvalidateAll() const;
    HWND GetWindow() const noexcept { return m_window; }

private:
    bool EnsureTreeView();

    Microsoft::WRL::ComPtr<INameSpaceTreeControl> m_control;
    ShellTabsTreeView::HighlightResolver m_resolver;
    HWND m_window = nullptr;
    std::unique_ptr<ShellTabsTreeView> m_treeView;
};

}  // namespace shelltabs

