#pragma once

#include <windows.h>
#include <CommCtrl.h>
#include <shobjidl_core.h>
#include <wrl/client.h>

#include <optional>
#include <string>
#include <unordered_map>
#include <vector>

namespace shelltabs {

// Applies NameColorStore colors to one Explorer window's folder view and navigation tree.
//
// Follows the QTTabBar approach: the browser is asked for the classic SysListView32 folder view
// (FVO_VISTALAYOUT) because the default DirectUI items view renders text through DirectWrite and
// exposes no per-item draw hook. The list view's NM_CUSTOMDRAW (sent to SHELLDLL_DefView) and the
// tree's NM_CUSTOMDRAW (sent to NamespaceTreeControl) are then used to set each item's text color,
// which works identically in every view mode.
//
// Also injects a "Name Color" submenu into the folder view's item context menu.
class NameColorView {
public:
    NameColorView() = default;
    ~NameColorView();

    NameColorView(const NameColorView&) = delete;
    NameColorView& operator=(const NameColorView&) = delete;

    // Requests the classic list view for views created by this browser from now on.
    static void ApplyFolderViewMode(IShellBrowser* browser, bool classicListView);

    // (Re)binds to the browser's current view and tree. Safe to call on every navigation.
    void Attach(IShellBrowser* browser, HWND frame, bool enabled);
    void Detach();

    // Called when NameColorStore changed (possibly in another process).
    void OnColorsChanged();

private:
    static LRESULT CALLBACK DefViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                UINT_PTR id, DWORD_PTR refData);
    static LRESULT CALLBACK TreeParentSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR id, DWORD_PTR refData);
    static bool HandlePopupCommand(UINT command);

    void AttachDefView(HWND defView);
    void DetachDefView();
    void AttachTree(HWND frame);
    void DetachTree();
    void ApplyListViewTheme();

    LRESULT HandleListCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT defaultResult);
    LRESULT HandleTreeCustomDraw(NMTVCUSTOMDRAW* draw, LRESULT defaultResult);
    std::optional<COLORREF> ResolveListItemColor(int index);
    std::optional<COLORREF> ResolveTreeItemColor(HTREEITEM item);

    void InjectContextMenu(HMENU menu);
    void ExecuteCommand(UINT command);
    bool CollectSelection(std::vector<std::wstring>* paths, std::wstring* commonExtension, bool* anyColored);

    bool m_enabled = false;
    bool m_darkMode = false;
    Microsoft::WRL::ComPtr<IShellBrowser> m_browser;
    Microsoft::WRL::ComPtr<IFolderView> m_folderView;
    Microsoft::WRL::ComPtr<IShellFolder> m_folder;
    HWND m_frame = nullptr;
    HWND m_defView = nullptr;
    HWND m_listView = nullptr;
    HWND m_treeParent = nullptr;
    HWND m_tree = nullptr;
    bool m_contextMenuPending = false;

    // Pending context menu state (valid between injection and command execution).
    std::vector<std::wstring> m_menuPaths;
    std::wstring m_menuExtension;

    std::unordered_map<HTREEITEM, std::optional<COLORREF>> m_treeCache;
};

}  // namespace shelltabs
