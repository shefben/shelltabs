#pragma once

#ifndef _WIN32_IE
#define _WIN32_IE 0x0700
#endif

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <CommCtrl.h>
#include <ShlObj.h>

#include <functional>
#include <memory>
#include <string>
#include <unordered_map>

#include <wrl/client.h>

#include "PaneHooks.h"
#include "Utilities.h"

namespace shelltabs {

class ShellTabsListView {
public:
    using HighlightResolver = std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)>;

    ShellTabsListView();
    ~ShellTabsListView();

    bool Initialize(HWND parent, IFolderView2* folderView, HighlightResolver resolver);

    HWND GetWindow() const noexcept { return m_window; }
    HWND GetListView() const noexcept { return m_listView; }

    void HandleInvalidationTargets(const PaneHighlightInvalidationTargets& targets) const;
    bool TryResolveHighlight(int index, PaneHighlight* highlight);

private:
    struct CachedItem {
        UniquePidl pidl;
        std::wstring displayName;
        int imageIndex = -1;
    };

    static ATOM EnsureWindowClass();
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK ListViewSubclassProc(HWND hwnd,
                                                 UINT message,
                                                 WPARAM wParam,
                                                 LPARAM lParam,
                                                 UINT_PTR subclassId,
                                                 DWORD_PTR referenceData);

    static ShellTabsListView* FromWindow(HWND hwnd);
    static ShellTabsListView* FromListView(HWND hwnd);

    bool CreateHostWindow(HWND parent);
    bool EnsureListView();
    void DestroyListView();
    void OnSize(int width, int height);
    LRESULT OnNotify(const NMHDR* header);
    bool HandleGetDispInfo(NMLVDISPINFOW* info);
    void HandleCacheHint(const NMLVCACHEHINT* hint);
    void HandleItemChanging(NMLVITEMCHANGE* change);
    CachedItem* EnsureCachedItem(int index);
    void PruneCache(int keepFrom, int keepTo);
    int ResolveIconIndex(PCIDLIST_ABSOLUTE pidl) const;

    HWND m_window = nullptr;
    HWND m_listView = nullptr;
    Microsoft::WRL::ComPtr<IFolderView2> m_folderView;
    HighlightResolver m_highlightResolver;
    int m_itemCount = 0;
    std::unordered_map<int, CachedItem> m_cache;
};

}  // namespace shelltabs

