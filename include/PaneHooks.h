#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <windows.h>
#include <CommCtrl.h>

#include <string>

namespace shelltabs {

struct PaneHighlight {
    bool hasTextColor = false;
    COLORREF textColor = 0;
    bool hasBackgroundColor = false;
    COLORREF backgroundColor = 0;
};

class PaneHighlightProvider {
public:
    virtual ~PaneHighlightProvider() = default;
    virtual bool TryGetListViewHighlight(HWND listView, int itemIndex, PaneHighlight* highlight) = 0;
    virtual bool TryGetTreeViewHighlight(HWND treeView, HTREEITEM item, PaneHighlight* highlight) = 0;
};

class PaneHookRouter {
public:
    explicit PaneHookRouter(PaneHighlightProvider* provider = nullptr);

    void SetHighlightProvider(PaneHighlightProvider* provider);
    void SetListView(HWND listView);
    void SetTreeView(HWND treeView);
    void Reset();

    bool HandleNotify(const NMHDR* header, LRESULT* result);

private:
    bool HandleListCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result);
    bool HandleTreeCustomDraw(NMTVCUSTOMDRAW* draw, LRESULT* result);

    PaneHighlightProvider* m_provider = nullptr;
    HWND m_listView = nullptr;
    HWND m_treeView = nullptr;
};

void RegisterPaneHighlight(const std::wstring& path, const PaneHighlight& highlight);
void UnregisterPaneHighlight(const std::wstring& path);
void ClearPaneHighlights();
bool TryGetPaneHighlight(const std::wstring& path, PaneHighlight* highlight);

}  // namespace shelltabs
