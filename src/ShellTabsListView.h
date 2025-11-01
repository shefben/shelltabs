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

namespace Gdiplus {
class Bitmap;
}

#include <wrl/client.h>

#include "PaneHooks.h"
#include "Utilities.h"

namespace shelltabs {

class ShellTabsListView {
public:
    using HighlightResolver = std::function<bool(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight)>;
    struct BackgroundSource {
        std::wstring cacheKey;
        Gdiplus::Bitmap* bitmap = nullptr;
    };
    using BackgroundResolver = std::function<BackgroundSource()>;
    using AccentColorResolver = std::function<bool(COLORREF* accent, COLORREF* text)>;

    ShellTabsListView();
    ~ShellTabsListView();

    bool Initialize(HWND parent,
                    IFolderView2* folderView,
                    HighlightResolver resolver,
                    BackgroundResolver backgroundResolver = {},
                    AccentColorResolver accentResolver = {},
                    bool useAccentColors = true);

    HWND GetWindow() const noexcept { return m_window; }
    HWND GetListView() const noexcept { return m_listView; }

    void HandleInvalidationTargets(const PaneHighlightInvalidationTargets& targets) const;
    bool TryResolveHighlight(int index, PaneHighlight* highlight);

    void SetBackgroundResolver(BackgroundResolver resolver);
    void SetAccentColorResolver(AccentColorResolver resolver);
    void SetUseAccentColors(bool enabled);

private:
    struct CachedItem {
        UniquePidl pidl;
        std::wstring displayName;
        int imageIndex = -1;
    };

    struct BackgroundSurface {
        std::unique_ptr<Gdiplus::Bitmap> bitmap;
        SIZE size{0, 0};
        std::wstring cacheKey;
    };

    struct AccentResources {
        HBRUSH backgroundBrush = nullptr;
        HPEN focusPen = nullptr;
        COLORREF accentColor = 0;
        COLORREF textColor = 0;
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
    bool HandleCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result);
    bool PaintBackground(HDC dc);
    bool EnsureBackgroundSurface(const RECT& client, const BackgroundSource& source);
    void ResetBackgroundSurface();
    void ResetAccentResources();
    bool EnsureAccentResources(AccentResources** resources);
    bool ShouldUseAccentColors() const;
    void HandleListViewThemeChanged();

    HWND m_window = nullptr;
    HWND m_listView = nullptr;
    Microsoft::WRL::ComPtr<IFolderView2> m_folderView;
    HighlightResolver m_highlightResolver;
    BackgroundResolver m_backgroundResolver;
    AccentColorResolver m_accentResolver;
    bool m_useAccentColors = true;
    int m_itemCount = 0;
    std::unordered_map<int, CachedItem> m_cache;
    BackgroundSurface m_backgroundSurface;
    AccentResources m_accentResources;
};

}  // namespace shelltabs

