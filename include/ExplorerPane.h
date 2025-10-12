#pragma once
#include <shobjidl.h>
#include <wrl/client.h>
#include <functional>
#include <string>

namespace shelltabs {

        class ExplorerPane {
        public:
                ExplorerPane();
                ~ExplorerPane();

                HRESULT Create(HWND parent, const RECT& rc);
                void Destroy();
                HWND GetHwnd() const;
                HWND GetListViewHwnd() const;
                void InvalidateView() const;

                HRESULT NavigateToPIDL(PCIDLIST_ABSOLUTE pidl);
                HRESULT NavigateToShellItem(IShellItem* item);
                HRESULT NavigateToPath(const std::wstring& path);
                HRESULT SetRect(const RECT& rc);

                using NavigationCallback = std::function<void(const std::wstring&)>;
                void SetNavigationCallback(NavigationCallback callback);
                std::wstring CurrentPath() const;

        private:
                class BrowserEvents;

                void HandleNavigationCompleted(PCIDLIST_ABSOLUTE pidl);
                void HandleViewCreated(IShellView* view);
                void RemoveSubclass();
                bool HandleNotify(NMHDR* header, LRESULT* result);
                bool HandleCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result);
                bool GetItemPath(int index, std::wstring* path) const;

                static LRESULT CALLBACK ViewSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                                         UINT_PTR id, DWORD_PTR refData);

                Microsoft::WRL::ComPtr<IExplorerBrowser> m_browser;
                Microsoft::WRL::ComPtr<IExplorerBrowserEvents> m_events;
                Microsoft::WRL::ComPtr<IFolderView2> m_folderView;
                HWND m_hwnd = nullptr;
                HWND m_defView = nullptr;
                HWND m_listView = nullptr;
                DWORD m_adviseCookie = 0;
                bool m_subclassed = false;

                NavigationCallback m_onNavigate;
                std::wstring m_currentPath;
        };

} // namespace shelltabs
