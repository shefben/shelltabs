#pragma once

#include <windows.h>
#include <CommCtrl.h>
#include <ShObjIdl_core.h>

#include <string>
#include <wrl/client.h>

namespace shelltabs {

class FolderViewColorizer {
public:
    FolderViewColorizer();
    ~FolderViewColorizer();

    void Attach(const Microsoft::WRL::ComPtr<IShellBrowser>& browser);
    void Detach();
    void Refresh();

private:
    Microsoft::WRL::ComPtr<IShellBrowser> m_shellBrowser;
    Microsoft::WRL::ComPtr<IFolderView2> m_folderView;
    HWND m_defView = nullptr;
    HWND m_listView = nullptr;

    bool ResolveView();
    void ResetView();
    bool EnsureSubclass();
    void RemoveSubclass();
    bool HandleNotify(NMHDR* header, LRESULT* result);
    bool HandleCustomDraw(NMLVCUSTOMDRAW* customDraw, LRESULT* result);
    bool GetItemPath(int index, std::wstring* path) const;

    static LRESULT CALLBACK SubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                         DWORD_PTR refData);
};

}  // namespace shelltabs

