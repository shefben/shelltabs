#pragma once

#include <CommCtrl.h>
#include <shobjidl.h>
#include <wrl/client.h>

#include <string>

namespace shelltabs {

// Subclasses the shell view hosted inside modern IFileDialog instances so the
// filename colour overrides show up in open/save dialogs as well.
class CommonDialogColorizer {
public:
    CommonDialogColorizer();
    ~CommonDialogColorizer();

    bool Attach(IFileDialog* dialog);
    void Detach();
    void Refresh();
    void InvalidateView() const;

    // Broadcasts a repaint notification to every active dialog colourizer.
    static void NotifyColorDataChanged();

private:
    class DialogEvents;

    void UpdateCurrentFolder();
    bool ResolveView();
    bool EnsureSubclass();
    void RemoveSubclass();
    bool HandleNotify(NMHDR* header, LRESULT* result);
    bool HandleCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result);
    bool GetItemPath(int index, std::wstring* path) const;

    static HWND FindDescendantByClass(HWND root, const wchar_t* className);
    static LRESULT CALLBACK SubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                         UINT_PTR id, DWORD_PTR refData);

    Microsoft::WRL::ComPtr<IFileDialog> m_dialog;
    Microsoft::WRL::ComPtr<IFileDialogEvents> m_events;
    Microsoft::WRL::ComPtr<IFolderView2> m_folderView;
    DWORD m_adviseCookie = 0;
    HWND m_dialogHwnd = nullptr;
    HWND m_defView = nullptr;
    HWND m_listView = nullptr;
    bool m_subclassed = false;
    std::wstring m_currentFolder;
};

}  // namespace shelltabs

