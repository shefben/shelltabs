#pragma once

#include <windows.h>

#include <CommCtrl.h>
#include <atomic>
#include <mutex>
#include <string>
#include <vector>

#include <exdisp.h>
#include <shobjidl.h>
#include <wrl/client.h>

namespace shelltabs {

struct PaneTheme {
    LOGFONTW font{};
    bool useCustomFont = false;

    COLORREF textColor = CLR_INVALID;
    COLORREF backgroundColor = CLR_INVALID;
    COLORREF selectedTextColor = CLR_INVALID;
    COLORREF selectedBackgroundColor = CLR_INVALID;
    COLORREF hotTextColor = CLR_INVALID;
    COLORREF hotBackgroundColor = CLR_INVALID;
};

class ExplorerWindowHook {
public:
    ExplorerWindowHook();
    ~ExplorerWindowHook();

    bool Initialize(IUnknown* site, IWebBrowser2* browser);
    void Shutdown();

    void Attach();

    void UpdateListTheme(const PaneTheme& theme);
    void UpdateTreeTheme(const PaneTheme& theme);

    static void AttachForExplorer(HWND explorer);

private:
    struct ScopedFont {
        HFONT handle = nullptr;
        void Reset();
        void Adopt(HFONT font);
        HFONT Get() const { return handle; }
    };

    enum class NotifyResult {
        kUnhandled,
        kModify,
        kHandled,
    };

    void RegisterThreadHook();
    void UnregisterThreadHook();

    static LRESULT CALLBACK CbtHookProc(int code, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK TreeParentSubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR id, DWORD_PTR refData);
    static LRESULT CALLBACK DefViewSubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                                UINT_PTR id, DWORD_PTR refData);

    void HandleCreateWindow(HWND created, const CBT_CREATEWND* data);
    void OnTreeWindowCreated(HWND tree);
    void OnListWindowCreated(HWND list);

    void AttachTreeParent(HWND parent);
    void DetachTreeParent();
    void AttachDefView(HWND defView);
    void DetachDefView();

    NotifyResult HandleTreeNotify(NMHDR* header, LRESULT* result);
    NotifyResult HandleTreeCustomDraw(NMTVCUSTOMDRAW* customDraw, LRESULT* result);
    NotifyResult HandleListNotify(NMHDR* header, LRESULT* result);
    NotifyResult HandleListCustomDraw(NMLVCUSTOMDRAW* customDraw, LRESULT* result);
    bool EnsureFolderView();
    void ResetFolderView();

    void ApplyTreeTheme();
    void ApplyListTheme();
    static HFONT CreateFontFromTheme(const PaneTheme& theme);
    void RegisterFrameHook();
    void UnregisterFrameHook();

    static constexpr UINT_PTR kTreeSubclassId = 0x54524B48;   // 'TRKH'
    static constexpr UINT_PTR kListSubclassId = 0x4C534B48;   // 'LSKH'

    Microsoft::WRL::ComPtr<IUnknown> site_;
    Microsoft::WRL::ComPtr<IWebBrowser2> browser_;
    Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider_;
    Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser_;
    Microsoft::WRL::ComPtr<IFolderView2> folderView_;

    HWND frame_ = nullptr;
    HWND tree_ = nullptr;
    HWND treeParent_ = nullptr;
    HWND defView_ = nullptr;
    HWND listView_ = nullptr;

    HHOOK cbtHook_ = nullptr;
    DWORD threadId_ = 0;

    bool treeSubclassed_ = false;
    bool defViewSubclassed_ = false;

    PaneTheme listTheme_{};
    PaneTheme treeTheme_{};

    ScopedFont listFont_{};
    ScopedFont treeFont_{};

};

}  // namespace shelltabs
