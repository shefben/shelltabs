#pragma once

#include "Utilities.h"
#include "FtpClient.h"

#include <string>
#include <vector>

#include <ocidl.h>
#include <shobjidl.h>
#include <shlobj.h>
#include <wrl/client.h>

namespace shelltabs {

enum class NamespaceMenuKind {
    FtpBackground,
    HttpBackground,
    FtpRootItem,
    HttpRootItem,
    FtpRemoteBackground,
    HttpRemoteBackground,
};

struct NamespaceMenuContext {
    NamespaceMenuKind kind = NamespaceMenuKind::FtpBackground;
    HWND ownerWindow = nullptr;
    PCIDLIST_ABSOLUTE folderPidl = nullptr;

    // Absolute PIDL of the item itself (for navigation on Open)
    PCIDLIST_ABSOLUTE itemAbsolutePidl = nullptr;

    // For root item menus: index of the entry in OptionsStore
    int entryIndex = -1;

    // For FTP remote operations
    FtpConnectionOptions ftpOptions{};
    FtpCredential ftpCredential{};
    std::vector<std::wstring> pathSegments;

    // For HTTP remote operations
    std::wstring httpUrl;

    // For item-specific operations
    std::wstring itemName;
    bool itemIsDirectory = false;
    bool itemEnabled = true;
};

class NamespaceContextMenu : public IContextMenu3, public IObjectWithSite {
public:
    explicit NamespaceContextMenu(const NamespaceMenuContext& context);
    ~NamespaceContextMenu();

    NamespaceContextMenu(const NamespaceContextMenu&) = delete;
    NamespaceContextMenu& operator=(const NamespaceContextMenu&) = delete;

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) override;
    IFACEMETHODIMP InvokeCommand(CMINVOKECOMMANDINFO* pici) override;
    IFACEMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, CHAR* pszName, UINT cchMax) override;

    // IContextMenu2
    IFACEMETHODIMP HandleMenuMsg(UINT uMsg, WPARAM wParam, LPARAM lParam) override;

    // IContextMenu3
    IFACEMETHODIMP HandleMenuMsg2(UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT* plResult) override;

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* pUnkSite) override;
    IFACEMETHODIMP GetSite(REFIID riid, void** ppvSite) override;

private:
    enum class MenuCommand : UINT {
        Open = 0,
        Add,
        Edit,
        Remove,
        EnableDisable,
        Refresh,
        NewFolder,
        UploadFiles,
        OpenInBrowser,
        Delete,
        Count,
    };

    void HandleAdd();
    void HandleEdit();
    void HandleRemove();
    void HandleEnableDisable();
    void HandleRefresh();
    void HandleNewFolder();
    void HandleUploadFiles();
    void HandleOpenInBrowser();
    void HandleDelete();

    std::atomic<ULONG> refCount_{1};
    NamespaceMenuContext context_;
    UINT idCmdFirst_ = 0;
    Microsoft::WRL::ComPtr<IUnknown> site_;
    UniquePidl ownedItemPidl_;  // Owns the item PIDL for navigation
};

}  // namespace shelltabs
