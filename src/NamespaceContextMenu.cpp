#include "NamespaceContextMenu.h"
#include "OptionsStore.h"
#include "Logging.h"

#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <strsafe.h>

#include <atomic>
#include <string>

using Microsoft::WRL::ComPtr;

namespace shelltabs {

namespace {

// Simple input dialog using TaskDialog with an edit control fallback.
// Uses a basic MessageBox-based prompt: shows an input dialog via a tiny
// modal window with an edit control.
class InputDialog {
public:
    static bool Show(HWND parent, const wchar_t* title, const wchar_t* prompt,
                     const wchar_t* defaultValue, std::wstring* result) {
        if (!result) return false;
        result->clear();

        // Use a minimal dialog template built in memory
        struct DialogData {
            HWND parent;
            const wchar_t* title;
            const wchar_t* prompt;
            const wchar_t* defaultValue;
            std::wstring result;
            bool accepted;
        };

        DialogData data{parent, title, prompt, defaultValue, {}, false};

        // Build an in-memory dialog template
        // DLGTEMPLATE + items for static, edit, OK, Cancel
        alignas(4) BYTE buffer[2048]{};
        BYTE* ptr = buffer;

        auto WriteWord = [&](WORD value) {
            memcpy(ptr, &value, sizeof(WORD));
            ptr += sizeof(WORD);
        };
        auto WriteDword = [&](DWORD value) {
            memcpy(ptr, &value, sizeof(DWORD));
            ptr += sizeof(DWORD);
        };
        auto WriteWString = [&](const wchar_t* str) {
            size_t len = (wcslen(str) + 1) * sizeof(wchar_t);
            memcpy(ptr, str, len);
            ptr += len;
        };
        auto Align4 = [&]() {
            while ((reinterpret_cast<uintptr_t>(ptr) & 3) != 0) *ptr++ = 0;
        };

        // DLGTEMPLATE
        auto* dlg = reinterpret_cast<DLGTEMPLATE*>(ptr);
        dlg->style = DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE | DS_SETFONT;
        dlg->cdit = 4;  // static, edit, OK, Cancel
        dlg->x = 0; dlg->y = 0; dlg->cx = 220; dlg->cy = 75;
        ptr += sizeof(DLGTEMPLATE);
        WriteWord(0);  // menu
        WriteWord(0);  // class
        WriteWString(title);  // title
        WriteWord(8);  // font size
        WriteWString(L"MS Shell Dlg");  // font name
        Align4();

        // Static text (prompt label)
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
        item->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
        item->x = 7; item->y = 7; item->cx = 206; item->cy = 12;
        item->id = 100;
        ptr += sizeof(DLGITEMTEMPLATE);
        WriteWord(0xFFFF); WriteWord(0x0082);  // static class
        WriteWString(prompt);
        WriteWord(0);  // creation data
        Align4();

        // Edit control
        item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
        item->style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = 7; item->y = 22; item->cx = 206; item->cy = 14;
        item->id = 101;
        ptr += sizeof(DLGITEMTEMPLATE);
        WriteWord(0xFFFF); WriteWord(0x0081);  // edit class
        WriteWString(defaultValue ? defaultValue : L"");
        WriteWord(0);
        Align4();

        // OK button
        item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON;
        item->x = 108; item->y = 44; item->cx = 50; item->cy = 14;
        item->id = IDOK;
        ptr += sizeof(DLGITEMTEMPLATE);
        WriteWord(0xFFFF); WriteWord(0x0080);  // button class
        WriteWString(L"OK");
        WriteWord(0);
        Align4();

        // Cancel button
        item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
        item->x = 163; item->y = 44; item->cx = 50; item->cy = 14;
        item->id = IDCANCEL;
        ptr += sizeof(DLGITEMTEMPLATE);
        WriteWord(0xFFFF); WriteWord(0x0080);
        WriteWString(L"Cancel");
        WriteWord(0);

        INT_PTR ret = DialogBoxIndirectParamW(nullptr, reinterpret_cast<DLGTEMPLATE*>(buffer),
                                              parent, DialogProc, reinterpret_cast<LPARAM>(&data));
        if (ret == IDOK && data.accepted) {
            *result = std::move(data.result);
            return true;
        }
        return false;
    }

private:
    static INT_PTR CALLBACK DialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        struct DialogData {
            HWND parent;
            const wchar_t* title;
            const wchar_t* prompt;
            const wchar_t* defaultValue;
            std::wstring result;
            bool accepted;
        };

        switch (msg) {
            case WM_INITDIALOG: {
                SetWindowLongPtrW(hwnd, DWLP_USER, lParam);
                auto* data = reinterpret_cast<DialogData*>(lParam);
                if (data->defaultValue) {
                    SetDlgItemTextW(hwnd, 101, data->defaultValue);
                }
                return TRUE;
            }
            case WM_COMMAND:
                if (LOWORD(wParam) == IDOK) {
                    auto* data = reinterpret_cast<DialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                    if (data) {
                        wchar_t buf[1024]{};
                        GetDlgItemTextW(hwnd, 101, buf, ARRAYSIZE(buf));
                        data->result = buf;
                        data->accepted = true;
                    }
                    EndDialog(hwnd, IDOK);
                    return TRUE;
                }
                if (LOWORD(wParam) == IDCANCEL) {
                    EndDialog(hwnd, IDCANCEL);
                    return TRUE;
                }
                break;
        }
        return FALSE;
    }
};

std::wstring JoinFtpPath(const std::vector<std::wstring>& segments) {
    std::wstring path;
    for (const auto& segment : segments) {
        path += L"/";
        path += segment;
    }
    if (path.empty()) path = L"/";
    return path;
}

}  // namespace

NamespaceContextMenu::NamespaceContextMenu(const NamespaceMenuContext& context)
    : context_(context) {
    // Clone the item PIDL so it outlives the caller's stack
    if (context.itemAbsolutePidl) {
        ownedItemPidl_ = ClonePidl(context.itemAbsolutePidl);
        context_.itemAbsolutePidl = ownedItemPidl_.get();
    }
}

NamespaceContextMenu::~NamespaceContextMenu() = default;

IFACEMETHODIMP NamespaceContextMenu::QueryInterface(REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (riid == IID_IUnknown || riid == IID_IContextMenu) {
        *ppv = static_cast<IContextMenu*>(this);
    } else if (riid == IID_IContextMenu2) {
        *ppv = static_cast<IContextMenu2*>(this);
    } else if (riid == IID_IContextMenu3) {
        *ppv = static_cast<IContextMenu3*>(this);
    } else if (riid == IID_IObjectWithSite) {
        *ppv = static_cast<IObjectWithSite*>(this);
    } else {
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) NamespaceContextMenu::AddRef() {
    return ++refCount_;
}

IFACEMETHODIMP_(ULONG) NamespaceContextMenu::Release() {
    ULONG count = --refCount_;
    if (count == 0) delete this;
    return count;
}

IFACEMETHODIMP NamespaceContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst,
                                                       UINT idCmdLast, UINT uFlags) {
    UNREFERENCED_PARAMETER(idCmdLast);
    const bool defaultOnly = (uFlags & CMF_DEFAULTONLY) != 0;

    idCmdFirst_ = idCmdFirst;
    UINT insertIndex = indexMenu;
    UINT commandsUsed = 0;

    auto InsertItem = [&](MenuCommand cmd, const wchar_t* text, UINT extraFlags = 0) {
        MENUITEMINFOW mii{};
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_ID | MIIM_STRING | MIIM_FTYPE | MIIM_STATE;
        mii.fType = MFT_STRING;
        mii.fState = MFS_ENABLED | extraFlags;
        mii.wID = idCmdFirst + static_cast<UINT>(cmd);
        mii.dwTypeData = const_cast<wchar_t*>(text);
        InsertMenuItemW(hmenu, insertIndex++, TRUE, &mii);
        UINT cmdVal = static_cast<UINT>(cmd) + 1;
        if (cmdVal > commandsUsed) commandsUsed = cmdVal;
    };

    auto InsertSeparator = [&]() {
        MENUITEMINFOW mii{};
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_FTYPE;
        mii.fType = MFT_SEPARATOR;
        InsertMenuItemW(hmenu, insertIndex++, TRUE, &mii);
    };

    switch (context_.kind) {
        case NamespaceMenuKind::FtpBackground:
            if (defaultOnly) break;
            InsertItem(MenuCommand::Add, L"Add FTP Site...");
            InsertSeparator();
            InsertItem(MenuCommand::Refresh, L"Refresh");
            break;

        case NamespaceMenuKind::HttpBackground:
            if (defaultOnly) break;
            InsertItem(MenuCommand::Add, L"Add Web Folder...");
            InsertSeparator();
            InsertItem(MenuCommand::Refresh, L"Refresh");
            break;

        case NamespaceMenuKind::FtpRootItem:
            InsertItem(MenuCommand::Open, L"Open", MFS_DEFAULT);
            if (!defaultOnly) {
                InsertItem(MenuCommand::Edit, L"Edit...");
                InsertSeparator();
                InsertItem(MenuCommand::EnableDisable, context_.itemEnabled ? L"Disable" : L"Enable");
                InsertItem(MenuCommand::Remove, L"Remove");
            }
            break;

        case NamespaceMenuKind::HttpRootItem:
            InsertItem(MenuCommand::Open, L"Open", MFS_DEFAULT);
            if (!defaultOnly) {
                InsertItem(MenuCommand::Edit, L"Edit...");
                InsertSeparator();
                InsertItem(MenuCommand::EnableDisable, context_.itemEnabled ? L"Disable" : L"Enable");
                InsertItem(MenuCommand::Remove, L"Remove");
            }
            break;

        case NamespaceMenuKind::FtpRemoteBackground:
            if (defaultOnly) break;
            InsertItem(MenuCommand::NewFolder, L"New Folder...");
            InsertItem(MenuCommand::UploadFiles, L"Upload Files...");
            InsertSeparator();
            InsertItem(MenuCommand::Refresh, L"Refresh");
            break;

        case NamespaceMenuKind::HttpRemoteBackground:
            if (defaultOnly) break;
            InsertItem(MenuCommand::OpenInBrowser, L"Open in Browser");
            InsertSeparator();
            InsertItem(MenuCommand::Refresh, L"Refresh");
            break;
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, commandsUsed);
}

IFACEMETHODIMP NamespaceContextMenu::InvokeCommand(CMINVOKECOMMANDINFO* pici) {
    if (!pici) return E_POINTER;

    // We don't support verb strings
    if (IS_INTRESOURCE(pici->lpVerb)) {
        UINT offset = LOWORD(pici->lpVerb);
        auto cmd = static_cast<MenuCommand>(offset);
        switch (cmd) {
            case MenuCommand::Open:
                if (context_.itemAbsolutePidl) {
                    // Navigate in-place via IShellBrowser::BrowseObject
                    ComPtr<IShellBrowser> browser;
                    if (site_ && SUCCEEDED(IUnknown_QueryService(
                            site_.Get(), SID_STopLevelBrowser, IID_PPV_ARGS(&browser)))) {
                        browser->BrowseObject(context_.itemAbsolutePidl, SBSP_DEFBROWSER | SBSP_ABSOLUTE);
                    } else {
                        // Fallback: open in a new Explorer window
                        SHELLEXECUTEINFOW sei{};
                        sei.cbSize = sizeof(sei);
                        sei.fMask = SEE_MASK_IDLIST;
                        sei.hwnd = pici ? pici->hwnd : context_.ownerWindow;
                        sei.nShow = SW_SHOWNORMAL;
                        sei.lpIDList = const_cast<PIDLIST_ABSOLUTE>(context_.itemAbsolutePidl);
                        ShellExecuteExW(&sei);
                    }
                }
                return S_OK;
            case MenuCommand::Add:       HandleAdd(); return S_OK;
            case MenuCommand::Edit:      HandleEdit(); return S_OK;
            case MenuCommand::Remove:    HandleRemove(); return S_OK;
            case MenuCommand::EnableDisable: HandleEnableDisable(); return S_OK;
            case MenuCommand::Refresh:   HandleRefresh(); return S_OK;
            case MenuCommand::NewFolder: HandleNewFolder(); return S_OK;
            case MenuCommand::UploadFiles: HandleUploadFiles(); return S_OK;
            case MenuCommand::OpenInBrowser: HandleOpenInBrowser(); return S_OK;
            case MenuCommand::Delete:    HandleDelete(); return S_OK;
            default: break;
        }
    }
    return E_INVALIDARG;
}

IFACEMETHODIMP NamespaceContextMenu::GetCommandString(UINT_PTR, UINT, UINT*, CHAR*, UINT) {
    return E_NOTIMPL;
}

IFACEMETHODIMP NamespaceContextMenu::HandleMenuMsg(UINT, WPARAM, LPARAM) {
    return S_OK;
}

IFACEMETHODIMP NamespaceContextMenu::HandleMenuMsg2(UINT, WPARAM, LPARAM, LRESULT* plResult) {
    if (plResult) *plResult = 0;
    return S_OK;
}

IFACEMETHODIMP NamespaceContextMenu::SetSite(IUnknown* pUnkSite) {
    site_ = pUnkSite;
    return S_OK;
}

IFACEMETHODIMP NamespaceContextMenu::GetSite(REFIID riid, void** ppvSite) {
    if (!ppvSite) return E_POINTER;
    *ppvSite = nullptr;
    if (!site_) return E_FAIL;
    return site_->QueryInterface(riid, ppvSite);
}

void NamespaceContextMenu::HandleAdd() {
    bool isFtp = (context_.kind == NamespaceMenuKind::FtpBackground);

    if (isFtp) {
        std::wstring host;
        if (!InputDialog::Show(context_.ownerWindow, L"Add FTP Site", L"Host (e.g. ftp.example.com):", L"", &host))
            return;
        if (host.empty()) return;

        std::wstring portStr;
        InputDialog::Show(context_.ownerWindow, L"Add FTP Site", L"Port (default 21):", L"21", &portStr);
        uint16_t port = 21;
        if (!portStr.empty()) {
            int val = _wtoi(portStr.c_str());
            if (val > 0 && val <= 65535) port = static_cast<uint16_t>(val);
        }

        std::wstring displayName;
        InputDialog::Show(context_.ownerWindow, L"Add FTP Site", L"Display name:", host.c_str(), &displayName);
        if (displayName.empty()) displayName = host;

        auto options = OptionsStore::Instance().Get();
        FtpSiteEntry entry;
        entry.host = host;
        entry.port = port;
        entry.displayName = displayName;
        entry.enabled = true;
        options.ftpSiteEntries.push_back(std::move(entry));
        OptionsStore::Instance().Set(options);
        OptionsStore::Instance().Save();
    } else {
        std::wstring url;
        if (!InputDialog::Show(context_.ownerWindow, L"Add Web Folder", L"URL (e.g. https://example.com/files/):", L"", &url))
            return;
        if (url.empty()) return;

        std::wstring displayName;
        InputDialog::Show(context_.ownerWindow, L"Add Web Folder", L"Display name:", url.c_str(), &displayName);
        if (displayName.empty()) displayName = url;

        auto options = OptionsStore::Instance().Get();
        WebFolderEntry entry;
        entry.url = url;
        entry.displayName = displayName;
        entry.enabled = true;
        options.webFolderEntries.push_back(std::move(entry));
        OptionsStore::Instance().Set(options);
        OptionsStore::Instance().Save();
    }

    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleEdit() {
    if (context_.entryIndex < 0) return;

    bool isFtp = (context_.kind == NamespaceMenuKind::FtpRootItem);
    auto options = OptionsStore::Instance().Get();

    if (isFtp) {
        if (context_.entryIndex >= static_cast<int>(options.ftpSiteEntries.size())) return;
        auto& entry = options.ftpSiteEntries[context_.entryIndex];

        std::wstring host;
        if (!InputDialog::Show(context_.ownerWindow, L"Edit FTP Site", L"Host:", entry.host.c_str(), &host))
            return;
        if (!host.empty()) entry.host = host;

        std::wstring portStr;
        InputDialog::Show(context_.ownerWindow, L"Edit FTP Site", L"Port:", std::to_wstring(entry.port).c_str(), &portStr);
        if (!portStr.empty()) {
            int val = _wtoi(portStr.c_str());
            if (val > 0 && val <= 65535) entry.port = static_cast<uint16_t>(val);
        }

        std::wstring displayName;
        InputDialog::Show(context_.ownerWindow, L"Edit FTP Site", L"Display name:", entry.displayName.c_str(), &displayName);
        if (!displayName.empty()) entry.displayName = displayName;
    } else {
        if (context_.entryIndex >= static_cast<int>(options.webFolderEntries.size())) return;
        auto& entry = options.webFolderEntries[context_.entryIndex];

        std::wstring url;
        if (!InputDialog::Show(context_.ownerWindow, L"Edit Web Folder", L"URL:", entry.url.c_str(), &url))
            return;
        if (!url.empty()) entry.url = url;

        std::wstring displayName;
        InputDialog::Show(context_.ownerWindow, L"Edit Web Folder", L"Display name:", entry.displayName.c_str(), &displayName);
        if (!displayName.empty()) entry.displayName = displayName;
    }

    OptionsStore::Instance().Set(options);
    OptionsStore::Instance().Save();

    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleRemove() {
    if (context_.entryIndex < 0) return;

    int answer = MessageBoxW(context_.ownerWindow, L"Are you sure you want to remove this entry?",
                             L"ShellTabs", MB_YESNO | MB_ICONQUESTION);
    if (answer != IDYES) return;

    auto options = OptionsStore::Instance().Get();

    if (context_.kind == NamespaceMenuKind::FtpRootItem) {
        if (context_.entryIndex < static_cast<int>(options.ftpSiteEntries.size())) {
            options.ftpSiteEntries.erase(options.ftpSiteEntries.begin() + context_.entryIndex);
        }
    } else {
        if (context_.entryIndex < static_cast<int>(options.webFolderEntries.size())) {
            options.webFolderEntries.erase(options.webFolderEntries.begin() + context_.entryIndex);
        }
    }

    OptionsStore::Instance().Set(options);
    OptionsStore::Instance().Save();

    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleEnableDisable() {
    if (context_.entryIndex < 0) return;

    auto options = OptionsStore::Instance().Get();

    if (context_.kind == NamespaceMenuKind::FtpRootItem) {
        if (context_.entryIndex < static_cast<int>(options.ftpSiteEntries.size())) {
            options.ftpSiteEntries[context_.entryIndex].enabled = !options.ftpSiteEntries[context_.entryIndex].enabled;
        }
    } else {
        if (context_.entryIndex < static_cast<int>(options.webFolderEntries.size())) {
            options.webFolderEntries[context_.entryIndex].enabled = !options.webFolderEntries[context_.entryIndex].enabled;
        }
    }

    OptionsStore::Instance().Set(options);
    OptionsStore::Instance().Save();

    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleRefresh() {
    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleNewFolder() {
    std::wstring folderName;
    if (!InputDialog::Show(context_.ownerWindow, L"New Folder", L"Folder name:", L"New Folder", &folderName))
        return;
    if (folderName.empty()) return;

    std::wstring remotePath = JoinFtpPath(context_.pathSegments) + L"/" + folderName;

    FtpClient client;
    HRESULT hr = client.CreateDirectory(context_.ftpOptions, &context_.ftpCredential,
                                         remotePath, context_.ownerWindow);
    if (FAILED(hr)) {
        MessageBoxW(context_.ownerWindow, L"Failed to create directory. The server may not allow this operation.",
                    L"ShellTabs", MB_OK | MB_ICONERROR);
        return;
    }

    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_MKDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleUploadFiles() {
    ComPtr<IFileOpenDialog> dialog;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
    if (FAILED(hr)) return;

    FILEOPENDIALOGOPTIONS opts = 0;
    dialog->GetOptions(&opts);
    dialog->SetOptions(opts | FOS_ALLOWMULTISELECT | FOS_FORCEFILESYSTEM);
    dialog->SetTitle(L"Select files to upload");

    hr = dialog->Show(context_.ownerWindow);
    if (FAILED(hr)) return;

    ComPtr<IShellItemArray> items;
    hr = dialog->GetResults(&items);
    if (FAILED(hr)) return;

    DWORD count = 0;
    items->GetCount(&count);

    std::wstring remotePath = JoinFtpPath(context_.pathSegments);
    FtpClient client;
    bool anyFailed = false;

    for (DWORD i = 0; i < count; ++i) {
        ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(i, &item))) continue;

        PWSTR localPath = nullptr;
        if (FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &localPath))) continue;

        std::wstring local(localPath);
        CoTaskMemFree(localPath);

        // Extract filename from local path
        const wchar_t* fileName = PathFindFileNameW(local.c_str());
        std::wstring remote = remotePath + L"/" + fileName;

        FtpTransferResult result{};
        hr = client.UploadFile(context_.ftpOptions, &context_.ftpCredential,
                               local, remote, &result, context_.ownerWindow);
        if (FAILED(hr)) {
            anyFailed = true;
        }
    }

    if (anyFailed) {
        MessageBoxW(context_.ownerWindow, L"Some files failed to upload.", L"ShellTabs", MB_OK | MB_ICONWARNING);
    }

    if (context_.folderPidl) {
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

void NamespaceContextMenu::HandleOpenInBrowser() {
    if (!context_.httpUrl.empty()) {
        ShellExecuteW(context_.ownerWindow, L"open", context_.httpUrl.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }
}

void NamespaceContextMenu::HandleDelete() {
    std::wstring msg = L"Are you sure you want to delete '" + context_.itemName + L"'?";
    int answer = MessageBoxW(context_.ownerWindow, msg.c_str(), L"ShellTabs", MB_YESNO | MB_ICONQUESTION);
    if (answer != IDYES) return;

    std::wstring remotePath = JoinFtpPath(context_.pathSegments) + L"/" + context_.itemName;

    FtpClient client;
    HRESULT hr;
    if (context_.itemIsDirectory) {
        hr = client.DeleteDirectory(context_.ftpOptions, &context_.ftpCredential,
                                     remotePath, context_.ownerWindow);
    } else {
        hr = client.DeleteFile(context_.ftpOptions, &context_.ftpCredential,
                                remotePath, context_.ownerWindow);
    }

    if (FAILED(hr)) {
        MessageBoxW(context_.ownerWindow, L"Failed to delete. The server may not allow this operation.",
                    L"ShellTabs", MB_OK | MB_ICONERROR);
        return;
    }

    if (context_.folderPidl) {
        LONG eventId = context_.itemIsDirectory ? SHCNE_RMDIR : SHCNE_DELETE;
        SHChangeNotify(eventId, SHCNF_IDLIST, context_.folderPidl, nullptr);
        SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, context_.folderPidl, nullptr);
    }
}

}  // namespace shelltabs
