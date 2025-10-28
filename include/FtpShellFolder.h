#pragma once

#include "FtpClient.h"
#include "FtpPidl.h"
#include "Module.h"

#include <atomic>
#include <string>
#include <string_view>
#include <vector>

#include <propkey.h>
#include <shobjidl.h>
#include <wrl/client.h>

namespace shelltabs::ftp {

class FtpShellFolder : public IShellFolder2, public IPersistFolder2 {
public:
    FtpShellFolder();
    FtpShellFolder(const FtpUrlParts& root, const std::vector<std::wstring>& segments);
    ~FtpShellFolder();

    FtpShellFolder(const FtpShellFolder&) = delete;
    FtpShellFolder& operator=(const FtpShellFolder&) = delete;

    static HRESULT Create(const FtpUrlParts& root, const std::vector<std::wstring>& segments, REFIID riid, void** ppv);

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IShellFolder
    IFACEMETHODIMP ParseDisplayName(HWND hwnd, IBindCtx* pbc, PWSTR pszName, ULONG* pchEaten, PIDLIST_RELATIVE* ppidl,
                                   ULONG* pdwAttributes) override;
    IFACEMETHODIMP EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList) override;
    IFACEMETHODIMP BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx* pbc, REFIID riid, void** ppv) override;
    IFACEMETHODIMP BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx* pbc, REFIID riid, void** ppv) override;
    IFACEMETHODIMP CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) override;
    IFACEMETHODIMP CreateViewObject(HWND hwnd, REFIID riid, void** ppv) override;
    IFACEMETHODIMP GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, ULONG* rgfInOut) override;
    IFACEMETHODIMP GetUIObjectOf(HWND hwnd, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid, UINT* prgfInOut,
                                 void** ppv) override;
    IFACEMETHODIMP GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET* pName) override;
    IFACEMETHODIMP SetNameOf(HWND hwnd, PCUITEMID_CHILD pidl, PCWSTR pszName, SHGDNF uFlags,
                             PIDLIST_RELATIVE* ppidlOut) override;

    // IShellFolder2
    IFACEMETHODIMP GetDefaultSearchGUID(GUID* pguid) override;
    IFACEMETHODIMP EnumSearches(IEnumExtraSearch** ppEnum) override;
    IFACEMETHODIMP GetDefaultColumn(DWORD dwRes, ULONG* pSort, ULONG* pDisplay) override;
    IFACEMETHODIMP GetDefaultColumnState(UINT iColumn, SHCOLSTATEF* pcsFlags) override;
    IFACEMETHODIMP GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID* pscid, VARIANT* pv) override;
    IFACEMETHODIMP GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS* pDetails) override;
    IFACEMETHODIMP MapColumnToSCID(UINT iColumn, SHCOLUMNID* pscid) override;

    // IPersist
    IFACEMETHODIMP GetClassID(CLSID* pClassID) override;

    // IPersistFolder
    IFACEMETHODIMP Initialize(PCIDLIST_ABSOLUTE pidl) override;

    // IPersistFolder2
    IFACEMETHODIMP GetCurFolder(PIDLIST_ABSOLUTE* ppidl) override;

private:
    std::atomic<ULONG> refCount_{1};
    FtpUrlParts rootParts_{};
    std::vector<std::wstring> pathSegments_;
    UniquePidl absolutePidl_;
    bool initialized_ = false;

    HRESULT EnsurePidl();
    HRESULT ParseInputToSegments(std::wstring_view input, std::vector<std::wstring>* segments, bool* isDirectory) const;
    bool ExtractRelativeSegments(PCUIDLIST_RELATIVE pidl, std::vector<std::wstring>* segments, bool* isDirectory) const;
    std::wstring BuildFolderPath(const std::vector<std::wstring>& extra) const;
    std::wstring BuildFileName(const std::vector<std::wstring>& segments, std::wstring* directoryOut) const;
    HRESULT DownloadFileToStream(const std::vector<std::wstring>& segments, Microsoft::WRL::ComPtr<IStream>* stream) const;
    HRESULT BindToChild(const std::vector<std::wstring>& segments, REFIID riid, void** ppv) const;
};

}  // namespace shelltabs::ftp

