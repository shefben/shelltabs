#pragma once

#include "HttpClient.h"
#include "HttpPidl.h"
#include "Module.h"

#include <atomic>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include <propkey.h>
#include <shobjidl.h>
#include <wrl/client.h>

namespace shelltabs::http {

class HttpShellFolder : public IShellFolder2, public IPersistFolder2 {
public:
    HttpShellFolder();
    HttpShellFolder(const HttpUrlParts& root, const std::vector<std::wstring>& segments);
    ~HttpShellFolder();

    HttpShellFolder(const HttpShellFolder&) = delete;
    HttpShellFolder& operator=(const HttpShellFolder&) = delete;

    static HRESULT Create(const HttpUrlParts& root, const std::vector<std::wstring>& segments, REFIID riid, void** ppv);
    static HRESULT CreateWithParentPidl(const HttpUrlParts& root, const std::vector<std::wstring>& segments,
                                        PCIDLIST_ABSOLUTE parentPidl, PCUIDLIST_RELATIVE childPidl,
                                        REFIID riid, void** ppv);

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
    class ViewCallback;
    friend class ViewCallback;

    std::atomic<ULONG> refCount_{1};
    HttpUrlParts rootParts_{};
    std::vector<std::wstring> pathSegments_;
    UniquePidl absolutePidl_;
    bool initialized_ = false;
    bool isNamespaceRoot_ = false;  // True when this is the top-level "Web Folders" node
    std::wstring filterString_;
    Microsoft::WRL::ComPtr<IShellFolderViewCB> viewCallback_;

    HRESULT EnsurePidl();
    bool ExtractRelativeSegments(PCUIDLIST_RELATIVE pidl, std::vector<std::wstring>* segments, bool* isDirectory) const;
    std::wstring BuildRemotePath(const std::vector<std::wstring>& extra) const;
    HRESULT DownloadFileToStream(const std::vector<std::wstring>& segments, Microsoft::WRL::ComPtr<IStream>* stream) const;
    HRESULT BindToChild(const std::vector<std::wstring>& segments, REFIID riid, void** ppv) const;
    HRESULT EnumRootEntries(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList);
    HRESULT EnumRemoteDirectory(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList);

    // Static download queue keyed by host for parallel downloads
    static HttpDownloadQueue* GetDownloadQueue(const std::wstring& host, int maxConcurrent, int speedLimitKBps);

    static std::mutex s_queueMutex;
    static std::unordered_map<std::wstring, std::unique_ptr<HttpDownloadQueue>> s_downloadQueues;
};

}  // namespace shelltabs::http
