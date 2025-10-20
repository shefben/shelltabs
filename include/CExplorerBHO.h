#pragma once

#include <windows.h>

#include <atomic>
#include <memory>
#include <string>
#include <vector>

#include <exdisp.h>
#include <ocidl.h>
#include <shlobj.h>

#include <wrl/client.h>
namespace shelltabs {

class CExplorerBHO : public IObjectWithSite, public IDispatch {
public:
    CExplorerBHO();
    ~CExplorerBHO();
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IDispatch
    IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo) override;
    IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
    IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid,
                                 DISPID* rgDispId) override;
    IFACEMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                          DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                          UINT* puArgErr) override;

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* site) override;
    IFACEMETHODIMP GetSite(REFIID riid, void** site) override;

private:
    void Disconnect();
    HRESULT EnsureBandVisible();
    HRESULT ConnectEvents();
    void DisconnectEvents();
    HRESULT ResolveBrowserFromSite(IUnknown* site, IWebBrowser2** browser);
    void UpdateBreadcrumbSubclass();
    void RemoveBreadcrumbSubclass();
    HWND FindBreadcrumbToolbar() const;
    HWND FindBreadcrumbToolbarInWindow(HWND root) const;
    HWND GetTopLevelExplorerWindow() const;
    bool InstallBreadcrumbSubclass(HWND toolbar);
    void UpdateExplorerViewSubclass();
    void RemoveExplorerViewSubclass();
    bool InstallExplorerViewSubclass(HWND viewWindow, HWND listView, HWND treeView);
    bool HandleExplorerViewMessage(HWND source, UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* result);
    void HandleExplorerContextMenuInit(HWND hwnd, HMENU menu);
    void HandleExplorerCommand(UINT commandId);
    void HandleExplorerMenuDismiss(HMENU menu);
    bool CollectSelectedFolderPaths(std::vector<std::wstring>& paths) const;
    void DispatchOpenInNewTab(const std::vector<std::wstring>& paths) const;
    void ClearPendingOpenInNewTabState();
    void EnsureBreadcrumbHook();
    void RemoveBreadcrumbHook();
    bool IsBreadcrumbToolbarCandidate(HWND hwnd) const;
    bool IsBreadcrumbToolbarAncestor(HWND hwnd) const;
    bool IsWindowOwnedByThisExplorer(HWND hwnd) const;
    bool HandleBreadcrumbPaint(HWND hwnd);
    enum class BreadcrumbDiscoveryStage {
        None,
        ServiceUnavailable,
        ServiceWindowMissing,
        ServiceToolbarMissing,
        FrameMissing,
        RebarMissing,
        ParentMissing,
        ToolbarMissing,
        Discovered,
    };
    void LogBreadcrumbStage(BreadcrumbDiscoveryStage stage, const wchar_t* format, ...) const;
    static LRESULT CALLBACK BreadcrumbCbtProc(int code, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK BreadcrumbSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR subclassId, DWORD_PTR refData);
    static LRESULT CALLBACK ExplorerViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                     UINT_PTR subclassId, DWORD_PTR refData);

    std::atomic<long> m_refCount;
    Microsoft::WRL::ComPtr<IUnknown> m_site;
    Microsoft::WRL::ComPtr<IWebBrowser2> m_webBrowser;
    Microsoft::WRL::ComPtr<IConnectionPoint> m_connectionPoint;
    DWORD m_connectionCookie = 0;
    bool m_bandVisible = false;
    bool m_shouldRetryEnsure = true;
    Microsoft::WRL::ComPtr<IShellBrowser> m_shellBrowser;
    HWND m_breadcrumbToolbar = nullptr;
    bool m_breadcrumbSubclassInstalled = false;
    bool m_breadcrumbGradientEnabled = false;
    bool m_breadcrumbFontGradientEnabled = false;
    int m_breadcrumbGradientTransparency = 45;
    int m_breadcrumbFontTransparency = 0;
    bool m_breadcrumbHookRegistered = false;
    enum class BreadcrumbLogState {
        Unknown,
        Disabled,
        Searching,
    };
    BreadcrumbLogState m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    bool m_loggedBreadcrumbToolbarMissing = false;
    bool m_bufferedPaintInitialized = false;
    bool m_gdiplusInitialized = false;
    ULONG_PTR m_gdiplusToken = 0;
    mutable BreadcrumbDiscoveryStage m_lastBreadcrumbStage = BreadcrumbDiscoveryStage::None;
    Microsoft::WRL::ComPtr<IShellView> m_shellView;
    HWND m_shellViewWindow = nullptr;
    bool m_shellViewWindowSubclassInstalled = false;
    HWND m_listView = nullptr;
    HWND m_treeView = nullptr;
    bool m_listViewSubclassInstalled = false;
    bool m_treeViewSubclassInstalled = false;
    HMENU m_trackedContextMenu = nullptr;
    std::vector<std::wstring> m_pendingOpenInNewTabPaths;
    bool m_contextMenuInserted = false;
    static constexpr UINT kOpenInNewTabCommandId = 0xE170;
    static constexpr UINT kMaxTrackedSelection = 16;
};

}  // namespace shelltabs

