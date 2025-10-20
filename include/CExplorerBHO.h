#pragma once

#include <windows.h>

#include <atomic>
#include <memory>

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
    HWND GetTopLevelExplorerWindow() const;
    bool HandleBreadcrumbPaint(HWND hwnd);
    static LRESULT CALLBACK BreadcrumbSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
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
    bool m_bufferedPaintInitialized = false;
    bool m_gdiplusInitialized = false;
    ULONG_PTR m_gdiplusToken = 0;
};

}  // namespace shelltabs

