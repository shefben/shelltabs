#pragma once

#include <windows.h>

#include <atomic>

#include <exdisp.h>
#include <ocidl.h>

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

    std::atomic<long> m_refCount;
    Microsoft::WRL::ComPtr<IUnknown> m_site;
    Microsoft::WRL::ComPtr<IWebBrowser2> m_webBrowser;
    Microsoft::WRL::ComPtr<IConnectionPoint> m_connectionPoint;
    DWORD m_connectionCookie = 0;
    bool m_bandVisible = false;
    bool m_shouldRetryEnsure = true;
};

}  // namespace shelltabs

