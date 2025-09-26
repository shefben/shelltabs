#pragma once

#include <atomic>
#include <string>

#include <exdisp.h>
#include <ocidl.h>
#include <wrl/client.h>

namespace shelltabs {

class TabBand;

class BrowserEvents : public IDispatch {
public:
    explicit BrowserEvents(TabBand* owner);

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IDispatch
    IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo) override;
    IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
    IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override;
    IFACEMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams,
                          VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override;

    HRESULT Connect(const Microsoft::WRL::ComPtr<IWebBrowser2>& browser);
    void Disconnect();

private:
    bool HandleNewWindowEvent(DISPID dispIdMember, DISPPARAMS* params);

    std::atomic<long> m_refCount;
    TabBand* m_owner;
    Microsoft::WRL::ComPtr<IConnectionPoint> m_connectionPoint;
    DWORD m_cookie = 0;
};

}  // namespace shelltabs

