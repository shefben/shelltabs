#include "CExplorerBHO.h"

#include <combaseapi.h>
#include <oleauto.h>
#include <shlobj.h>

#include <string>

#include "ComUtils.h"
#include "Guids.h"
#include "Module.h"

namespace shelltabs {

CExplorerBHO::CExplorerBHO() : m_refCount(1) {
    ModuleAddRef();
}

CExplorerBHO::~CExplorerBHO() {
    Disconnect();
    ModuleRelease();
}

IFACEMETHODIMP CExplorerBHO::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IObjectWithSite || riid == IID_IDispatch) {
        *object = static_cast<IObjectWithSite*>(this);
        AddRef();
        return S_OK;
    }
    *object = nullptr;
    return E_NOINTERFACE;
}

IFACEMETHODIMP_(ULONG) CExplorerBHO::AddRef() {
    return static_cast<ULONG>(++m_refCount);
}

IFACEMETHODIMP_(ULONG) CExplorerBHO::Release() {
    const ULONG count = static_cast<ULONG>(--m_refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}

IFACEMETHODIMP CExplorerBHO::GetTypeInfoCount(UINT* pctinfo) {
    if (pctinfo) {
        *pctinfo = 0;
    }
    return S_OK;
}

IFACEMETHODIMP CExplorerBHO::GetTypeInfo(UINT, LCID, ITypeInfo**) {
    return E_NOTIMPL;
}

IFACEMETHODIMP CExplorerBHO::GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) {
    return E_NOTIMPL;
}

IFACEMETHODIMP CExplorerBHO::Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) {
    return S_OK;
}

void CExplorerBHO::Disconnect() {
    m_webBrowser.Reset();
    m_site.Reset();
    m_hasAttemptedEnsure = false;
}

HRESULT CExplorerBHO::EnsureBandVisible() {
    if (m_hasAttemptedEnsure || !m_webBrowser) {
        return S_OK;
    }

    m_hasAttemptedEnsure = true;

    Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
    HRESULT hr = m_webBrowser.As(&serviceProvider);
    if (FAILED(hr) || !serviceProvider) {
        return hr;
    }

    Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
    hr = serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&shellBrowser));
    if (FAILED(hr) || !shellBrowser) {
        return hr;
    }

    const std::wstring clsidString = GuidToString(CLSID_ShellTabsBand);
    if (clsidString.empty()) {
        return E_FAIL;
    }

    VARIANT bandId;
    VariantInit(&bandId);
    bandId.vt = VT_BSTR;
    bandId.bstrVal = SysAllocString(clsidString.c_str());
    if (!bandId.bstrVal) {
        return E_OUTOFMEMORY;
    }

    VARIANT show;
    VariantInit(&show);
    show.vt = VT_BOOL;
    show.boolVal = VARIANT_TRUE;

    hr = m_webBrowser->ShowBrowserBar(&bandId, &show, nullptr);

    VariantClear(&bandId);
    VariantClear(&show);
    return hr;
}

IFACEMETHODIMP CExplorerBHO::SetSite(IUnknown* site) {
    if (!site) {
        Disconnect();
        return S_OK;
    }

    Microsoft::WRL::ComPtr<IWebBrowser2> browser;
    HRESULT hr = site->QueryInterface(IID_PPV_ARGS(&browser));
    if (FAILED(hr) || !browser) {
        Disconnect();
        return S_OK;
    }

    m_site = site;
    m_webBrowser = browser;

    // Attempt to surface the deskband for the current window. Ignore failures because Explorer may
    // reject the call if policy forbids bands or the window is not ready yet.
    EnsureBandVisible();
    return S_OK;
}

IFACEMETHODIMP CExplorerBHO::GetSite(REFIID riid, void** site) {
    if (!site) {
        return E_POINTER;
    }
    *site = nullptr;
    if (!m_site) {
        return E_FAIL;
    }
    return m_site->QueryInterface(riid, site);
}

}  // namespace shelltabs

