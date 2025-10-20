#include "CExplorerBHO.h"

#include <combaseapi.h>
#include <objbase.h>
#include <exdispid.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shlguid.h>

#include <string>
#include <wrl/client.h>

#include "ComUtils.h"
#include "Guids.h"
#include "Module.h"
#include "Utilities.h"
// --- CExplorerBHO private state (treat these as class members) ---

namespace {

bool ExtractVariantBool(const VARIANT& value, bool* result) {
    if (!result) {
        return false;
    }

    switch (value.vt) {
        case VT_BOOL:
            *result = (value.boolVal != VARIANT_FALSE);
            return true;
        case VT_BOOL | VT_BYREF:
            if (!value.pboolVal) {
                return false;
            }
            *result = (*value.pboolVal != VARIANT_FALSE);
            return true;
        case VT_VARIANT | VT_BYREF:
            if (!value.pvarVal) {
                return false;
            }
            return ExtractVariantBool(*value.pvarVal, result);
        default:
            break;
    }

    return false;
}

bool ExtractVariantLong(const VARIANT& value, LONG* result) {
    if (!result) {
        return false;
    }

    switch (value.vt) {
        case VT_I4:
        case VT_INT:
            *result = value.lVal;
            return true;
        case VT_UI4:
        case VT_UINT:
            *result = static_cast<LONG>(value.ulVal);
            return true;
        case VT_I2:
            *result = value.iVal;
            return true;
        case VT_UI2:
            *result = value.uiVal;
            return true;
        case VT_BOOL:
            *result = (value.boolVal != VARIANT_FALSE) ? 1 : 0;
            return true;
        case VT_I4 | VT_BYREF:
        case VT_INT | VT_BYREF:
            if (!value.plVal) {
                return false;
            }
            *result = *value.plVal;
            return true;
        case VT_UI4 | VT_BYREF:
        case VT_UINT | VT_BYREF:
            if (!value.pulVal) {
                return false;
            }
            *result = static_cast<LONG>(*value.pulVal);
            return true;
        case VT_I2 | VT_BYREF:
            if (!value.piVal) {
                return false;
            }
            *result = *value.piVal;
            return true;
        case VT_UI2 | VT_BYREF:
            if (!value.puiVal) {
                return false;
            }
            *result = *value.puiVal;
            return true;
        case VT_BOOL | VT_BYREF:
            if (!value.pboolVal) {
                return false;
            }
            *result = (*value.pboolVal != VARIANT_FALSE) ? 1 : 0;
            return true;
        case VT_VARIANT | VT_BYREF:
            if (!value.pvarVal) {
                return false;
            }
            return ExtractVariantLong(*value.pvarVal, result);
        default:
            break;
    }

    return false;
}

bool ExtractVariantDispatch(const VARIANT& value, IDispatch** result) {
    if (!result) {
        return false;
    }

    *result = nullptr;

    switch (value.vt) {
        case VT_DISPATCH:
            if (!value.pdispVal) {
                return false;
            }
            *result = value.pdispVal;
            (*result)->AddRef();
            return true;
        case VT_UNKNOWN:
            if (!value.punkVal) {
                return false;
            }
            return SUCCEEDED(value.punkVal->QueryInterface(IID_PPV_ARGS(result)));
        case VT_DISPATCH | VT_BYREF:
            if (!value.ppdispVal || !*value.ppdispVal) {
                return false;
            }
            *result = *value.ppdispVal;
            (*result)->AddRef();
            return true;
        case VT_UNKNOWN | VT_BYREF:
            if (!value.ppunkVal || !*value.ppunkVal) {
                return false;
            }
            return SUCCEEDED((*value.ppunkVal)->QueryInterface(IID_PPV_ARGS(result)));
        case VT_VARIANT | VT_BYREF:
            if (!value.pvarVal) {
                return false;
            }
            return ExtractVariantDispatch(*value.pvarVal, result);
        default:
            break;
    }

    return false;
}

#ifndef WBSTATE_USERVISIBLE
#define WBSTATE_USERVISIBLE 0x00000001
#endif

}  // namespace

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
    if (riid == IID_IUnknown || riid == IID_IObjectWithSite) {
        *object = static_cast<IObjectWithSite*>(this);
    } else if (riid == IID_IDispatch) {
        *object = static_cast<IDispatch*>(this);
    } else {
        *object = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
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

void CExplorerBHO::Disconnect() {
    DisconnectEvents();
    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();
    ResetEnsureState();
}


void CExplorerBHO::ResetEnsureState() {
    m_bandVisible = false;
    m_shouldRetryEnsure = true;
}


HRESULT CExplorerBHO::EnsureBandVisible() {
    return GuardExplorerCall(
        L"CExplorerBHO::EnsureBandVisible",
        [&]() -> HRESULT {
            if (!m_webBrowser || !m_shouldRetryEnsure) {
                return S_OK;
            }

            Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
            HRESULT hr = m_webBrowser.As(&serviceProvider);
            if ((!serviceProvider || FAILED(hr)) && m_site) {
                serviceProvider = nullptr;
                hr = m_site.As(&serviceProvider);
            }
            if (FAILED(hr) || !serviceProvider) {
                return hr;
            }

            Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
            hr = serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&shellBrowser));
            if ((FAILED(hr) || !shellBrowser)) {
                hr = serviceProvider->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&shellBrowser));
                if (FAILED(hr) || !shellBrowser) {
                    return hr;
                }
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

            if (SUCCEEDED(hr)) {
                m_bandVisible = true;
                m_shouldRetryEnsure = false;
            } else if (hr == E_ACCESSDENIED || HRESULT_CODE(hr) == ERROR_ACCESS_DENIED) {
                m_shouldRetryEnsure = false;
            }

            return hr;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP CExplorerBHO::SetSite(IUnknown* site) {
    return GuardExplorerCall(
        L"CExplorerBHO::SetSite",
        [&]() -> HRESULT {
            if (!site) {
                Disconnect();
                return S_OK;
            }

            Disconnect();

            Microsoft::WRL::ComPtr<IWebBrowser2> browser;
            HRESULT hr = ResolveBrowserFromSite(site, &browser);
            if (FAILED(hr) || !browser) {
                return S_OK;
            }

            m_site = site;
            m_webBrowser = browser;
            m_shouldRetryEnsure = true;

            m_shellBrowser.Reset();

            ConnectEvents();

            Microsoft::WRL::ComPtr<IServiceProvider> siteProvider;
            if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&siteProvider))) && siteProvider) {
                if (!m_shellBrowser) {
                    siteProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_shellBrowser));
                }
                if (!m_shellBrowser) {
                    siteProvider->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&m_shellBrowser));
                }
            } else {
                Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
                if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&shellBrowser))) && shellBrowser) {
                    m_shellBrowser = shellBrowser;
                    m_shellBrowser.As(&siteProvider);
                }
            }

            if (!m_shellBrowser) {
                Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
                if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&shellBrowser))) && shellBrowser) {
                    m_shellBrowser = shellBrowser;
                }
            }

            if (!siteProvider && m_shellBrowser) {
                m_shellBrowser.As(&siteProvider);
            }
            EnsureBandVisible();
            return S_OK;

        },
        []() -> HRESULT { return E_FAIL; });
}

HRESULT CExplorerBHO::ResolveBrowserFromSite(IUnknown* site, IWebBrowser2** browser) {
    if (!browser) {
        return E_POINTER;
    }

    *browser = nullptr;

    if (!site) {
        return E_POINTER;
    }

    Microsoft::WRL::ComPtr<IWebBrowser2> candidate;
    HRESULT hr = site->QueryInterface(IID_PPV_ARGS(&candidate));
    if (SUCCEEDED(hr) && candidate) {
        *browser = candidate.Detach();
        return S_OK;
    }

    Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
    hr = site->QueryInterface(IID_PPV_ARGS(&serviceProvider));
    if (SUCCEEDED(hr) && serviceProvider) {
        hr = serviceProvider->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&candidate));
        if (SUCCEEDED(hr) && candidate) {
            *browser = candidate.Detach();
            return S_OK;
        }

        hr = serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&candidate));
        if (SUCCEEDED(hr) && candidate) {
            *browser = candidate.Detach();
            return S_OK;
        }
    }

    Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
    hr = site->QueryInterface(IID_PPV_ARGS(&shellBrowser));
    if (SUCCEEDED(hr) && shellBrowser) {
        serviceProvider = nullptr;
        hr = shellBrowser.As(&serviceProvider);
        if (SUCCEEDED(hr) && serviceProvider) {
            hr = serviceProvider->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&candidate));
            if (SUCCEEDED(hr) && candidate) {
                *browser = candidate.Detach();
                return S_OK;
            }
        }
    }

    return E_NOINTERFACE;
}

IFACEMETHODIMP CExplorerBHO::GetSite(REFIID riid, void** site) {
    return GuardExplorerCall(
        L"CExplorerBHO::GetSite",
        [&]() -> HRESULT {
            if (!site) {
                return E_POINTER;
            }
            *site = nullptr;
            if (!m_site) {
                return E_FAIL;
            }
            return m_site->QueryInterface(riid, site);
        },
        []() -> HRESULT { return E_FAIL; });
}

HRESULT CExplorerBHO::ConnectEvents() {
    return GuardExplorerCall(
        L"CExplorerBHO::ConnectEvents",
        [&]() -> HRESULT {
            if (!m_webBrowser || m_connectionCookie != 0) {
                return S_OK;
            }

            Microsoft::WRL::ComPtr<IConnectionPointContainer> container;
            HRESULT hr = m_webBrowser.As(&container);
            if (FAILED(hr) || !container) {
                return hr;
            }

            Microsoft::WRL::ComPtr<IConnectionPoint> connectionPoint;
            hr = container->FindConnectionPoint(DIID_DWebBrowserEvents2, &connectionPoint);
            if (FAILED(hr) || !connectionPoint) {
                return hr;
            }

            DWORD cookie = 0;
            hr = connectionPoint->Advise(static_cast<IDispatch*>(this), &cookie);
            if (FAILED(hr)) {
                return hr;
            }

            m_connectionPoint = connectionPoint;
            m_connectionCookie = cookie;
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

void CExplorerBHO::DisconnectEvents() {
    if (m_connectionPoint && m_connectionCookie != 0) {
        m_connectionPoint->Unadvise(m_connectionCookie);
    }
    m_connectionPoint.Reset();
    m_connectionCookie = 0;
}

IFACEMETHODIMP CExplorerBHO::Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS* params, VARIANT*, EXCEPINFO*,
                                    UINT*) {
    return GuardExplorerCall(
        L"CExplorerBHO::Invoke",
        [&]() -> HRESULT {
            switch (dispIdMember) {
                case DISPID_ONVISIBLE: {
                    if (!params || params->cArgs < 1) {
                        if (m_shouldRetryEnsure || !m_bandVisible) {
                            EnsureBandVisible();
                        }
                        break;
                    }
                    bool visible = false;
                    if (!ExtractVariantBool(params->rgvarg[0], &visible)) {
                        if (m_shouldRetryEnsure || !m_bandVisible) {
                            EnsureBandVisible();
                        }
                        break;
                    }
                    if (!visible) {
                        ResetEnsureState();
                        break;
                    }
                    if (!m_bandVisible || m_shouldRetryEnsure) {
                        EnsureBandVisible();
                    }
                    break;
                }
                case DISPID_WINDOWSTATECHANGED: {
                    if (!params || params->cArgs < 2) {
                        if (m_shouldRetryEnsure || !m_bandVisible) {
                            EnsureBandVisible();
                        }
                        break;
                    }
                    LONG stateFlags = 0;
                    LONG validMask = 0;
                    const bool haveState = ExtractVariantLong(params->rgvarg[1], &stateFlags);
                    const bool haveMask = ExtractVariantLong(params->rgvarg[0], &validMask);
                    if (haveState && haveMask && (validMask & WBSTATE_USERVISIBLE) != 0) {
                        if ((stateFlags & WBSTATE_USERVISIBLE) == 0) {
                            ResetEnsureState();
                            break;
                        }
                        if (!m_bandVisible || m_shouldRetryEnsure) {
                            EnsureBandVisible();
                        }
                    } else if (!m_bandVisible || m_shouldRetryEnsure) {
                        EnsureBandVisible();
                    }
                    break;
                }
                case DISPID_NAVIGATECOMPLETE2:
                case DISPID_DOCUMENTCOMPLETE: {
                    if (!params || params->cArgs < 2) {
                        if (m_shouldRetryEnsure || !m_bandVisible) {
                            EnsureBandVisible();
                        }
                        break;
                    }

                    Microsoft::WRL::ComPtr<IDispatch> dispatch;
                    if (!ExtractVariantDispatch(params->rgvarg[1], dispatch.GetAddressOf()) || !dispatch) {
                        if (m_shouldRetryEnsure || !m_bandVisible) {
                            EnsureBandVisible();
                        }
                        break;
                    }

                    Microsoft::WRL::ComPtr<IWebBrowser2> sourceBrowser;
                    if (FAILED(dispatch.As(&sourceBrowser)) || !sourceBrowser) {
                        if (m_shouldRetryEnsure || !m_bandVisible) {
                            EnsureBandVisible();
                        }
                        break;
                    }

                    if (m_webBrowser && IsEqualObject(sourceBrowser.Get(), m_webBrowser.Get())) {
                        if (!m_bandVisible || m_shouldRetryEnsure) {
                            EnsureBandVisible();
                        }
                    }
                    break;
                }
                case DISPID_ONQUIT:
                    Disconnect();
                    break;
                default:
                    break;
            }

            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

}  // namespace shelltabs

