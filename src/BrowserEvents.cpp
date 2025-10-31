#include "BrowserEvents.h"

#include <exdispid.h>
#include <OleAuto.h>

#include <mutex>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>

#include "TabBand.h"

namespace shelltabs {

namespace {
std::wstring VariantToString(const VARIANT& var) {
    if (var.vt == VT_BSTR && var.bstrVal) {
        return std::wstring(var.bstrVal, SysStringLen(var.bstrVal));
    }
    if (var.vt == (VT_BSTR | VT_BYREF) && var.pbstrVal && *var.pbstrVal) {
        BSTR value = *var.pbstrVal;
        return std::wstring(value, SysStringLen(value));
    }
    if (var.vt == (VT_VARIANT | VT_BYREF) && var.pvarVal) {
        return VariantToString(*var.pvarVal);
    }
    return {};
}

VARIANT_BOOL* ExtractCancelPointer(VARIANT& var) {
    if (var.vt == (VT_BOOL | VT_BYREF)) {
        return var.pboolVal;
    }
    if (var.vt == (VT_VARIANT | VT_BYREF) && var.pvarVal) {
        return ExtractCancelPointer(*var.pvarVal);
    }
    return nullptr;
}

HRESULT LoadBrowserEventTypeInfo(ITypeInfo** typeInfo) {
    if (!typeInfo) {
        return E_POINTER;
    }

    static std::once_flag onceFlag;
    static Microsoft::WRL::ComPtr<ITypeInfo> cachedTypeInfo;
    static HRESULT cachedResult = E_FAIL;

    std::call_once(onceFlag, []() {
        Microsoft::WRL::ComPtr<ITypeLib> typeLib;
        HRESULT hr = LoadRegTypeLib(LIBID_SHDocVw, 1, 1, LOCALE_SYSTEM_DEFAULT, &typeLib);
        if (FAILED(hr)) {
            constexpr const wchar_t* kTypeLibs[] = {L"ieframe.dll", L"shdocvw.dll"};
            for (const auto* path : kTypeLibs) {
                Microsoft::WRL::ComPtr<ITypeLib> fallbackLib;
                hr = LoadTypeLibEx(path, REGKIND_NONE, &fallbackLib);
                if (SUCCEEDED(hr)) {
                    typeLib = fallbackLib;
                    break;
                }
            }
        }

        if (SUCCEEDED(hr)) {
            Microsoft::WRL::ComPtr<ITypeInfo> info;
            hr = typeLib->GetTypeInfoOfGuid(DIID_DWebBrowserEvents2, &info);
            if (SUCCEEDED(hr)) {
                cachedTypeInfo = info;
            }
        }

        cachedResult = hr;
    });

    if (FAILED(cachedResult)) {
        return cachedResult;
    }

    return cachedTypeInfo.CopyTo(typeInfo);
}
}  // namespace

BrowserEvents::BrowserEvents(TabBand* owner)
    : m_refCount(1), m_owner(owner) {}

IFACEMETHODIMP BrowserEvents::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IDispatch) {
        *object = static_cast<IDispatch*>(this);
        AddRef();
        return S_OK;
    }
    *object = nullptr;
    return E_NOINTERFACE;
}

IFACEMETHODIMP_(ULONG) BrowserEvents::AddRef() {
    return static_cast<ULONG>(++m_refCount);
}

IFACEMETHODIMP_(ULONG) BrowserEvents::Release() {
    const ULONG count = static_cast<ULONG>(--m_refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}

IFACEMETHODIMP BrowserEvents::GetTypeInfoCount(UINT* pctinfo) {
    if (!pctinfo) {
        return E_POINTER;
    }

    Microsoft::WRL::ComPtr<ITypeInfo> typeInfo;
    const HRESULT hr = LoadBrowserEventTypeInfo(&typeInfo);
    *pctinfo = SUCCEEDED(hr) ? 1u : 0u;
    return S_OK;
}

IFACEMETHODIMP BrowserEvents::GetTypeInfo(UINT iTInfo, LCID, ITypeInfo** ppTInfo) {
    if (!ppTInfo) {
        return E_POINTER;
    }
    if (iTInfo != 0) {
        return DISP_E_BADINDEX;
    }

    return LoadBrowserEventTypeInfo(ppTInfo);
}

IFACEMETHODIMP BrowserEvents::GetIDsOfNames(REFIID, LPOLESTR* rgszNames, UINT cNames, LCID, DISPID* rgDispId) {
    if (!rgszNames || !rgDispId) {
        return E_POINTER;
    }

    Microsoft::WRL::ComPtr<ITypeInfo> typeInfo;
    HRESULT hr = LoadBrowserEventTypeInfo(&typeInfo);
    if (FAILED(hr)) {
        return hr;
    }

    return typeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

IFACEMETHODIMP BrowserEvents::Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS* pDispParams, VARIANT*,
                                     EXCEPINFO*, UINT*) {
    if (!m_owner) {
        return S_OK;
    }

    switch (dispIdMember) {
        case DISPID_BEFORENAVIGATE2:
            if (HandleBeforeNavigate(pDispParams)) {
                return S_OK;
            }
            break;
        case DISPID_DOCUMENTCOMPLETE:
        case DISPID_NAVIGATECOMPLETE2:
            m_owner->OnBrowserNavigate();
            break;
        case DISPID_NEWWINDOW2:
        case DISPID_NEWWINDOW3:
            HandleNewWindowEvent(dispIdMember, pDispParams);
            break;
        case DISPID_ONQUIT:
            m_owner->OnBrowserQuit();
            break;
        default:
            break;
    }
    return S_OK;
}

bool BrowserEvents::HandleNewWindowEvent(DISPID dispIdMember, DISPPARAMS* params) {
    if (!m_owner || !params) {
        return false;
    }

    VARIANT_BOOL* cancel = nullptr;
    std::wstring url;

    if (dispIdMember == DISPID_NEWWINDOW3) {
        if (params->cArgs < 5) {
            return false;
        }
        cancel = ExtractCancelPointer(params->rgvarg[3]);
        url = VariantToString(params->rgvarg[0]);
    } else {
        if (params->cArgs < 2) {
            return false;
        }
        cancel = ExtractCancelPointer(params->rgvarg[0]);
    }

    if (!cancel) {
        return false;
    }

    if (m_owner->OnBrowserNewWindow(url)) {
        *cancel = VARIANT_TRUE;
        return true;
    }

    return false;
}

bool BrowserEvents::HandleBeforeNavigate(DISPPARAMS* params) {
    if (!m_owner || !params || params->cArgs < 7) {
        return false;
    }

    VARIANT_BOOL* cancel = ExtractCancelPointer(params->rgvarg[0]);
    if (!cancel) {
        return false;
    }

    const std::wstring url = VariantToString(params->rgvarg[5]);
    if (url.empty()) {
        return false;
    }

    if ((GetKeyState(VK_CONTROL) & 0x8000) == 0) {
        return false;
    }

    if (!m_owner->OnCtrlBeforeNavigate(url)) {
        return false;
    }

    *cancel = VARIANT_TRUE;
    return true;
}

HRESULT BrowserEvents::Connect(const Microsoft::WRL::ComPtr<IWebBrowser2>& browser) {
    if (!browser) {
        return E_INVALIDARG;
    }

    Microsoft::WRL::ComPtr<IConnectionPointContainer> container;
    HRESULT hr = browser.As(&container);
    if (FAILED(hr)) {
        return hr;
    }

    Microsoft::WRL::ComPtr<IConnectionPoint> connectionPoint;
    hr = container->FindConnectionPoint(DIID_DWebBrowserEvents2, &connectionPoint);
    if (FAILED(hr)) {
        return hr;
    }

    DWORD cookie = 0;
    hr = connectionPoint->Advise(this, &cookie);
    if (FAILED(hr)) {
        return hr;
    }

    m_connectionPoint = connectionPoint;
    m_cookie = cookie;
    return S_OK;
}

void BrowserEvents::Disconnect() {
    if (m_connectionPoint && m_cookie != 0) {
        m_connectionPoint->Unadvise(m_cookie);
    }
    m_connectionPoint.Reset();
    m_cookie = 0;
}

}  // namespace shelltabs

