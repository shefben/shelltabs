#include "CExplorerBHO.h"

#include <combaseapi.h>
#include <exdispid.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shlguid.h>
#include <CommCtrl.h>
#include <uxtheme.h>
#include <OleIdl.h>

#include <array>
#include <cwchar>
#include <string>
#include <wrl/client.h>

#include "ComUtils.h"
#include "Guids.h"
#include "Logging.h"
#include "Module.h"
#include "OptionsStore.h"
#include "Utilities.h"

namespace {

bool MatchesClass(HWND hwnd, const wchar_t* className) {
    if (!hwnd || !className) {
        return false;
    }
    wchar_t buffer[256];
    const int length = GetClassNameW(hwnd, buffer, ARRAYSIZE(buffer));
    if (length <= 0) {
        return false;
    }
    return _wcsicmp(buffer, className) == 0;
}

HWND FindDescendantWindow(HWND parent, const wchar_t* className) {
    if (!parent || !className) {
        return nullptr;
    }
    for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT)) {
        if (MatchesClass(child, className)) {
            return child;
        }
        if (HWND found = FindDescendantWindow(child, className)) {
            return found;
        }
    }
    return nullptr;
}

}  // namespace
// --- CExplorerBHO private state (treat these as class members) ---

namespace shelltabs {

CExplorerBHO::CExplorerBHO() : m_refCount(1) {
    ModuleAddRef();
    m_bufferedPaintInitialized = SUCCEEDED(BufferedPaintInit());
}

CExplorerBHO::~CExplorerBHO() {
    Disconnect();
    if (m_bufferedPaintInitialized) {
        BufferedPaintUnInit();
        m_bufferedPaintInitialized = false;
    }
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
    RemoveBreadcrumbSubclass();
    DisconnectEvents();
    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();
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
                UpdateBreadcrumbSubclass();
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
            UpdateBreadcrumbSubclass();
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

IFACEMETHODIMP CExplorerBHO::Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*,
                                    UINT*) {
    return GuardExplorerCall(
        L"CExplorerBHO::Invoke",
        [&]() -> HRESULT {
            switch (dispIdMember) {
                case DISPID_ONVISIBLE:
                case DISPID_WINDOWSTATECHANGED:
                    if (!m_bandVisible) {
                        EnsureBandVisible();
                        UpdateBreadcrumbSubclass();
                    }
                    break;
                case DISPID_DOCUMENTCOMPLETE:
                case DISPID_NAVIGATECOMPLETE2:
                    UpdateBreadcrumbSubclass();
                    break;
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

HWND CExplorerBHO::GetTopLevelExplorerWindow() const {
    HWND hwnd = nullptr;
    if (m_shellBrowser && SUCCEEDED(m_shellBrowser->GetWindow(&hwnd)) && hwnd) {
        return hwnd;
    }
    if (m_webBrowser) {
        SHANDLE_PTR raw = 0;
        if (SUCCEEDED(m_webBrowser->get_HWND(&raw)) && raw) {
            return reinterpret_cast<HWND>(raw);
        }
    }
    return nullptr;
}

HWND CExplorerBHO::FindBreadcrumbToolbar() const {
    auto queryBreadcrumbToolbar = [&](const Microsoft::WRL::ComPtr<IServiceProvider>& provider) -> HWND {
        if (!provider) {
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
        if (FAILED(provider->QueryService(CLSID_CBreadcrumbBar, IID_PPV_ARGS(&oleWindow))) || !oleWindow) {
            return nullptr;
        }

        HWND bandWindow = nullptr;
        if (FAILED(oleWindow->GetWindow(&bandWindow)) || !bandWindow) {
            return nullptr;
        }

        HWND toolbar = FindWindowExW(bandWindow, nullptr, TOOLBARCLASSNAME, nullptr);
        if (!toolbar) {
            toolbar = FindDescendantWindow(bandWindow, TOOLBARCLASSNAME);
        }
        return toolbar;
    };

    if (m_shellBrowser) {
        Microsoft::WRL::ComPtr<IServiceProvider> provider;
        if (SUCCEEDED(m_shellBrowser.As(&provider))) {
            if (HWND fromService = queryBreadcrumbToolbar(provider)) {
                return fromService;
            }
        }
    }

    if (m_webBrowser) {
        Microsoft::WRL::ComPtr<IServiceProvider> provider;
        if (SUCCEEDED(m_webBrowser.As(&provider))) {
            if (HWND fromService = queryBreadcrumbToolbar(provider)) {
                return fromService;
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return nullptr;
    }

    HWND travelBand = FindDescendantWindow(frame, L"TravelBand");
    if (!travelBand) {
        return nullptr;
    }

    HWND rebar = GetParent(travelBand);
    if (!rebar) {
        return nullptr;
    }

    HWND breadcrumbParent = FindWindowExW(rebar, nullptr, L"Breadcrumb Parent", nullptr);
    if (!breadcrumbParent) {
        breadcrumbParent = FindDescendantWindow(rebar, L"Breadcrumb Parent");
    }
    if (!breadcrumbParent) {
        return nullptr;
    }

    HWND toolbar = FindWindowExW(breadcrumbParent, nullptr, TOOLBARCLASSNAME, nullptr);
    if (!toolbar) {
        toolbar = FindDescendantWindow(breadcrumbParent, TOOLBARCLASSNAME);
    }
    return toolbar;
}

void CExplorerBHO::RemoveBreadcrumbSubclass() {
    if (m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        if (IsWindow(m_breadcrumbToolbar)) {
            RemoveWindowSubclass(m_breadcrumbToolbar, BreadcrumbSubclassProc, reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_breadcrumbToolbar, nullptr, TRUE);
        }
    }
    m_breadcrumbToolbar = nullptr;
    m_breadcrumbSubclassInstalled = false;
}

void CExplorerBHO::UpdateBreadcrumbSubclass() {
    auto& store = OptionsStore::Instance();
    store.Load();
    const ShellTabsOptions options = store.Get();
    m_breadcrumbGradientEnabled = options.enableBreadcrumbGradient;

    if (!m_breadcrumbGradientEnabled) {
        if (m_breadcrumbSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Breadcrumb gradient disabled; removing subclass");
        }
        RemoveBreadcrumbSubclass();
        return;
    }

    HWND toolbar = FindBreadcrumbToolbar();
    if (!toolbar) {
        if (m_breadcrumbSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Breadcrumb toolbar not found; removing subclass");
        }
        RemoveBreadcrumbSubclass();
        return;
    }

    if (toolbar == m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        InvalidateRect(toolbar, nullptr, TRUE);
        return;
    }

    RemoveBreadcrumbSubclass();

    if (SetWindowSubclass(toolbar, BreadcrumbSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_breadcrumbToolbar = toolbar;
        m_breadcrumbSubclassInstalled = true;
        LogMessage(LogLevel::Info, L"Installed breadcrumb gradient subclass on hwnd=%p", toolbar);
        InvalidateRect(toolbar, nullptr, TRUE);
    }
}

bool CExplorerBHO::HandleBreadcrumbPaint(HWND hwnd) {
    if (!m_breadcrumbGradientEnabled) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC target = BeginPaint(hwnd, &ps);
    if (!target) {
        return true;
    }

    RECT client{};
    if (!GetClientRect(hwnd, &client)) {
        EndPaint(hwnd, &ps);
        return true;
    }

    BP_PAINTPARAMS params{};
    params.cbSize = sizeof(params);
    params.dwFlags = BPPF_ERASE;
    HDC paintDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(target, &client, BPBF_TOPDOWNDIB, &params, &paintDc);
    if (!buffer || !paintDc) {
        DefSubclassProc(hwnd, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(target), PRF_CLIENT);
        if (buffer) {
            EndBufferedPaint(buffer, FALSE);
        }
        EndPaint(hwnd, &ps);
        return true;
    }

    DefSubclassProc(hwnd, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(paintDc), PRF_CLIENT);

    HFONT font = reinterpret_cast<HFONT>(SendMessage(hwnd, WM_GETFONT, 0, 0));
    HFONT oldFont = nullptr;
    if (font) {
        oldFont = static_cast<HFONT>(SelectObject(paintDc, font));
    }

    SetBkMode(paintDc, TRANSPARENT);

    constexpr std::array<COLORREF, 7> kRainbowColors = {
        RGB(255, 59, 48),   // red
        RGB(255, 149, 0),   // orange
        RGB(255, 204, 0),   // yellow
        RGB(52, 199, 89),   // green
        RGB(0, 122, 255),   // blue
        RGB(88, 86, 214),   // indigo
        RGB(175, 82, 222)   // violet
    };

    const int buttonCount = static_cast<int>(SendMessage(hwnd, TB_BUTTONCOUNT, 0, 0));
    int colorIndex = 0;
    for (int i = 0; i < buttonCount; ++i) {
        TBBUTTON button{};
        if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&button))) {
            continue;
        }
        if ((button.fsStyle & TBSTYLE_SEP) != 0 || (button.fsState & TBSTATE_HIDDEN) != 0) {
            continue;
        }

        LRESULT textLength = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand, 0);
        if (textLength <= 0) {
            continue;
        }

        std::wstring text(static_cast<size_t>(textLength) + 1, L'\0');
        LRESULT copied = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand, reinterpret_cast<LPARAM>(text.data()));
        if (copied <= 0) {
            continue;
        }
        text.resize(static_cast<size_t>(copied));
        if (text.empty()) {
            continue;
        }

        RECT itemRect{};
        if (!SendMessage(hwnd, TB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&itemRect))) {
            continue;
        }

        const COLORREF color = kRainbowColors[static_cast<size_t>(colorIndex % kRainbowColors.size())];
        ++colorIndex;
        SetTextColor(paintDc, color);

        RECT textRect = itemRect;
        textRect.left += 4;
        textRect.right -= 4;

        DrawTextW(paintDc, text.c_str(), static_cast<int>(text.size()), &textRect,
                  DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
    }

    if (oldFont) {
        SelectObject(paintDc, oldFont);
    }

    EndBufferedPaint(buffer, TRUE);
    EndPaint(hwnd, &ps);
    return true;
}

LRESULT CALLBACK CExplorerBHO::BreadcrumbSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                      UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_PAINT:
            if (self->HandleBreadcrumbPaint(hwnd)) {
                return 0;
            }
            break;
        case WM_NCDESTROY:
            self->RemoveBreadcrumbSubclass();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

}  // namespace shelltabs

