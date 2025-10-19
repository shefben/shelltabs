#include "CExplorerBHO.h"

#include <combaseapi.h>
#include <exdispid.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shlguid.h>

#include <string>
#include <CommCtrl.h>     // SetWindowSubclass, RemoveWindowSubclass
#include <shobjidl.h>
#include <wrl/client.h>

#include "SplitHost.h"    // your new split host header
#include "TabManager.h"   // for reading split state (group/tab indices)

#ifndef WM_SHELLTABS_APPLY_SPLIT
#define WM_SHELLTABS_APPLY_SPLIT (WM_APP + 0x3F1)
#endif

#include "ComUtils.h"
#include "CommonDialogColorizer.h"
#include "ExplorerWindowHook.h"
#include "Guids.h"
#include "NamespaceTreeColorizer.h"
#include "Module.h"
#include "Utilities.h"
#include "FileColorOverrides.h"
// --- CExplorerBHO private state (treat these as class members) ---

namespace shelltabs {

CExplorerBHO::CExplorerBHO() : m_refCount(1) {
    ModuleAddRef();
}

CExplorerBHO::~CExplorerBHO() {
    Disconnect();
    ModuleRelease();
}
// Helper to get the Explorer frame HWND from IWebBrowser2
static HWND GetExplorerFrameHwnd(IWebBrowser2* wb) {
	SHANDLE_PTR h = 0;
	if (wb && SUCCEEDED(wb->get_HWND(&h)) && h) {
		return reinterpret_cast<HWND>(h);
	}
	return nullptr;
}

// Find DefView and content parent
static HWND FindDefViewParent(HWND shellBrowserHwnd) {
	UNREFERENCED_PARAMETER(shellBrowserHwnd);  // fixes C4100
	return nullptr;
}
static HWND FindDefViewParentFromBrowser(IServiceProvider* sp, HWND* outDefView) {
	if (outDefView) *outDefView = nullptr;
	if (!sp) return nullptr;

	Microsoft::WRL::ComPtr<IShellBrowser> browser;
	if (FAILED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser))) || !browser)
		return nullptr;

	Microsoft::WRL::ComPtr<IShellView> view;
	if (FAILED(browser->QueryActiveShellView(&view)) || !view)
		return nullptr;

	HWND hwndDefView = nullptr;
	if (FAILED(view->GetWindow(&hwndDefView)) || !IsWindow(hwndDefView))
		return nullptr;

	if (outDefView) *outDefView = hwndDefView;
	return GetParent(hwndDefView);
}

// near top: keep the WM id you already defined
#ifndef WM_SHELLTABS_APPLY_SPLIT
#define WM_SHELLTABS_APPLY_SPLIT (WM_APP + 0x3F1)
#endif
// --- Implement teardown so we don't leave DefView hidden ---
void CExplorerBHO::RemoveSplitIfAny() {
        if (m_splitHost && IsWindow(m_splitHost)) {
                DestroyWindow(m_splitHost);
        }
        m_splitHost = nullptr;
        if (m_defView && IsWindow(m_defView)) {
                ShowWindow(m_defView, SW_SHOW);
        }
        FileColorOverrides::Instance().ClearEphemeral();
}

// zero-arg overload for legacy callers
void CExplorerBHO::ApplySplitIfNeeded() {
	ApplySplitIfNeeded(false);
}

// main implementation; initialize consts where declared and use TabManager API that actually exists
void CExplorerBHO::ApplySplitIfNeeded(bool doSwap) {
	// ensure we know the hosting parent and the SHELLDLL_DefView
	Microsoft::WRL::ComPtr<IServiceProvider> sp;
	if ((!m_contentParent || !IsWindow(m_contentParent)) ||
		(!m_defView || !IsWindow(m_defView))) {
		if (m_site) m_site.As(&sp);
		m_contentParent = FindDefViewParentFromBrowser(sp.Get(), &m_defView);
	}
	if (!m_contentParent || !IsWindow(m_contentParent) || !m_defView) {
		RemoveSplitIfAny();
		return;
	}

	// Use the singleton TabManager and existing API
	auto& tm = shelltabs::TabManager::Get();

	// Selected group = "active group"
	const int g = tm.SelectedLocation().groupIndex;
	if (g < 0) {
		RemoveSplitIfAny();
		return;
	}

	const shelltabs::TabGroup* group = tm.GetGroup(g);
	if (!group || !group->splitView) {
		RemoveSplitIfAny();
		return;
	}

	// Create host if needed
	if (!m_splitHost || !IsWindow(m_splitHost)) {
		shelltabs::SplitHost::DestroyIfExistsOn(m_contentParent);
		if (auto host = shelltabs::SplitHost::CreateAndAttach(m_contentParent)) {
			m_splitHost = host->Hwnd();
			ShowWindow(m_defView, SW_HIDE);
		}
		else {
			return; // bail without hiding DefView
		}
	}

	// Resolve left/right PIDLs from the group's tabs (no non-existent ResolveTabAbsolutePIDL calls)
	const int leftIndex = group->splitPrimary;
	const int rightIndex = (group->splitSecondary >= 0) ? group->splitSecondary : group->splitPrimary;

	PCIDLIST_ABSOLUTE leftPIDL = nullptr;
	PCIDLIST_ABSOLUTE rightPIDL = nullptr;

	if (leftIndex >= 0 && leftIndex < static_cast<int>(group->tabs.size()))
		leftPIDL = group->tabs[static_cast<size_t>(leftIndex)].pidl.get();

	if (rightIndex >= 0 && rightIndex < static_cast<int>(group->tabs.size()))
		rightPIDL = group->tabs[static_cast<size_t>(rightIndex)].pidl.get();

	if (auto host = shelltabs::SplitHost::FromHwnd(m_splitHost)) {
		if (doSwap) host->Swap();
		else        host->SetFolders(leftPIDL, rightPIDL);
	}
}


static LRESULT CALLBACK ExplorerFrameSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
	UINT_PTR, DWORD_PTR refData) {
	CExplorerBHO* self = reinterpret_cast<CExplorerBHO*>(refData);
	if (!self) return DefSubclassProc(hwnd, msg, wp, lp);

	if (msg == WM_SHELLTABS_APPLY_SPLIT) {
		self->ApplySplitIfNeeded(wp == 1 /*swap*/);
		return 0;
	}
	return DefSubclassProc(hwnd, msg, wp, lp);
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
        if (m_dialogColorizer) {
                m_dialogColorizer->Detach();
                m_dialogColorizer.reset();
        }
        if (m_windowHook) {
                m_windowHook->Shutdown();
                m_windowHook.reset();
        }
        if (m_treeColorizer) {
                m_treeColorizer->Detach();
                m_treeColorizer.reset();
        }
        // UNSUBCLASS before we lose IWebBrowser2
        if (m_webBrowser) {
                if (HWND frame = GetExplorerFrameHwnd(m_webBrowser.Get())) {
                        RemoveWindowSubclass(frame, ExplorerFrameSubclassProc, 1);
                }
	}

	// Tear down any split host we created in this window
	RemoveSplitIfAny();
	m_contentParent = nullptr;
	m_defView = nullptr;
	m_splitHost = nullptr;

        // Event sink last
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

            if (m_dialogColorizer) {
                m_dialogColorizer->Detach();
                m_dialogColorizer.reset();
            }

            Disconnect();

            Microsoft::WRL::ComPtr<IFileDialog> fileDialog;
            if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&fileDialog))) && fileDialog) {
                auto colorizer = std::make_unique<CommonDialogColorizer>();
                if (colorizer && colorizer->Attach(fileDialog.Get())) {
                    m_dialogColorizer = std::move(colorizer);
                }
                return S_OK;
            }

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

            if (!m_windowHook) {
                m_windowHook = std::make_unique<ExplorerWindowHook>();
            }
            if (m_windowHook && !m_windowHook->Initialize(site, browser.Get())) {
                m_windowHook->Shutdown();
                m_windowHook.reset();
            }

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

            if (siteProvider) {
                if (!m_treeColorizer) {
                    m_treeColorizer = std::make_unique<NamespaceTreeColorizer>();
                }
                if (m_treeColorizer) {
                    m_treeColorizer->Detach();
                    m_treeColorizer->Attach(siteProvider);
                }
            }

            if (HWND frame = GetExplorerFrameHwnd(m_webBrowser.Get())) {
                SetWindowSubclass(frame, ExplorerFrameSubclassProc, 1, reinterpret_cast<DWORD_PTR>(this));
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

IFACEMETHODIMP CExplorerBHO::Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*,
                                    UINT*) {
    return GuardExplorerCall(
        L"CExplorerBHO::Invoke",
        [&]() -> HRESULT {
            switch (dispIdMember) {
                case DISPID_DOCUMENTCOMPLETE:
                case DISPID_NAVIGATECOMPLETE2:
                    ApplySplitIfNeeded();
                    break;
                case DISPID_ONVISIBLE:
                case DISPID_WINDOWSTATECHANGED:
                    if (!m_bandVisible) {
                        EnsureBandVisible();
                    }
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

}  // namespace shelltabs

