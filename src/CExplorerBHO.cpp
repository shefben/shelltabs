#include "CExplorerBHO.h"

#include <combaseapi.h>
#include <exdispid.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shlguid.h>
#include <CommCtrl.h>
#include <windowsx.h>
#include <uxtheme.h>
#include <OleIdl.h>
#include <gdiplus.h>

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdarg>
#include <cwchar>
#include <cwctype>
#include <mutex>
#include <string>
#include <unordered_map>
#include <vector>
#include <wrl/client.h>
#include <shobjidl_core.h>

#include "ComUtils.h"
#include "Guids.h"
#include "Logging.h"
#include "Module.h"
#include "OptionsStore.h"
#include "ShellTabsMessages.h"
#include "Utilities.h"

#ifndef TBSTATE_HOT
#define TBSTATE_HOT 0x80
#endif

#ifndef SFVIDM_CLIENT_OPENWINDOW
#define SFVIDM_CLIENT_OPENWINDOW 0x705B
#endif

namespace {

#ifndef ListView_GetItemW
BOOL ListView_GetItemW(HWND hwnd, LVITEMW* item) {
    return static_cast<BOOL>(SendMessageW(hwnd, LVM_GETITEMW, 0, reinterpret_cast<LPARAM>(item)));
}
#endif

#ifndef TreeView_GetItemW
BOOL TreeView_GetItemW(HWND hwnd, TVITEMEXW* item) {
    return static_cast<BOOL>(SendMessageW(hwnd, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(item)));
}
#endif

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

std::wstring NormalizeMenuText(const std::wstring& value) {
    if (value.empty()) {
        return {};
    }

    std::wstring normalized;
    normalized.reserve(value.size());
    for (wchar_t ch : value) {
        if (ch == L'&' || ch == L'.' || ch == 0x2026) {  // 0x2026 = ellipsis
            continue;
        }
        normalized.push_back(static_cast<wchar_t>(towlower(ch)));
    }

    const size_t first = normalized.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos) {
        return {};
    }
    const size_t last = normalized.find_last_not_of(L" \t\r\n");
    return normalized.substr(first, last - first + 1);
}

COLORREF SampleAverageColor(HDC dc, const RECT& rect) {
    if (!dc || rect.left >= rect.right || rect.top >= rect.bottom) {
        return GetSysColor(COLOR_WINDOW);
    }

    const LONG left = std::max(rect.left, static_cast<LONG>(0));
    const LONG top = std::max(rect.top, static_cast<LONG>(0));
    const LONG right = std::max(rect.right - 1, left);
    const LONG bottom = std::max(rect.bottom - 1, top);

    const std::array<POINT, 4> samplePoints = {{{left, top},
                                               {right, top},
                                               {left, bottom},
                                               {right, bottom}}};

    int totalRed = 0;
    int totalGreen = 0;
    int totalBlue = 0;
    int count = 0;

    for (const auto& point : samplePoints) {
        const COLORREF pixel = GetPixel(dc, point.x, point.y);
        if (pixel == CLR_INVALID) {
            continue;
        }
        totalRed += GetRValue(pixel);
        totalGreen += GetGValue(pixel);
        totalBlue += GetBValue(pixel);
        ++count;
    }

    if (count == 0) {
        return GetSysColor(COLOR_WINDOW);
    }

    return RGB(totalRed / count, totalGreen / count, totalBlue / count);
}

bool TryGetMenuItemText(HMENU menu, UINT position, std::wstring& text) {
    text.clear();
    if (!menu) {
        return false;
    }

    MENUITEMINFOW info{};
    info.cbSize = sizeof(info);
    info.fMask = MIIM_STRING;
    info.dwTypeData = nullptr;
    info.cch = 0;
    if (!GetMenuItemInfoW(menu, position, TRUE, &info)) {
        return false;
    }

    if (info.cch == 0) {
        return true;
    }

    std::wstring buffer;
    buffer.resize(info.cch);
    info.dwTypeData = buffer.data();
    info.cch = static_cast<UINT>(buffer.size());
    if (!GetMenuItemInfoW(menu, position, TRUE, &info)) {
        return false;
    }

    buffer.resize(info.cch);
    text = std::move(buffer);
    return true;
}

bool FindMenuItemById(HMENU menu, UINT commandId, UINT* position) {
    if (!menu) {
        return false;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0) {
        return false;
    }

    for (int i = 0; i < count; ++i) {
        const UINT id = GetMenuItemID(menu, i);
        if (id == commandId) {
            if (position) {
                *position = static_cast<UINT>(i);
            }
            return true;
        }
    }

    return false;
}

bool FindOpenInNewWindowMenuItem(HMENU menu, UINT* position, UINT* commandId) {
    if (!menu) {
        return false;
    }

    const UINT candidates[] = {SFVIDM_CLIENT_OPENWINDOW, 0x705A, 0x7059, 0x7020};
    for (UINT candidate : candidates) {
        UINT pos = 0;
        if (FindMenuItemById(menu, candidate, &pos)) {
            if (position) {
                *position = pos;
            }
            if (commandId) {
                *commandId = candidate;
            }
            return true;
        }
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0) {
        return false;
    }

    static const wchar_t* kTargets[] = {L"open in new window", L"open new window"};

    for (int i = 0; i < count; ++i) {
        const UINT id = GetMenuItemID(menu, i);
        if (id == UINT_MAX) {
            continue;
        }

        std::wstring text;
        if (!TryGetMenuItemText(menu, i, text)) {
            continue;
        }

        const std::wstring normalized = NormalizeMenuText(text);
        if (normalized.empty()) {
            continue;
        }

        for (const wchar_t* target : kTargets) {
            if (normalized == target) {
                if (position) {
                    *position = static_cast<UINT>(i);
                }
                if (commandId) {
                    *commandId = id;
                }
                return true;
            }
        }
    }

    return false;
}

constexpr wchar_t kOpenInNewTabLabel[] = L"Open in new tab";

struct BreadcrumbHookEntry {
    HHOOK hook = nullptr;
    std::vector<shelltabs::CExplorerBHO*> observers;
};

std::mutex g_breadcrumbHookMutex;
std::unordered_map<DWORD, BreadcrumbHookEntry> g_breadcrumbHooks;

}  // namespace
// --- CExplorerBHO private state (treat these as class members) ---

namespace shelltabs {

CExplorerBHO::CExplorerBHO() : m_refCount(1) {
    ModuleAddRef();
    m_bufferedPaintInitialized = SUCCEEDED(BufferedPaintInit());

    Gdiplus::GdiplusStartupInput gdiplusInput;
    if (Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusInput, nullptr) == Gdiplus::Ok) {
        m_gdiplusInitialized = true;
    } else {
        m_gdiplusToken = 0;
        LogMessage(LogLevel::Warning, L"Failed to initialize GDI+; breadcrumb gradient disabled");
    }
}

CExplorerBHO::~CExplorerBHO() {
    Disconnect();
    if (m_bufferedPaintInitialized) {
        BufferedPaintUnInit();
        m_bufferedPaintInitialized = false;
    }
    if (m_gdiplusInitialized) {
        Gdiplus::GdiplusShutdown(m_gdiplusToken);
        m_gdiplusInitialized = false;
        m_gdiplusToken = 0;
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
    RemoveBreadcrumbHook();
    RemoveBreadcrumbSubclass();
    RemoveExplorerViewSubclass();
    DisconnectEvents();
    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();
    m_bandVisible = false;
    m_shouldRetryEnsure = true;
    m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    m_loggedBreadcrumbToolbarMissing = false;
    m_lastBreadcrumbStage = BreadcrumbDiscoveryStage::None;
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
                LogMessage(LogLevel::Info, L"CExplorerBHO::SetSite detaching from site");
                Disconnect();
                return S_OK;
            }

            LogMessage(LogLevel::Info, L"CExplorerBHO::SetSite attaching to site=%p", site);
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
            UpdateExplorerViewSubclass();
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
                    UpdateExplorerViewSubclass();
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
        // fall through to normalize the window handle below
    } else if (m_webBrowser) {
        SHANDLE_PTR raw = 0;
        if (SUCCEEDED(m_webBrowser->get_HWND(&raw)) && raw) {
            hwnd = reinterpret_cast<HWND>(raw);
        }
    }

    if (!hwnd) {
        return nullptr;
    }

    HWND ancestor = GetAncestor(hwnd, GA_ROOTOWNER);
    if (ancestor) {
        hwnd = ancestor;
    }

    ancestor = GetAncestor(hwnd, GA_ROOT);
    if (ancestor) {
        hwnd = ancestor;
    }

    // Walk up the parent chain in case GetAncestor returned a child window.
    HWND current = hwnd;
    HWND parent = nullptr;
    int safety = 0;
    while (current && safety++ < 32) {
        parent = GetParent(current);
        if (!parent) {
            break;
        }
        current = parent;
    }

    return current ? current : hwnd;
}

void CExplorerBHO::LogBreadcrumbStage(BreadcrumbDiscoveryStage stage, const wchar_t* format, ...) const {
    if (!format) {
        return;
    }
    if (m_lastBreadcrumbStage == stage) {
        return;
    }

    m_lastBreadcrumbStage = stage;

    va_list args;
    va_start(args, format);
    LogMessageV(LogLevel::Info, format, args);
    va_end(args);
}

HWND CExplorerBHO::FindBreadcrumbToolbar() const {
    auto queryBreadcrumbToolbar = [&](const Microsoft::WRL::ComPtr<IServiceProvider>& provider,
                                     const wchar_t* source) -> HWND {
        if (!provider) {
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IUnknown> breadcrumbService;
        HRESULT hr = provider->QueryService(CLSID_CBreadcrumbBar, IID_PPV_ARGS(&breadcrumbService));
        if (FAILED(hr) || !breadcrumbService) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceUnavailable,
                               L"Breadcrumb QueryService(%s) failed: 0x%08X", source ? source : L"?", hr);
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
        hr = breadcrumbService.As(&oleWindow);
        if (FAILED(hr) || !oleWindow) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceWindowMissing,
                               L"Breadcrumb service missing IOleWindow (%s): 0x%08X", source ? source : L"?", hr);
            return nullptr;
        }

        HWND bandWindow = nullptr;
        hr = oleWindow->GetWindow(&bandWindow);
        if (FAILED(hr) || !bandWindow) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceWindowMissing,
                               L"Breadcrumb service window unavailable (%s): 0x%08X", source ? source : L"?", hr);
            return nullptr;
        }

        HWND toolbar = FindWindowExW(bandWindow, nullptr, TOOLBARCLASSNAME, nullptr);
        if (!toolbar) {
            toolbar = FindDescendantWindow(bandWindow, TOOLBARCLASSNAME);
        }
        if (toolbar) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                               L"Breadcrumb toolbar located via %s service (hwnd=%p)", source ? source : L"?",
                               toolbar);
        } else {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceToolbarMissing,
                               L"Breadcrumb service band (%s hwnd=%p) missing toolbar child", source ? source : L"?",
                               bandWindow);
        }
        return toolbar;
    };

    if (m_shellBrowser) {
        Microsoft::WRL::ComPtr<IServiceProvider> provider;
        if (SUCCEEDED(m_shellBrowser.As(&provider))) {
            if (HWND fromService = queryBreadcrumbToolbar(provider, L"IShellBrowser")) {
                return fromService;
            }
        }
    }

    if (m_webBrowser) {
        Microsoft::WRL::ComPtr<IServiceProvider> provider;
        if (SUCCEEDED(m_webBrowser.As(&provider))) {
            if (HWND fromService = queryBreadcrumbToolbar(provider, L"IWebBrowser2")) {
                return fromService;
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::FrameMissing,
                           L"Top-level Explorer window unavailable during breadcrumb search");
        return nullptr;
    }

    HWND travelBand = FindDescendantWindow(frame, L"TravelBand");
    HWND rebar = travelBand ? GetParent(travelBand) : nullptr;
    if (!rebar) {
        rebar = FindDescendantWindow(frame, L"ReBarWindow32");
    }
    if (!rebar) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::RebarMissing,
                           L"Failed to locate Explorer rebar while searching for breadcrumbs");
        return FindBreadcrumbToolbarInWindow(frame);
    }

    HWND breadcrumbParent = FindWindowExW(rebar, nullptr, L"Breadcrumb Parent", nullptr);
    if (!breadcrumbParent) {
        breadcrumbParent = FindDescendantWindow(rebar, L"Breadcrumb Parent");
    }
    if (!breadcrumbParent) {
        breadcrumbParent = FindDescendantWindow(frame, L"Breadcrumb Parent");
    }
    if (!breadcrumbParent) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::ParentMissing,
                           L"Failed to find 'Breadcrumb Parent' window during breadcrumb search");
        return FindBreadcrumbToolbarInWindow(frame);
    }

    HWND toolbar = FindWindowExW(breadcrumbParent, nullptr, TOOLBARCLASSNAME, nullptr);
    if (!toolbar) {
        toolbar = FindDescendantWindow(breadcrumbParent, TOOLBARCLASSNAME);
    }
    if (!toolbar) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::ToolbarMissing,
                           L"'Breadcrumb Parent' hwnd=%p missing ToolbarWindow32 child", breadcrumbParent);
        return FindBreadcrumbToolbarInWindow(breadcrumbParent);
    }

    LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                       L"Breadcrumb toolbar located via window enumeration (hwnd=%p)", toolbar);
    return toolbar;
}

HWND CExplorerBHO::FindBreadcrumbToolbarInWindow(HWND root) const {
    if (!root) {
        return nullptr;
    }

    struct EnumData {
        const CExplorerBHO* self = nullptr;
        HWND toolbar = nullptr;
    } data{this, nullptr};

    EnumChildWindows(
        root,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* data = reinterpret_cast<EnumData*>(param);
            if (!data || data->toolbar) {
                return FALSE;
            }
            if (!MatchesClass(hwnd, TOOLBARCLASSNAME)) {
                return TRUE;
            }
            if (!data->self->IsBreadcrumbToolbarCandidate(hwnd)) {
                return TRUE;
            }
            data->toolbar = hwnd;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    if (data.toolbar) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                           L"Breadcrumb toolbar located via deep enumeration (hwnd=%p)", data.toolbar);
    }

    return data.toolbar;
}

bool CExplorerBHO::IsBreadcrumbToolbarAncestor(HWND hwnd) const {
    HWND current = hwnd;
    bool sawRebar = false;
    int depth = 0;
    while (current && depth++ < 16) {
        if (MatchesClass(current, L"Breadcrumb Parent") || MatchesClass(current, L"Address Band Root") ||
            MatchesClass(current, L"AddressBandRoot") || MatchesClass(current, L"CabinetAddressBand") ||
            MatchesClass(current, L"NavigationBand")) {
            return true;
        }
        if (MatchesClass(current, L"ReBarWindow32")) {
            sawRebar = true;
        }
        if (MatchesClass(current, L"CabinetWClass")) {
            break;
        }
        current = GetParent(current);
    }
    return sawRebar;
}

bool CExplorerBHO::IsBreadcrumbToolbarCandidate(HWND hwnd) const {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }
    if (!MatchesClass(hwnd, TOOLBARCLASSNAME)) {
        return false;
    }

    if (!IsBreadcrumbToolbarAncestor(hwnd)) {
        return false;
    }

    LRESULT buttonCount = SendMessage(hwnd, TB_BUTTONCOUNT, 0, 0);
    if (buttonCount <= 0) {
        return false;
    }

    const int maxToCheck = static_cast<int>(std::min<LRESULT>(buttonCount, 5));
    std::array<wchar_t, 260> buffer{};
    TBBUTTON button{};
    for (int i = 0; i < maxToCheck; ++i) {
        if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&button))) {
            continue;
        }
        if ((button.fsStyle & TBSTYLE_SEP) != 0 || (button.fsState & TBSTATE_HIDDEN) != 0) {
            continue;
        }
        buffer.fill(L'\0');
        LRESULT copied = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand, reinterpret_cast<LPARAM>(buffer.data()));
        if (copied > 0 && buffer[0] != L'\0') {
            return true;
        }
    }

    return false;
}

bool CExplorerBHO::IsWindowOwnedByThisExplorer(HWND hwnd) const {
    HWND frame = GetTopLevelExplorerWindow();
    if (!frame || !IsWindow(frame)) {
        return false;
    }

    HWND current = hwnd;
    int depth = 0;
    while (current && depth++ < 32) {
        if (current == frame) {
            return true;
        }
        current = GetParent(current);
    }

    HWND root = GetAncestor(hwnd, GA_ROOT);
    return root == frame;
}

void CExplorerBHO::UpdateExplorerViewSubclass() {
    RemoveExplorerViewSubclass();

    if (!m_shellBrowser) {
        return;
    }

    Microsoft::WRL::ComPtr<IShellView> shellView;
    HRESULT hr = m_shellBrowser->QueryActiveShellView(&shellView);
    if (FAILED(hr) || !shellView) {
        return;
    }

    HWND viewWindow = nullptr;
    hr = shellView->GetWindow(&viewWindow);
    if (FAILED(hr) || !viewWindow) {
        return;
    }

    HWND listView = FindDescendantWindow(viewWindow, L"SysListView32");
    if (!listView) {
        LogMessage(LogLevel::Warning,
                   L"Explorer view subclass setup aborted: list view not found for view window %p", viewWindow);
        return;
    }

    HWND treeView = FindDescendantWindow(viewWindow, L"SysTreeView32");
    if (!treeView) {
        LogMessage(LogLevel::Info,
                   L"Explorer view subclass setup continuing without tree view (view=%p list=%p)", viewWindow,
                   listView);
    }

    if (!InstallExplorerViewSubclass(viewWindow, listView, treeView)) {
        LogMessage(LogLevel::Warning,
                   L"Explorer view subclass installation failed (view=%p list=%p tree=%p)", viewWindow, listView,
                   treeView);
        return;
    }

    m_shellView = shellView;
    m_shellViewWindow = viewWindow;
}

bool CExplorerBHO::InstallExplorerViewSubclass(HWND viewWindow, HWND listView, HWND treeView) {
    bool installed = false;

    if (viewWindow && IsWindow(viewWindow)) {
        if (SetWindowSubclass(viewWindow, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
            m_shellViewWindowSubclassInstalled = true;
            installed = true;
            LogMessage(LogLevel::Info, L"Installed shell view window subclass (view=%p)", viewWindow);
        } else {
            LogLastError(L"SetWindowSubclass(shell view window)", GetLastError());
            m_shellViewWindowSubclassInstalled = false;
        }
    } else {
        m_shellViewWindowSubclassInstalled = false;
    }

    HWND frameWindow = GetTopLevelExplorerWindow();
    if (frameWindow && frameWindow != viewWindow && frameWindow != listView && frameWindow != treeView &&
        IsWindow(frameWindow)) {
        if (SetWindowSubclass(frameWindow, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
            m_frameWindow = frameWindow;
            m_frameSubclassInstalled = true;
            installed = true;
            LogMessage(LogLevel::Info, L"Installed explorer frame subclass (frame=%p)", frameWindow);
        } else {
            LogLastError(L"SetWindowSubclass(explorer frame)", GetLastError());
            m_frameSubclassInstalled = false;
            m_frameWindow = nullptr;
        }
    } else {
        m_frameSubclassInstalled = false;
        m_frameWindow = nullptr;
    }

    if (listView && IsWindow(listView)) {
        if (SetWindowSubclass(listView, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
            m_listView = listView;
            m_listViewSubclassInstalled = true;
            installed = true;
            LogMessage(LogLevel::Info, L"Installed explorer list view subclass (list=%p)", listView);
        } else {
            LogLastError(L"SetWindowSubclass(list view)", GetLastError());
        }
    }

    if (treeView && treeView != listView && IsWindow(treeView)) {
        if (SetWindowSubclass(treeView, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
            m_treeView = treeView;
            m_treeViewSubclassInstalled = true;
            LogMessage(LogLevel::Info, L"Installed explorer tree view subclass (tree=%p)", treeView);
        } else {
            LogLastError(L"SetWindowSubclass(tree view)", GetLastError());
        }
    }

    if (installed) {
        ClearPendingOpenInNewTabState();
        LogMessage(LogLevel::Info, L"Explorer view subclass ready (view=%p list=%p tree=%p)", viewWindow,
                   listView, treeView);
    } else {
        LogMessage(LogLevel::Warning,
                   L"Explorer view subclass installation skipped: no valid targets (view=%p list=%p tree=%p)",
                   viewWindow, listView, treeView);
    }

    return installed;
}

void CExplorerBHO::RemoveExplorerViewSubclass() {
    if (m_shellViewWindow && m_shellViewWindowSubclassInstalled && IsWindow(m_shellViewWindow)) {
        RemoveWindowSubclass(m_shellViewWindow, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }
    if (m_frameWindow && m_frameSubclassInstalled && IsWindow(m_frameWindow)) {
        RemoveWindowSubclass(m_frameWindow, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }
    if (m_listView && m_listViewSubclassInstalled && IsWindow(m_listView)) {
        RemoveWindowSubclass(m_listView, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }
    if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
        RemoveWindowSubclass(m_treeView, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }

    m_shellViewWindowSubclassInstalled = false;
    m_frameWindow = nullptr;
    m_frameSubclassInstalled = false;
    m_listView = nullptr;
    m_treeView = nullptr;
    m_listViewSubclassInstalled = false;
    m_treeViewSubclassInstalled = false;
    m_shellViewWindow = nullptr;
    m_shellView.Reset();
    ClearPendingOpenInNewTabState();
}

bool CExplorerBHO::HandleExplorerViewMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                             LRESULT* result) {
    if (!result) {
        return false;
    }

    const UINT optionsChangedMessage = GetOptionsChangedMessage();
    if (optionsChangedMessage != 0 && msg == optionsChangedMessage) {
        UpdateBreadcrumbSubclass();
        if (m_breadcrumbToolbar && m_breadcrumbSubclassInstalled && IsWindow(m_breadcrumbToolbar)) {
            InvalidateRect(m_breadcrumbToolbar, nullptr, TRUE);
        }
        *result = 0;
        return true;
    }

    switch (msg) {
        case WM_INITMENUPOPUP: {
            if (HIWORD(lParam) == 0) {
                HandleExplorerContextMenuInit(hwnd, reinterpret_cast<HMENU>(wParam));
            }
            break;
        }
        case WM_CONTEXTMENU: {
            POINT screenPoint{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            PrepareContextMenuSelection(reinterpret_cast<HWND>(wParam), screenPoint);
            break;
        }
        case WM_COMMAND: {
            const UINT commandId = LOWORD(wParam);
            if (commandId == kOpenInNewTabCommandId) {
                HandleExplorerCommand(commandId);
                *result = 0;
                return true;
            }
            break;
        }
        case WM_MENUCOMMAND: {
            HMENU menu = reinterpret_cast<HMENU>(lParam);
            const UINT position = static_cast<UINT>(wParam);
            if (menu && GetMenuItemID(menu, position) == kOpenInNewTabCommandId) {
                HandleExplorerCommand(kOpenInNewTabCommandId);
                *result = 0;
                return true;
            }
            break;
        }
        case WM_UNINITMENUPOPUP: {
            HandleExplorerMenuDismiss(reinterpret_cast<HMENU>(wParam));
            break;
        }
        case WM_CANCELMODE: {
            HandleExplorerMenuDismiss(m_trackedContextMenu);
            break;
        }
        default:
            break;
    }

    return false;
}

void CExplorerBHO::HandleExplorerContextMenuInit(HWND source, HMENU menu) {
    LogMessage(LogLevel::Info, L"Explorer context menu init (menu=%p source=%p inserted=%d tracking=%p)", menu, source,
               m_contextMenuInserted ? 1 : 0, m_trackedContextMenu);

    if (!menu) {
        LogMessage(LogLevel::Warning, L"Context menu init aborted: menu handle missing");
        return;
    }

    if (m_contextMenuInserted) {
        LogMessage(LogLevel::Info, L"Context menu init skipped: already inserted for this cycle");
        return;
    }

    if (m_trackedContextMenu && menu != m_trackedContextMenu) {
        LogMessage(LogLevel::Info, L"Context menu init skipped: still tracking menu %p", m_trackedContextMenu);
        return;
    }

    ClearPendingOpenInNewTabState();

    std::vector<std::wstring> paths;
    if (!CollectSelectedFolderPaths(paths) || paths.empty()) {
        LogMessage(LogLevel::Info, L"Context menu init aborted: no eligible folder selection detected");
        return;
    }

    UINT position = 0;
    const bool anchorFound = FindOpenInNewWindowMenuItem(menu, &position, nullptr);

    if (GetMenuState(menu, kOpenInNewTabCommandId, MF_BYCOMMAND) != static_cast<UINT>(-1)) {
        LogMessage(LogLevel::Info, L"Context menu already contains Open In New Tab entry");
        return;
    }

    MENUITEMINFOW item{};
    item.cbSize = sizeof(item);
    item.fMask = MIIM_ID | MIIM_STRING | MIIM_FTYPE | MIIM_STATE;
    item.fType = MFT_STRING;
    item.fState = MFS_ENABLED;
    item.wID = kOpenInNewTabCommandId;
    item.dwTypeData = const_cast<wchar_t*>(kOpenInNewTabLabel);

    UINT insertPosition = 0;
    if (anchorFound) {
        insertPosition = position + 1;
    } else {
        LogMessage(LogLevel::Info, L"Context menu init continuing without explicit anchor");
        const int itemCount = GetMenuItemCount(menu);
        insertPosition = itemCount > 0 ? static_cast<UINT>(itemCount) : 0;
    }

    if (!InsertMenuItemW(menu, insertPosition, TRUE, &item)) {
        LogLastError(L"InsertMenuItem(Open In New Tab)", GetLastError());
        return;
    }

    m_pendingOpenInNewTabPaths = std::move(paths);
    m_contextMenuInserted = true;
    m_trackedContextMenu = menu;
    LogMessage(LogLevel::Info, L"Open In New Tab inserted at position %u for %zu paths", insertPosition + 1,
               m_pendingOpenInNewTabPaths.size());
}

void CExplorerBHO::PrepareContextMenuSelection(HWND sourceWindow, POINT screenPoint) {
    HWND target = sourceWindow;
    if (!target || !IsWindow(target)) {
        target = GetFocus();
    }

    if (!target || (!IsWindow(target))) {
        return;
    }

    if (target == m_listView) {
        if (screenPoint.x == -1 && screenPoint.y == -1) {
            return;
        }

        POINT clientPoint = screenPoint;
        if (!ScreenToClient(target, &clientPoint)) {
            return;
        }

        LVHITTESTINFO hit{};
        hit.pt = clientPoint;
        const int index = ListView_SubItemHitTest(m_listView, &hit);
        if (index < 0 || (hit.flags & LVHT_ONITEM) == 0) {
            return;
        }

        const UINT state = ListView_GetItemState(m_listView, index, LVIS_SELECTED);
        if ((state & LVIS_SELECTED) != 0) {
            return;
        }

        ListView_SetItemState(m_listView, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_SetItemState(m_listView, index, LVIS_SELECTED | LVIS_FOCUSED,
                              LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(m_listView, index, FALSE);
        LogMessage(LogLevel::Info, L"Context menu selection synchronized to list view item %d", index);
        return;
    }

    if (target == m_treeView) {
        if (screenPoint.x == -1 && screenPoint.y == -1) {
            return;
        }

        POINT clientPoint = screenPoint;
        if (!ScreenToClient(target, &clientPoint)) {
            return;
        }

        TVHITTESTINFO hit{};
        hit.pt = clientPoint;
        HTREEITEM item = TreeView_HitTest(m_treeView, &hit);
        if (!item || (hit.flags & (TVHT_ONITEM | TVHT_ONITEMBUTTON | TVHT_ONITEMINDENT)) == 0) {
            return;
        }

        HTREEITEM current = TreeView_GetSelection(m_treeView);
        if (current == item) {
            return;
        }

        TreeView_SelectItem(m_treeView, item);
        LogMessage(LogLevel::Info, L"Context menu selection synchronized to tree view item %p", item);
    }
}

void CExplorerBHO::HandleExplorerCommand(UINT commandId) {
    if (commandId != kOpenInNewTabCommandId) {
        return;
    }

    std::vector<std::wstring> paths = m_pendingOpenInNewTabPaths;
    if (paths.empty()) {
        if (!CollectSelectedFolderPaths(paths)) {
            LogMessage(LogLevel::Warning, L"Open In New Tab command aborted: unable to resolve folder selection");
            ClearPendingOpenInNewTabState();
            return;
        }
    }

    LogMessage(LogLevel::Info, L"Open In New Tab command executing for %zu paths", paths.size());
    DispatchOpenInNewTab(paths);
    ClearPendingOpenInNewTabState();
}

void CExplorerBHO::HandleExplorerMenuDismiss(HMENU menu) {
    if (!m_trackedContextMenu) {
        return;
    }

    if (!menu || menu == m_trackedContextMenu) {
        LogMessage(LogLevel::Info, L"Explorer context menu dismissed (menu=%p)", menu);
        ClearPendingOpenInNewTabState();
    }
}

bool CExplorerBHO::CollectSelectedFolderPaths(std::vector<std::wstring>& paths) const {
    paths.clear();

    bool success = false;

    if (CollectPathsFromShellViewSelection(paths)) {
        success = true;
    }

    if (!success) {
        if (CollectPathsFromFolderViewSelection(paths)) {
            success = true;
        }
    }

    if (!success) {
        if (CollectPathsFromListView(paths)) {
            success = true;
        }
    }

    if (!success) {
        if (CollectPathsFromTreeView(paths)) {
            success = true;
        }
    }

    if (!success) {
        LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths found no eligible folders");
    } else {
        LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths captured %zu path(s)", paths.size());
    }

    return success && !paths.empty();
}

bool CExplorerBHO::CollectPathsFromShellViewSelection(std::vector<std::wstring>& paths) const {
    if (!m_shellView) {
        LogMessage(LogLevel::Warning, L"CollectPathsFromShellViewSelection failed: shell view unavailable");
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItemArray> items;
    HRESULT hr = m_shellView->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&items));
    if (FAILED(hr) || !items) {
        LogMessage(LogLevel::Info, L"CollectPathsFromShellViewSelection skipped: selection unavailable (hr=0x%08lX)", hr);
        return false;
    }

    return CollectPathsFromItemArray(items.Get(), paths);
}

bool CExplorerBHO::CollectPathsFromFolderViewSelection(std::vector<std::wstring>& paths) const {
    if (!m_shellView) {
        return false;
    }

    Microsoft::WRL::ComPtr<IFolderView2> folderView;
    HRESULT hr = m_shellView.As(&folderView);
    if (FAILED(hr) || !folderView) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItemArray> items;
    hr = folderView->GetSelection(TRUE, &items);
    if (FAILED(hr) || !items) {
        LogMessage(LogLevel::Info,
                   L"CollectPathsFromFolderViewSelection skipped: unable to resolve folder view selection (hr=0x%08lX)",
                   hr);
        return false;
    }

    return CollectPathsFromItemArray(items.Get(), paths);
}

bool CExplorerBHO::CollectPathsFromItemArray(IShellItemArray* items, std::vector<std::wstring>& paths) const {
    if (!items) {
        return false;
    }

    DWORD count = 0;
    HRESULT hr = items->GetCount(&count);
    if (FAILED(hr) || count == 0) {
        LogMessage(LogLevel::Info, L"CollectPathsFromItemArray skipped: count=%lu hr=0x%08lX", count, hr);
        return false;
    }

    if (count > kMaxTrackedSelection) {
        LogMessage(LogLevel::Info, L"CollectPathsFromItemArray limiting selection from %lu to %u entries", count,
                   kMaxTrackedSelection);
    }

    bool appended = false;

    for (DWORD index = 0; index < count && paths.size() < kMaxTrackedSelection; ++index) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(index, &item)) || !item) {
            LogMessage(LogLevel::Warning, L"CollectPathsFromItemArray failed: unable to access item %lu", index);
            continue;
        }

        PIDLIST_ABSOLUTE pidl = nullptr;
        hr = SHGetIDListFromObject(item.Get(), &pidl);
        if (FAILED(hr) || !pidl) {
            PWSTR path = nullptr;
            hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path);
            if (FAILED(hr) || !path || path[0] == L'\0') {
                if (path) {
                    CoTaskMemFree(path);
                }
                LogMessage(LogLevel::Info,
                           L"CollectPathsFromItemArray skipping item %lu: missing PIDL and filesystem path (hr=0x%08lX)",
                           index, hr);
                continue;
            }

            std::wstring value(path);
            CoTaskMemFree(path);
            if (value.empty()) {
                continue;
            }
            if (std::find(paths.begin(), paths.end(), value) == paths.end() &&
                paths.size() < kMaxTrackedSelection) {
                paths.push_back(std::move(value));
                appended = true;
            }
            continue;
        }

        if (AppendPathFromPidl(pidl, paths)) {
            appended = true;
        }
        CoTaskMemFree(pidl);
    }

    return appended;
}

bool CExplorerBHO::CollectPathsFromListView(std::vector<std::wstring>& paths) const {
    if (!m_listView || !IsWindow(m_listView)) {
        return false;
    }

    int index = -1;
    bool appended = false;
    while ((index = ListView_GetNextItem(m_listView, index, LVNI_SELECTED)) != -1) {
        if (paths.size() >= kMaxTrackedSelection) {
            LogMessage(LogLevel::Info, L"CollectPathsFromListView truncated selection at %zu entries", paths.size());
            break;
        }

        LVITEMW item{};
        item.mask = LVIF_PARAM;
        item.iItem = index;
        if (!ListView_GetItemW(m_listView, &item)) {
            LogLastError(L"ListView_GetItem(selection)", GetLastError());
            continue;
        }

        if (AppendPathFromPidl(reinterpret_cast<PCIDLIST_ABSOLUTE>(item.lParam), paths)) {
            appended = true;
        }
    }

    if (!appended) {
        LogMessage(LogLevel::Info, L"CollectPathsFromListView found no folder selections");
    }

    return appended;
}

bool CExplorerBHO::CollectPathsFromTreeView(std::vector<std::wstring>& paths) const {
    if (!m_treeView || !IsWindow(m_treeView)) {
        return false;
    }

    if (paths.size() >= kMaxTrackedSelection) {
        return false;
    }

    HTREEITEM selection = TreeView_GetSelection(m_treeView);
    if (!selection) {
        LogMessage(LogLevel::Info, L"CollectPathsFromTreeView skipped: no selection");
        return false;
    }

    TVITEMEXW item{};
    item.mask = TVIF_PARAM;
    item.hItem = selection;
    if (!TreeView_GetItemW(m_treeView, &item)) {
        LogLastError(L"TreeView_GetItem(selection)", GetLastError());
        return false;
    }

    if (!AppendPathFromPidl(reinterpret_cast<PCIDLIST_ABSOLUTE>(item.lParam), paths)) {
        LogMessage(LogLevel::Info, L"CollectPathsFromTreeView skipped: selection not a filesystem folder");
        return false;
    }

    return !paths.empty();
}

bool CExplorerBHO::AppendPathFromPidl(PCIDLIST_ABSOLUTE pidl, std::vector<std::wstring>& paths) const {
    if (!pidl) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellFolder> parentFolder;
    PCUITEMID_CHILD child = nullptr;
    HRESULT hr = SHBindToParent(pidl, IID_PPV_ARGS(&parentFolder), &child);
    if (FAILED(hr) || !parentFolder || !child) {
        LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: unable to bind to parent (hr=0x%08lX)", hr);
        return false;
    }

    SFGAOF attributes = SFGAO_FOLDER | SFGAO_FILESYSTEM;
    hr = parentFolder->GetAttributesOf(1, &child, &attributes);
    if (FAILED(hr) || (attributes & SFGAO_FOLDER) == 0 || (attributes & SFGAO_FILESYSTEM) == 0) {
        LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: attributes=0x%08lX (hr=0x%08lX)", attributes, hr);
        return false;
    }

    PWSTR path = nullptr;
    hr = SHGetNameFromIDList(pidl, SIGDN_FILESYSPATH, &path);
    if (FAILED(hr) || !path || path[0] == L'\0') {
        if (path) {
            CoTaskMemFree(path);
        }
        LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: unable to resolve filesystem path (hr=0x%08lX)", hr);
        return false;
    }

    std::wstring value(path);
    CoTaskMemFree(path);

    if (value.empty()) {
        return false;
    }

    if (std::find(paths.begin(), paths.end(), value) != paths.end()) {
        return true;
    }

    paths.push_back(std::move(value));
    return true;
}

void CExplorerBHO::DispatchOpenInNewTab(const std::vector<std::wstring>& paths) const {
    if (paths.empty()) {
        LogMessage(LogLevel::Info, L"DispatchOpenInNewTab skipped: no paths provided");
        return;
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        LogMessage(LogLevel::Warning, L"DispatchOpenInNewTab failed: explorer frame not found");
        return;
    }

    HWND bandWindow = FindDescendantWindow(frame, L"ShellTabsBandWindow");
    if (!bandWindow || !IsWindow(bandWindow)) {
        LogMessage(LogLevel::Warning, L"DispatchOpenInNewTab failed: ShellTabs band window missing (frame=%p)", frame);
        return;
    }

    for (const std::wstring& path : paths) {
        if (path.empty()) {
            LogMessage(LogLevel::Warning, L"DispatchOpenInNewTab skipped empty path entry");
            continue;
        }

        OpenFolderMessagePayload payload{path.c_str(), path.size()};
        SendMessageW(bandWindow, WM_SHELLTABS_OPEN_FOLDER, reinterpret_cast<WPARAM>(&payload), 0);
        LogMessage(LogLevel::Info, L"Dispatched Open In New Tab request for %ls", path.c_str());
    }
}

void CExplorerBHO::ClearPendingOpenInNewTabState() {
    m_pendingOpenInNewTabPaths.clear();
    m_trackedContextMenu = nullptr;
    m_contextMenuInserted = false;
    LogMessage(LogLevel::Info, L"Cleared Open In New Tab pending state");
}

bool CExplorerBHO::InstallBreadcrumbSubclass(HWND toolbar) {
    if (!toolbar || !IsWindow(toolbar)) {
        return false;
    }

    if (toolbar == m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        return true;
    }

    RemoveBreadcrumbSubclass();

    if (SetWindowSubclass(toolbar, BreadcrumbSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_breadcrumbToolbar = toolbar;
        m_breadcrumbSubclassInstalled = true;
        m_loggedBreadcrumbToolbarMissing = false;
        LogMessage(LogLevel::Info, L"Installed breadcrumb gradient subclass on hwnd=%p", toolbar);
        InvalidateRect(toolbar, nullptr, TRUE);
        return true;
    }

    LogLastError(L"SetWindowSubclass(breadcrumb toolbar)", GetLastError());
    return false;
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
    if (m_breadcrumbLogState == BreadcrumbLogState::Searching) {
        m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    }
    m_loggedBreadcrumbToolbarMissing = false;
}

void CExplorerBHO::EnsureBreadcrumbHook() {
    if (m_breadcrumbHookRegistered) {
        return;
    }

    const DWORD threadId = GetCurrentThreadId();
    std::lock_guard<std::mutex> lock(g_breadcrumbHookMutex);
    auto& entry = g_breadcrumbHooks[threadId];
    if (std::find(entry.observers.begin(), entry.observers.end(), this) == entry.observers.end()) {
        entry.observers.push_back(this);
    }

    if (!entry.hook) {
        HHOOK hook = SetWindowsHookExW(WH_CBT, BreadcrumbCbtProc, nullptr, threadId);
        if (!hook) {
            LogLastError(L"SetWindowsHookEx(WH_CBT)", GetLastError());
            entry.observers.erase(std::remove(entry.observers.begin(), entry.observers.end(), this), entry.observers.end());
            if (entry.observers.empty()) {
                g_breadcrumbHooks.erase(threadId);
            }
            return;
        }
        entry.hook = hook;
        LogMessage(LogLevel::Info, L"Breadcrumb CBT hook installed for thread %lu", threadId);
    }

    m_breadcrumbHookRegistered = true;
}

void CExplorerBHO::RemoveBreadcrumbHook() {
    if (!m_breadcrumbHookRegistered) {
        return;
    }

    const DWORD threadId = GetCurrentThreadId();
    std::lock_guard<std::mutex> lock(g_breadcrumbHookMutex);
    auto it = g_breadcrumbHooks.find(threadId);
    if (it == g_breadcrumbHooks.end()) {
        m_breadcrumbHookRegistered = false;
        return;
    }

    auto& observers = it->second.observers;
    observers.erase(std::remove(observers.begin(), observers.end(), this), observers.end());
    if (observers.empty()) {
        if (it->second.hook) {
            UnhookWindowsHookEx(it->second.hook);
        }
        g_breadcrumbHooks.erase(it);
        LogMessage(LogLevel::Info, L"Breadcrumb CBT hook removed for thread %lu", threadId);
    }

    m_breadcrumbHookRegistered = false;
}

void CExplorerBHO::UpdateBreadcrumbSubclass() {
    auto& store = OptionsStore::Instance();
    store.Load();
    const ShellTabsOptions options = store.Get();
    m_breadcrumbGradientEnabled = options.enableBreadcrumbGradient;
    m_breadcrumbFontGradientEnabled = options.enableBreadcrumbFontGradient;
    m_breadcrumbGradientTransparency = std::clamp(options.breadcrumbGradientTransparency, 0, 100);
    m_breadcrumbFontTransparency = std::clamp(options.breadcrumbFontTransparency, 0, 100);
    m_useCustomBreadcrumbGradientColors = options.useCustomBreadcrumbGradientColors;
    m_breadcrumbGradientStartColor = options.breadcrumbGradientStartColor;
    m_breadcrumbGradientEndColor = options.breadcrumbGradientEndColor;
    m_useCustomBreadcrumbFontColors = options.useCustomBreadcrumbFontColors;
    m_breadcrumbFontGradientStartColor = options.breadcrumbFontGradientStartColor;
    m_breadcrumbFontGradientEndColor = options.breadcrumbFontGradientEndColor;

    const bool gradientsEnabled = (m_breadcrumbGradientEnabled || m_breadcrumbFontGradientEnabled);
    if (!gradientsEnabled || !m_gdiplusInitialized) {
        if (m_breadcrumbLogState != BreadcrumbLogState::Disabled) {
            LogMessage(LogLevel::Info,
                       L"Breadcrumb gradients inactive (background=%d text=%d gdiplus=%d); ensuring subclass removed",
                       m_breadcrumbGradientEnabled ? 1 : 0, m_breadcrumbFontGradientEnabled ? 1 : 0,
                       m_gdiplusInitialized ? 1 : 0);
            m_breadcrumbLogState = BreadcrumbLogState::Disabled;
        }
        if (m_breadcrumbSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Breadcrumb gradients disabled; removing subclass");
        }
        RemoveBreadcrumbHook();
        RemoveBreadcrumbSubclass();
        m_loggedBreadcrumbToolbarMissing = false;
        return;
    }

    EnsureBreadcrumbHook();

    if (m_breadcrumbLogState != BreadcrumbLogState::Searching) {
        LogMessage(LogLevel::Info,
                   L"Breadcrumb gradients enabled; locating toolbar (installed=%d background=%d text=%d)",
                   m_breadcrumbSubclassInstalled ? 1 : 0, m_breadcrumbGradientEnabled ? 1 : 0,
                   m_breadcrumbFontGradientEnabled ? 1 : 0);
        m_lastBreadcrumbStage = BreadcrumbDiscoveryStage::None;
        m_breadcrumbLogState = BreadcrumbLogState::Searching;
    }

    HWND toolbar = FindBreadcrumbToolbar();
    if (!toolbar) {
        if (!m_loggedBreadcrumbToolbarMissing) {
            LogMessage(LogLevel::Info, L"Breadcrumb toolbar not yet available; will retry");
            m_loggedBreadcrumbToolbarMissing = true;
        }
        if (m_breadcrumbSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Breadcrumb toolbar not found; removing subclass");
        }
        RemoveBreadcrumbSubclass();
        return;
    }

    if (m_loggedBreadcrumbToolbarMissing) {
        LogMessage(LogLevel::Info, L"Breadcrumb toolbar discovered after retry");
    }
    m_loggedBreadcrumbToolbarMissing = false;

    if (toolbar == m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        InvalidateRect(toolbar, nullptr, TRUE);
        return;
    }

    InstallBreadcrumbSubclass(toolbar);
}

bool CExplorerBHO::HandleBreadcrumbPaint(HWND hwnd) {
    if ((!m_breadcrumbGradientEnabled && !m_breadcrumbFontGradientEnabled) || !m_gdiplusInitialized) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC target = BeginPaint(hwnd, &ps);
    if (!target) {
        return true;
    }

    RECT client{};
    GetClientRect(hwnd, &client);

    BP_PAINTPARAMS params{};
    params.cbSize = sizeof(params);
    params.dwFlags = BPPF_ERASE;

    HDC paintDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(target, &client, BPBF_TOPDOWNDIB, &params, &paintDc);
    HDC drawDc = paintDc ? paintDc : target;

    // Allow the control to render itself so layout, glyphs, and hover states stay intact.
    DefSubclassProc(hwnd, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(drawDc), PRF_CLIENT);

    Gdiplus::Graphics graphics(drawDc);
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        if (buffer) {
            EndBufferedPaint(buffer, TRUE);
        }
        EndPaint(hwnd, &ps);
        return true;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
    graphics.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

    HTHEME theme = nullptr;
    if (IsAppThemed() && IsThemeActive()) {
        theme = OpenThemeData(hwnd, L"BreadcrumbBar");
        if (!theme) {
            theme = OpenThemeData(hwnd, L"Toolbar");
        }
    }

    const COLORREF highlightBlendColor =
        theme ? GetThemeSysColor(theme, COLOR_HIGHLIGHTTEXT) : GetSysColor(COLOR_HIGHLIGHTTEXT);

    HFONT fontHandle = reinterpret_cast<HFONT>(SendMessage(hwnd, WM_GETFONT, 0, 0));
    if (!fontHandle) {
        fontHandle = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    Gdiplus::Font font(drawDc, fontHandle);
    if (font.GetLastStatus() != Gdiplus::Ok) {
        if (buffer) {
            EndBufferedPaint(buffer, TRUE);
        }
        EndPaint(hwnd, &ps);
        return true;
    }

    Gdiplus::StringFormat format;
    format.SetAlignment(Gdiplus::StringAlignmentNear);
    format.SetLineAlignment(Gdiplus::StringAlignmentCenter);
    format.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
    format.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);

    const int gradientTransparency = std::clamp(m_breadcrumbGradientTransparency, 0, 100);
    const int gradientOpacityPercent = 100 - gradientTransparency;
    const int fontTransparency = m_breadcrumbFontGradientEnabled ? std::clamp(m_breadcrumbFontTransparency, 0, 100) : 0;
    const int fontOpacityPercent = 100 - fontTransparency;
    const BYTE textAlphaBase = static_cast<BYTE>(std::clamp(fontOpacityPercent * 255 / 100, 0, 255));
    const bool shouldClearDefaultText = m_breadcrumbGradientEnabled || m_breadcrumbFontGradientEnabled;

    static const std::array<COLORREF, 7> kRainbowColors = {
        RGB(255, 59, 48),   // red
        RGB(255, 149, 0),   // orange
        RGB(255, 204, 0),   // yellow
        RGB(52, 199, 89),   // green
        RGB(0, 122, 255),   // blue
        RGB(88, 86, 214),   // indigo
        RGB(175, 82, 222)   // violet
    };

    HIMAGELIST imageList = reinterpret_cast<HIMAGELIST>(SendMessage(hwnd, TB_GETIMAGELIST, 0, 0));
    if (!imageList) {
        imageList = reinterpret_cast<HIMAGELIST>(SendMessage(hwnd, TB_GETIMAGELIST, 1, 0));
    }
    int imageWidth = 0;
    int imageHeight = 0;
    if (imageList) {
        ImageList_GetIconSize(imageList, &imageWidth, &imageHeight);
    }

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

        RECT itemRect{};
        if (!SendMessage(hwnd, TB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&itemRect))) {
            continue;
        }

        const bool isPressed = (button.fsState & TBSTATE_PRESSED) != 0;
        const bool isHot = !isPressed && (button.fsState & TBSTATE_HOT) != 0;
        const bool hasDropdown = (button.fsStyle & BTNS_DROPDOWN) != 0;
        const bool hasIcon = imageList && imageWidth > 0 && imageHeight > 0 && button.iBitmap >= 0 &&
                              button.iBitmap != I_IMAGENONE;
        const bool useFontGradient = m_breadcrumbFontGradientEnabled;

        COLORREF startRgb = 0;
        COLORREF endRgb = 0;
        if (m_useCustomBreadcrumbGradientColors) {
            startRgb = m_breadcrumbGradientStartColor;
            endRgb = m_breadcrumbGradientEndColor;
            ++colorIndex;
        } else {
            const size_t startIndex = static_cast<size_t>(colorIndex % kRainbowColors.size());
            const size_t endIndex = static_cast<size_t>((colorIndex + 1) % kRainbowColors.size());
            ++colorIndex;
            startRgb = kRainbowColors[startIndex];
            endRgb = kRainbowColors[endIndex];
        }

        auto brightenForState = [&](const Gdiplus::Color& color) -> Gdiplus::Color {
            if (!isHot && !isPressed) {
                return color;
            }
            const float blendFactor = isPressed ? 0.75f : 0.55f;
            const BYTE blendRed = GetRValue(highlightBlendColor);
            const BYTE blendGreen = GetGValue(highlightBlendColor);
            const BYTE blendBlue = GetBValue(highlightBlendColor);
            auto blendChannel = [&](BYTE base, BYTE blend) -> BYTE {
                const double result = static_cast<double>(base) +
                                      (static_cast<double>(blend) - static_cast<double>(base)) * blendFactor;
                return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(result)), 0, 255));
            };
            return Gdiplus::Color(color.GetA(), blendChannel(color.GetR(), blendRed),
                                  blendChannel(color.GetG(), blendGreen), blendChannel(color.GetB(), blendBlue));
        };

        auto darkenChannel = [](BYTE channel) -> BYTE {
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(channel) * 35 / 100, 0, 255));
        };
        auto lightenChannel = [](BYTE channel) -> BYTE {
            const double lightenFactor = 0.55;
            const double boosted = static_cast<double>(channel) +
                                   (static_cast<double>(255 - channel) * lightenFactor);
            return static_cast<BYTE>(
                std::clamp<int>(static_cast<int>(std::lround(boosted)), 0, 255));
        };
        auto averageChannel = [](BYTE a, BYTE b) -> BYTE {
            return static_cast<BYTE>((static_cast<int>(a) + static_cast<int>(b)) / 2);
        };

        Gdiplus::RectF rectF(static_cast<Gdiplus::REAL>(itemRect.left),
                             static_cast<Gdiplus::REAL>(itemRect.top),
                             static_cast<Gdiplus::REAL>(itemRect.right - itemRect.left),
                             static_cast<Gdiplus::REAL>(itemRect.bottom - itemRect.top));

        BYTE baseAlpha = 200;
        if (isPressed) {
            baseAlpha = 235;
        } else if (isHot) {
            baseAlpha = 220;
        }

        const BYTE scaledAlpha = static_cast<BYTE>(std::clamp<int>(baseAlpha * gradientOpacityPercent / 100, 0, 255));
        const bool backgroundGradientVisible = (m_breadcrumbGradientEnabled && scaledAlpha > 0);

        Gdiplus::Color backgroundGradientStartColor;
        Gdiplus::Color backgroundGradientEndColor;
        bool hasBackgroundGradientColors = false;
        Gdiplus::Color backgroundSolidColor;
        bool hasBackgroundSolidColor = false;

        if (backgroundGradientVisible) {
            backgroundGradientStartColor = brightenForState(
                Gdiplus::Color(scaledAlpha, darkenChannel(GetRValue(startRgb)),
                               darkenChannel(GetGValue(startRgb)), darkenChannel(GetBValue(startRgb))));
            backgroundGradientEndColor = brightenForState(
                Gdiplus::Color(scaledAlpha, darkenChannel(GetRValue(endRgb)),
                               darkenChannel(GetGValue(endRgb)), darkenChannel(GetBValue(endRgb))));
            hasBackgroundGradientColors = true;
            Gdiplus::LinearGradientBrush backgroundBrush(rectF, backgroundGradientStartColor,
                                                         backgroundGradientEndColor, Gdiplus::LinearGradientModeHorizontal);
            backgroundBrush.SetGammaCorrection(TRUE);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
            graphics.FillRectangle(&backgroundBrush, rectF);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
        }

        if (hasIcon) {
            const int iconX = itemRect.left + 4;
            const LONG verticalSpace = ((itemRect.bottom - itemRect.top) - imageHeight) / 2;
            const LONG iconYOffset = std::max<LONG>(0, verticalSpace);
            const int iconY = static_cast<int>(itemRect.top + iconYOffset);
            ImageList_Draw(imageList, button.iBitmap, drawDc, iconX, iconY, ILD_TRANSPARENT);
        }

        BYTE textAlpha = textAlphaBase;
        if (isPressed) {
            textAlpha = static_cast<BYTE>(std::min<int>(255, textAlpha + 60));
        } else if (isHot) {
            textAlpha = static_cast<BYTE>(std::min<int>(255, textAlpha + 35));
        }
        COLORREF fontStartRgb = m_useCustomBreadcrumbFontColors ? m_breadcrumbFontGradientStartColor : startRgb;
        COLORREF fontEndRgb = m_useCustomBreadcrumbFontColors ? m_breadcrumbFontGradientEndColor : endRgb;
        const BYTE lightStartRed = lightenChannel(GetRValue(fontStartRgb));
        const BYTE lightStartGreen = lightenChannel(GetGValue(fontStartRgb));
        const BYTE lightStartBlue = lightenChannel(GetBValue(fontStartRgb));
        const BYTE lightEndRed = lightenChannel(GetRValue(fontEndRgb));
        const BYTE lightEndGreen = lightenChannel(GetGValue(fontEndRgb));
        const BYTE lightEndBlue = lightenChannel(GetBValue(fontEndRgb));
        const Gdiplus::Color brightFontStart = brightenForState(
            Gdiplus::Color(textAlpha, lightStartRed, lightStartGreen, lightStartBlue));
        const Gdiplus::Color brightFontEnd = brightenForState(
            Gdiplus::Color(textAlpha, lightEndRed, lightEndGreen, lightEndBlue));

        auto computeOpaqueFontColor = [&](const Gdiplus::Color& fontColor, bool useStart) {
            if (textAlpha >= 255) {
                return Gdiplus::Color(255, fontColor.GetR(), fontColor.GetG(), fontColor.GetB());
            }
            const double opacity = static_cast<double>(textAlpha) / 255.0;
            int backgroundRed = 0;
            int backgroundGreen = 0;
            int backgroundBlue = 0;
            if (hasBackgroundGradientColors) {
                const Gdiplus::Color& background = useStart ? backgroundGradientStartColor : backgroundGradientEndColor;
                backgroundRed = background.GetR();
                backgroundGreen = background.GetG();
                backgroundBlue = background.GetB();
            } else if (hasBackgroundSolidColor) {
                backgroundRed = backgroundSolidColor.GetR();
                backgroundGreen = backgroundSolidColor.GetG();
                backgroundBlue = backgroundSolidColor.GetB();
            } else {
                backgroundRed = GetRValue(highlightBlendColor);
                backgroundGreen = GetGValue(highlightBlendColor);
                backgroundBlue = GetBValue(highlightBlendColor);
            }
            auto blendComponent = [&](int foreground, int background) -> BYTE {
                const double value = static_cast<double>(background) +
                                     (static_cast<double>(foreground) - static_cast<double>(background)) * opacity;
                return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(value)), 0, 255));
            };
            return Gdiplus::Color(255, blendComponent(fontColor.GetR(), backgroundRed),
                                  blendComponent(fontColor.GetG(), backgroundGreen),
                                  blendComponent(fontColor.GetB(), backgroundBlue));
        };

        Gdiplus::Color textPaintStart = computeOpaqueFontColor(brightFontStart, true);
        Gdiplus::Color textPaintEnd = computeOpaqueFontColor(brightFontEnd, false);

        constexpr int kTextPadding = 8;
        const int iconReserve = hasIcon ? (imageWidth + 6) : 0;
        const int dropdownReserve = hasDropdown ? 12 : 0;
        const int availableTextWidth = (itemRect.right - itemRect.left) - iconReserve - dropdownReserve -
                                       (kTextPadding * 2);
        const bool iconOnlyButton = hasIcon && availableTextWidth <= 4 && (button.fsStyle & BTNS_SHOWTEXT) == 0;

        LRESULT textLength = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand, 0);
        if (!iconOnlyButton && textLength > 0) {
            std::wstring text(static_cast<size_t>(textLength) + 1, L'\0');
            LRESULT copied = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand,
                                         reinterpret_cast<LPARAM>(text.data()));
            if (copied > 0) {
                text.resize(static_cast<size_t>(copied));

                RECT textRect = itemRect;
                RECT clearRect = itemRect;
                textRect.left += kTextPadding + iconReserve;
                textRect.right -= kTextPadding;
                clearRect.left += kTextPadding + iconReserve;
                if (hasDropdown) {
                    clearRect.right -= dropdownReserve;
                    textRect.right -= dropdownReserve;
                }

                if (textRect.right > textRect.left) {
                    Gdiplus::RectF textRectF(static_cast<Gdiplus::REAL>(textRect.left),
                                             static_cast<Gdiplus::REAL>(textRect.top),
                                             static_cast<Gdiplus::REAL>(textRect.right - textRect.left),
                                             static_cast<Gdiplus::REAL>(textRect.bottom - textRect.top));

                    if (shouldClearDefaultText && !backgroundGradientVisible) {
                        if (clearRect.right > clearRect.left) {
                            const COLORREF averageBackground = SampleAverageColor(drawDc, clearRect);
                            Gdiplus::RectF clearRectF(static_cast<Gdiplus::REAL>(clearRect.left),
                                                      static_cast<Gdiplus::REAL>(clearRect.top),
                                                      static_cast<Gdiplus::REAL>(clearRect.right - clearRect.left),
                                                      static_cast<Gdiplus::REAL>(clearRect.bottom - clearRect.top));
                            backgroundSolidColor = Gdiplus::Color(255, GetRValue(averageBackground),
                                                                  GetGValue(averageBackground),
                                                                  GetBValue(averageBackground));
                            hasBackgroundSolidColor = true;
                            textPaintStart = computeOpaqueFontColor(brightFontStart, true);
                            textPaintEnd = computeOpaqueFontColor(brightFontEnd, false);
                            Gdiplus::SolidBrush clearBrush(backgroundSolidColor);
                            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
                            graphics.FillRectangle(&clearBrush, clearRectF);
                            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                        }
                    }

                    if (textAlpha > 0) {
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
                        if (useFontGradient) {
                            Gdiplus::LinearGradientBrush textBrush(
                                textRectF,
                                textPaintStart,
                                textPaintEnd,
                                Gdiplus::LinearGradientModeHorizontal);
                            textBrush.SetGammaCorrection(TRUE);
                            graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                &textBrush);
                        } else {
                            const BYTE avgRed = averageChannel(brightFontStart.GetR(), brightFontEnd.GetR());
                            const BYTE avgGreen = averageChannel(brightFontStart.GetG(), brightFontEnd.GetG());
                            const BYTE avgBlue = averageChannel(brightFontStart.GetB(), brightFontEnd.GetB());
                            Gdiplus::Color solidColor = computeOpaqueFontColor(
                                Gdiplus::Color(textAlpha, avgRed, avgGreen, avgBlue), true);
                            Gdiplus::SolidBrush textBrush(solidColor);
                            graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                &textBrush);
                        }
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                    }
                }
            }
        }

        if (hasDropdown) {
            const float arrowWidth = 6.0f;
            const float arrowHeight = 4.0f;
            const float centerX = rectF.X + rectF.Width - 9.0f;
            const float centerY = rectF.Y + rectF.Height / 2.0f;

            if (isHot || isPressed) {
                const float highlightWidth = arrowWidth + 6.0f;
                const float highlightHeight = std::max(4.0f, rectF.Height - 4.0f);
                Gdiplus::RectF highlightRect(centerX - highlightWidth / 2.0f,
                                             rectF.Y + 2.0f,
                                             highlightWidth,
                                             highlightHeight);
                const BYTE highlightAlpha = static_cast<BYTE>(isPressed ? 160 : 130);
                Gdiplus::Color highlightColor = brightenForState(
                    Gdiplus::Color(highlightAlpha, lightEndRed, lightEndGreen, lightEndBlue));
                Gdiplus::SolidBrush highlightBrush(highlightColor);
                graphics.FillRectangle(&highlightBrush, highlightRect);
            }

            Gdiplus::PointF arrow[3] = {
                {centerX - arrowWidth / 2.0f, centerY - arrowHeight / 2.0f},
                {centerX + arrowWidth / 2.0f, centerY - arrowHeight / 2.0f},
                {centerX, centerY + arrowHeight / 2.0f},
            };

            Gdiplus::RectF arrowRect(centerX - arrowWidth / 2.0f,
                                     centerY - arrowHeight / 2.0f,
                                     arrowWidth,
                                     arrowHeight);
            const bool useArrowGradient = useFontGradient || m_breadcrumbGradientEnabled;
            const BYTE arrowAlphaBase = std::max(textAlpha,
                                                 static_cast<BYTE>(backgroundGradientVisible ? scaledAlpha : 0));
            const int arrowBoost = isPressed ? 60 : (isHot ? 35 : 15);
            const BYTE arrowAlpha = static_cast<BYTE>(std::min(255, static_cast<int>(arrowAlphaBase) + arrowBoost));
            const Gdiplus::Color arrowStartColor(arrowAlpha, textPaintStart.GetR(), textPaintStart.GetG(),
                                                 textPaintStart.GetB());
            const Gdiplus::Color arrowEndColor(arrowAlpha, textPaintEnd.GetR(), textPaintEnd.GetG(),
                                               textPaintEnd.GetB());
            if (useArrowGradient) {
                Gdiplus::LinearGradientBrush arrowBrush(arrowRect, arrowStartColor, arrowEndColor,
                                                        Gdiplus::LinearGradientModeHorizontal);
                arrowBrush.SetGammaCorrection(TRUE);
                graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
            } else {
                const BYTE arrowRed = averageChannel(arrowStartColor.GetR(), arrowEndColor.GetR());
                const BYTE arrowGreen = averageChannel(arrowStartColor.GetG(), arrowEndColor.GetG());
                const BYTE arrowBlue = averageChannel(arrowStartColor.GetB(), arrowEndColor.GetB());
                Gdiplus::SolidBrush arrowBrush(Gdiplus::Color(arrowAlpha, arrowRed, arrowGreen, arrowBlue));
                graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
            }
        }
    }

    if (theme) {
        CloseThemeData(theme);
    }

    if (buffer) {
        EndBufferedPaint(buffer, TRUE);
    }
    EndPaint(hwnd, &ps);
    return true;
}

LRESULT CALLBACK CExplorerBHO::BreadcrumbCbtProc(int code, WPARAM wParam, LPARAM lParam) {
    HHOOK hookHandle = nullptr;

    if (code == HCBT_CREATEWND) {
        HWND hwnd = reinterpret_cast<HWND>(wParam);
        const CBT_CREATEWNDW* create = reinterpret_cast<CBT_CREATEWNDW*>(lParam);

        const wchar_t* className = nullptr;
        if (create && create->lpcs) {
            if (HIWORD(create->lpcs->lpszClass)) {
                className = create->lpcs->lpszClass;
            }
        }

        if (className) {
            if (_wcsicmp(className, TOOLBARCLASSNAMEW) != 0) {
                return CallNextHookEx(nullptr, code, wParam, lParam);
            }
        } else {
            wchar_t buffer[64]{};
            if (GetClassNameW(hwnd, buffer, ARRAYSIZE(buffer)) <= 0 ||
                _wcsicmp(buffer, TOOLBARCLASSNAMEW) != 0) {
                return CallNextHookEx(nullptr, code, wParam, lParam);
            }
        }

        std::vector<CExplorerBHO*> observers;
        {
            std::lock_guard<std::mutex> lock(g_breadcrumbHookMutex);
            auto it = g_breadcrumbHooks.find(GetCurrentThreadId());
            if (it != g_breadcrumbHooks.end()) {
                observers = it->second.observers;
                hookHandle = it->second.hook;
            }
        }

        if (!observers.empty()) {
            for (CExplorerBHO* observer : observers) {
                if (!observer ||
                    (!observer->m_breadcrumbGradientEnabled && !observer->m_breadcrumbFontGradientEnabled) ||
                    !observer->m_gdiplusInitialized) {
                    continue;
                }

                HWND start = hwnd;
                if (create && create->lpcs && create->lpcs->hwndParent) {
                    start = create->lpcs->hwndParent;
                }

                if (!observer->IsBreadcrumbToolbarAncestor(start)) {
                    continue;
                }
                if (!observer->IsWindowOwnedByThisExplorer(hwnd)) {
                    continue;
                }

                if (observer->InstallBreadcrumbSubclass(hwnd)) {
                    observer->LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                                                 L"Breadcrumb toolbar subclassed via CBT hook (hwnd=%p)", hwnd);
                }
            }
        }
    }

    return CallNextHookEx(hookHandle, code, wParam, lParam);
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
        case WM_ERASEBKGND:
            if (self->m_breadcrumbGradientEnabled) {
                return 1;
            }
            break;
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_UPDATEUISTATE:
            if (self->m_breadcrumbGradientEnabled || self->m_breadcrumbFontGradientEnabled) {
                InvalidateRect(hwnd, nullptr, TRUE);
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

LRESULT CALLBACK CExplorerBHO::ExplorerViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                       UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    LRESULT result = 0;
    if (self->HandleExplorerViewMessage(hwnd, msg, wParam, lParam, &result)) {
        return result;
    }

    if (msg == WM_NCDESTROY) {
        if (hwnd == self->m_listView) {
            self->m_listView = nullptr;
            self->m_listViewSubclassInstalled = false;
        } else if (hwnd == self->m_treeView) {
            self->m_treeView = nullptr;
            self->m_treeViewSubclassInstalled = false;
        } else if (hwnd == self->m_frameWindow) {
            self->m_frameWindow = nullptr;
            self->m_frameSubclassInstalled = false;
        } else if (hwnd == self->m_shellViewWindow) {
            self->m_shellViewWindowSubclassInstalled = false;
            self->m_shellViewWindow = nullptr;
        }

        if (!self->m_listView && !self->m_treeView && !self->m_shellViewWindow) {
            self->m_shellView.Reset();
            self->ClearPendingOpenInNewTabState();
        }

        RemoveWindowSubclass(hwnd, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(self));
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

}  // namespace shelltabs

