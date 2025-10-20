#include "CExplorerBHO.h"

#include <combaseapi.h>
#include <exdispid.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shlguid.h>
#include <CommCtrl.h>
#include <uxtheme.h>
#include <OleIdl.h>
#include <gdiplus.h>

#include <algorithm>
#include <array>
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
                    UINT id = GetMenuItemID(menu, i);
                    if (id == UINT_MAX) {
                        MENUITEMINFOW info{};
                        info.cbSize = sizeof(info);
                        info.fMask = MIIM_ID;
                        if (GetMenuItemInfoW(menu, i, TRUE, &info)) {
                            id = info.wID;
                        } else {
                            id = 0;
                        }
                    }
                    *commandId = id;
                }
                return true;
            }
        }
    }

    return false;
}

constexpr wchar_t kOpenInNewTabLabel[] = L"Open In New Tab";

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
    if (m_listView && m_listViewSubclassInstalled && IsWindow(m_listView)) {
        RemoveWindowSubclass(m_listView, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }
    if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
        RemoveWindowSubclass(m_treeView, ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }

    m_shellViewWindowSubclassInstalled = false;
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
            HandleExplorerContextMenuInit(hwnd, reinterpret_cast<HMENU>(wParam));
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
    if (!FindOpenInNewWindowMenuItem(menu, &position, nullptr)) {
        LogMessage(LogLevel::Warning, L"Context menu init aborted: 'Open in new window' anchor not found");
        return;
    }

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

    if (!InsertMenuItemW(menu, position + 1, TRUE, &item)) {
        LogLastError(L"InsertMenuItem(Open In New Tab)", GetLastError());
        return;
    }

    m_pendingOpenInNewTabPaths = std::move(paths);
    m_contextMenuInserted = true;
    m_trackedContextMenu = menu;
    LogMessage(LogLevel::Info, L"Open In New Tab inserted at position %u for %zu paths", position + 1,
               m_pendingOpenInNewTabPaths.size());
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
    if (!m_shellView) {
        LogMessage(LogLevel::Warning, L"CollectSelectedFolderPaths failed: shell view unavailable");
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItemArray> items;
    HRESULT hr = m_shellView->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&items));
    if (FAILED(hr) || !items) {
        LogMessage(LogLevel::Warning, L"CollectSelectedFolderPaths failed: unable to query selection (hr=0x%08lX)", hr);
        return false;
    }

    DWORD count = 0;
    hr = items->GetCount(&count);
    if (FAILED(hr) || count == 0) {
        LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths found no selected items (hr=0x%08lX count=%lu)", hr,
                   count);
        return false;
    }

    if (count > kMaxTrackedSelection) {
        LogMessage(LogLevel::Warning, L"CollectSelectedFolderPaths aborted: selection too large (%lu)", count);
        return false;
    }

    paths.reserve(static_cast<size_t>(count));

    for (DWORD index = 0; index < count; ++index) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(index, &item)) || !item) {
            LogMessage(LogLevel::Warning, L"CollectSelectedFolderPaths failed: unable to access item %lu", index);
            return false;
        }

        SFGAOF attributes = 0;
        hr = item->GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM, &attributes);
        if (FAILED(hr) || (attributes & SFGAO_FOLDER) == 0 || (attributes & SFGAO_FILESYSTEM) == 0) {
            LogMessage(LogLevel::Info,
                       L"CollectSelectedFolderPaths skipping item %lu: attributes=0x%08lX (hr=0x%08lX)", index,
                       attributes, hr);
            return false;
        }

        PWSTR path = nullptr;
        hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path);
        if (FAILED(hr) || !path || path[0] == L'\0') {
            if (path) {
                CoTaskMemFree(path);
            }
            LogMessage(LogLevel::Warning,
                       L"CollectSelectedFolderPaths failed: unable to resolve file system path for item %lu (hr=0x%08lX)",
                       index, hr);
            return false;
        }

        std::wstring value(path);
        CoTaskMemFree(path);

        if (value.empty()) {
            LogMessage(LogLevel::Warning, L"CollectSelectedFolderPaths failed: empty path for item %lu", index);
            return false;
        }

        paths.push_back(std::move(value));
    }

    LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths captured %zu path(s)", paths.size());
    return !paths.empty();
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
    format.SetAlignment(Gdiplus::StringAlignmentCenter);
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

        const size_t startIndex = static_cast<size_t>(colorIndex % kRainbowColors.size());
        const size_t endIndex = static_cast<size_t>((colorIndex + 1) % kRainbowColors.size());
        ++colorIndex;

        const COLORREF startRgb = kRainbowColors[startIndex];
        const COLORREF endRgb = kRainbowColors[endIndex];

        auto darkenChannel = [](BYTE channel) -> BYTE {
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(channel) * 35 / 100, 0, 255));
        };
        auto lightenChannel = [](BYTE channel) -> BYTE {
            const int boosted = channel + ((255 - channel) * 3) / 4;
            return static_cast<BYTE>(std::clamp<int>(boosted, 0, 255));
        };
        auto averageChannel = [](BYTE a, BYTE b) -> BYTE {
            return static_cast<BYTE>((static_cast<int>(a) + static_cast<int>(b)) / 2);
        };

        Gdiplus::RectF rectF(static_cast<Gdiplus::REAL>(itemRect.left),
                             static_cast<Gdiplus::REAL>(itemRect.top),
                             static_cast<Gdiplus::REAL>(itemRect.right - itemRect.left),
                             static_cast<Gdiplus::REAL>(itemRect.bottom - itemRect.top));

        BYTE baseAlpha = 200;
        if ((button.fsState & TBSTATE_PRESSED) != 0) {
            baseAlpha = 235;
        } else if ((button.fsState & TBSTATE_HOT) != 0) {
            baseAlpha = 220;
        }

        const BYTE scaledAlpha = static_cast<BYTE>(std::clamp<int>(baseAlpha * gradientOpacityPercent / 100, 0, 255));
        const bool backgroundGradientVisible = (m_breadcrumbGradientEnabled && scaledAlpha > 0);

        if (backgroundGradientVisible) {
            const Gdiplus::Color startColor(scaledAlpha, darkenChannel(GetRValue(startRgb)),
                                            darkenChannel(GetGValue(startRgb)),
                                            darkenChannel(GetBValue(startRgb)));
            const Gdiplus::Color endColor(scaledAlpha, darkenChannel(GetRValue(endRgb)),
                                          darkenChannel(GetGValue(endRgb)), darkenChannel(GetBValue(endRgb)));
            Gdiplus::LinearGradientBrush backgroundBrush(rectF, startColor, endColor,
                                                         Gdiplus::LinearGradientModeHorizontal);
            backgroundBrush.SetGammaCorrection(TRUE);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
            graphics.FillRectangle(&backgroundBrush, rectF);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
        }

        const BYTE textAlpha = textAlphaBase;

        LRESULT textLength = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand, 0);
        if (textLength > 0) {
            std::wstring text(static_cast<size_t>(textLength) + 1, L'\0');
            LRESULT copied = SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand,
                                         reinterpret_cast<LPARAM>(text.data()));
            if (copied > 0) {
                text.resize(static_cast<size_t>(copied));

                RECT textRect = itemRect;
                constexpr int kPadding = 8;
                textRect.left += kPadding;
                textRect.right -= kPadding;
                if ((button.fsStyle & BTNS_DROPDOWN) != 0) {
                    textRect.right -= 12;
                }

                if (textRect.right > textRect.left) {
                    const BYTE lightStartRed = lightenChannel(GetRValue(startRgb));
                    const BYTE lightStartGreen = lightenChannel(GetGValue(startRgb));
                    const BYTE lightStartBlue = lightenChannel(GetBValue(startRgb));
                    const BYTE lightEndRed = lightenChannel(GetRValue(endRgb));
                    const BYTE lightEndGreen = lightenChannel(GetGValue(endRgb));
                    const BYTE lightEndBlue = lightenChannel(GetBValue(endRgb));

                    const bool useFontGradient = m_breadcrumbFontGradientEnabled;
                    Gdiplus::RectF textRectF(static_cast<Gdiplus::REAL>(textRect.left),
                                             static_cast<Gdiplus::REAL>(textRect.top),
                                             static_cast<Gdiplus::REAL>(textRect.right - textRect.left),
                                             static_cast<Gdiplus::REAL>(textRect.bottom - textRect.top));

                    if (shouldClearDefaultText && !backgroundGradientVisible) {
                        const COLORREF averageBackground = SampleAverageColor(drawDc, textRect);
                        Gdiplus::SolidBrush clearBrush(Gdiplus::Color(255, GetRValue(averageBackground),
                                                                      GetGValue(averageBackground),
                                                                      GetBValue(averageBackground)));
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
                        graphics.FillRectangle(&clearBrush, textRectF);
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                    }

                    if (textAlpha > 0) {
                        if (useFontGradient) {
                            Gdiplus::LinearGradientBrush textBrush(
                                textRectF,
                                Gdiplus::Color(textAlpha, lightStartRed, lightStartGreen, lightStartBlue),
                                Gdiplus::Color(textAlpha, lightEndRed, lightEndGreen, lightEndBlue),
                                Gdiplus::LinearGradientModeHorizontal);
                            textBrush.SetGammaCorrection(TRUE);
                            graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                &textBrush);
                        } else {
                            const BYTE avgRed = averageChannel(lightStartRed, lightEndRed);
                            const BYTE avgGreen = averageChannel(lightStartGreen, lightEndGreen);
                            const BYTE avgBlue = averageChannel(lightStartBlue, lightEndBlue);
                            Gdiplus::SolidBrush textBrush(Gdiplus::Color(textAlpha, avgRed, avgGreen, avgBlue));
                            graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                &textBrush);
                        }
                    }
                }
            }
        }

        if ((button.fsStyle & BTNS_DROPDOWN) != 0) {
            const float arrowWidth = 6.0f;
            const float arrowHeight = 4.0f;
            const float centerX = rectF.X + rectF.Width - 9.0f;
            const float centerY = rectF.Y + rectF.Height / 2.0f;

            Gdiplus::PointF arrow[3] = {
                {centerX - arrowWidth / 2.0f, centerY - arrowHeight / 2.0f},
                {centerX + arrowWidth / 2.0f, centerY - arrowHeight / 2.0f},
                {centerX, centerY + arrowHeight / 2.0f},
            };

            Gdiplus::SolidBrush arrowBrush(Gdiplus::Color(220, 255, 255, 255));
            graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
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

