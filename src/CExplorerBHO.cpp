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
#include <limits>
#include <string>
#include <unordered_map>
#include <vector>
#include <wrl/client.h>
#include <shobjidl_core.h>

#include "BackgroundCache.h"
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

BYTE AverageColorChannel(BYTE a, BYTE b) {
    return static_cast<BYTE>((static_cast<int>(a) + static_cast<int>(b)) / 2);
}

Gdiplus::Color BrightenBreadcrumbColor(const Gdiplus::Color& color,
                                       bool isHot,
                                       bool isPressed,
                                       COLORREF highlightBackgroundColor) {
    if (!isHot && !isPressed) {
        return color;
    }

    const float blendFactor = isPressed ? 0.75f : 0.55f;
    const BYTE blendRed = GetRValue(highlightBackgroundColor);
    const BYTE blendGreen = GetGValue(highlightBackgroundColor);
    const BYTE blendBlue = GetBValue(highlightBackgroundColor);

    auto blendChannel = [&](BYTE base, BYTE blend) -> BYTE {
        const double result = static_cast<double>(base) +
                              (static_cast<double>(blend) - static_cast<double>(base)) * blendFactor;
        return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(result)), 0, 255));
    };

    return Gdiplus::Color(color.GetA(), blendChannel(color.GetR(), blendRed),
                          blendChannel(color.GetG(), blendGreen), blendChannel(color.GetB(), blendBlue));
}

constexpr wchar_t kOpenInNewTabLabel[] = L"Open in new tab";
constexpr int kProgressGradientSampleWidth = 256;

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
    DestroyProgressGradientResources();
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
    RemoveProgressSubclass();
    RemoveAddressEditSubclass();
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
    ClearFolderBackgrounds();
    m_currentFolderKey.clear();
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

HWND CExplorerBHO::FindProgressWindow() const {
    if (m_breadcrumbToolbar && IsWindow(m_breadcrumbToolbar)) {
        HWND breadcrumbParent = GetParent(m_breadcrumbToolbar);
        if (breadcrumbParent) {
            if (MatchesClass(breadcrumbParent, PROGRESS_CLASSW)) {
                return breadcrumbParent;
            }
            HWND progressParent = GetParent(breadcrumbParent);
            if (progressParent && MatchesClass(progressParent, PROGRESS_CLASSW)) {
                return progressParent;
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return nullptr;
    }

    struct EnumData {
        HWND progress = nullptr;
    } data{};

    EnumChildWindows(
        frame,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* data = reinterpret_cast<EnumData*>(param);
            if (!data || data->progress) {
                return FALSE;
            }
            if (!MatchesClass(hwnd, PROGRESS_CLASSW)) {
                return TRUE;
            }
            if (!FindDescendantWindow(hwnd, L"Breadcrumb Parent")) {
                return TRUE;
            }
            data->progress = hwnd;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    return data.progress;
}

HWND CExplorerBHO::FindAddressEditControl() const {
    auto resolveEdit = [&](HWND window) -> HWND {
        if (!window || !IsWindow(window)) {
            return nullptr;
        }
        HWND edit = nullptr;
        if (MatchesClass(window, L"ComboBoxEx32")) {
            edit = reinterpret_cast<HWND>(SendMessageW(window, CBEM_GETEDITCONTROL, 0, 0));
            if (!edit) {
                edit = FindDescendantWindow(window, L"Edit");
            }
        } else if (MatchesClass(window, L"Edit")) {
            edit = window;
        } else {
            edit = FindDescendantWindow(window, L"Edit");
        }
        if (!edit || !IsWindow(edit)) {
            return nullptr;
        }
        if (!MatchesClass(edit, L"Edit")) {
            return nullptr;
        }
        if (!IsBreadcrumbToolbarAncestor(edit) || !IsWindowOwnedByThisExplorer(edit)) {
            return nullptr;
        }
        return edit;
    };

    if (m_breadcrumbToolbar && IsWindow(m_breadcrumbToolbar)) {
        if (HWND parent = GetParent(m_breadcrumbToolbar)) {
            if (HWND edit = resolveEdit(parent)) {
                return edit;
            }
            if (HWND grandparent = GetParent(parent)) {
                if (HWND edit = resolveEdit(grandparent)) {
                    return edit;
                }
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return nullptr;
    }

    struct EnumData {
        const CExplorerBHO* self = nullptr;
        HWND edit = nullptr;
    } data{this, nullptr};

    EnumChildWindows(
        frame,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* data = reinterpret_cast<EnumData*>(param);
            if (!data || data->edit) {
                return FALSE;
            }
            if (!MatchesClass(hwnd, L"ComboBoxEx32") && !MatchesClass(hwnd, L"Edit")) {
                return TRUE;
            }
            HWND edit = nullptr;
            if (MatchesClass(hwnd, L"ComboBoxEx32")) {
                edit = reinterpret_cast<HWND>(SendMessageW(hwnd, CBEM_GETEDITCONTROL, 0, 0));
                if (!edit) {
                    edit = FindDescendantWindow(hwnd, L"Edit");
                }
            } else {
                edit = hwnd;
            }
            if (!edit || !IsWindow(edit) || !MatchesClass(edit, L"Edit")) {
                return TRUE;
            }
            if (!data->self->IsBreadcrumbToolbarAncestor(edit) ||
                !data->self->IsWindowOwnedByThisExplorer(edit)) {
                return TRUE;
            }
            data->edit = edit;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    return data.edit;
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
    UpdateCurrentFolderBackground();
}

bool CExplorerBHO::InstallExplorerViewSubclass(HWND viewWindow, HWND listView, HWND treeView) {
    bool installed = false;

    if (viewWindow && IsWindow(viewWindow)) {
        if (SetWindowSubclass(viewWindow, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this),
                              0)) {
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
        if (SetWindowSubclass(frameWindow, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this),
                              0)) {
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
        if (SetWindowSubclass(listView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
            m_listView = listView;
            m_listViewSubclassInstalled = true;
            installed = true;
            LogMessage(LogLevel::Info, L"Installed explorer list view subclass (list=%p)", listView);
        } else {
            LogLastError(L"SetWindowSubclass(list view)", GetLastError());
        }
    }

    if (treeView && treeView != listView && IsWindow(treeView)) {
        if (SetWindowSubclass(treeView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
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
        RemoveWindowSubclass(m_shellViewWindow, &CExplorerBHO::ExplorerViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }
    if (m_frameWindow && m_frameSubclassInstalled && IsWindow(m_frameWindow)) {
        RemoveWindowSubclass(m_frameWindow, &CExplorerBHO::ExplorerViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }
    if (m_listView && m_listViewSubclassInstalled && IsWindow(m_listView)) {
        RemoveWindowSubclass(m_listView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }
    if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
        RemoveWindowSubclass(m_treeView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
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

void CExplorerBHO::ClearFolderBackgrounds() {
    m_folderBackgroundBitmaps.clear();
    m_universalBackgroundBitmap.reset();
    m_folderBackgroundsEnabled = false;
}

std::wstring CExplorerBHO::NormalizeBackgroundKey(const std::wstring& path) const {
    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }

    std::transform(normalized.begin(), normalized.end(), normalized.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });
    return normalized;
}

void CExplorerBHO::ReloadFolderBackgrounds(const ShellTabsOptions& options) {
    ClearFolderBackgrounds();

    if (!m_gdiplusInitialized) {
        return;
    }

    if (!options.enableFolderBackgrounds) {
        InvalidateFolderBackgroundTargets();
        return;
    }

    m_folderBackgroundsEnabled = true;

    if (!options.universalFolderBackgroundImage.cachedImagePath.empty()) {
        auto bitmap = LoadBackgroundBitmap(options.universalFolderBackgroundImage.cachedImagePath);
        if (bitmap) {
            m_universalBackgroundBitmap = std::move(bitmap);
        } else {
            LogMessage(LogLevel::Warning, L"Failed to load universal folder background from %ls",
                       options.universalFolderBackgroundImage.cachedImagePath.c_str());
        }
    }

    for (const auto& entry : options.folderBackgroundEntries) {
        if (entry.folderPath.empty() || entry.image.cachedImagePath.empty()) {
            continue;
        }

        std::wstring key = NormalizeBackgroundKey(entry.folderPath);
        if (key.empty()) {
            continue;
        }

        if (m_folderBackgroundBitmaps.find(key) != m_folderBackgroundBitmaps.end()) {
            continue;
        }

        auto bitmap = LoadBackgroundBitmap(entry.image.cachedImagePath);
        if (!bitmap) {
            LogMessage(LogLevel::Warning, L"Failed to load background for %ls from %ls", entry.folderPath.c_str(),
                       entry.image.cachedImagePath.c_str());
            continue;
        }

        m_folderBackgroundBitmaps.emplace(std::move(key), std::move(bitmap));
    }

    InvalidateFolderBackgroundTargets();
}

Gdiplus::Bitmap* CExplorerBHO::ResolveCurrentFolderBackground() const {
    if (!m_folderBackgroundsEnabled || !m_gdiplusInitialized) {
        return nullptr;
    }

    if (!m_currentFolderKey.empty()) {
        auto it = m_folderBackgroundBitmaps.find(m_currentFolderKey);
        if (it != m_folderBackgroundBitmaps.end() && it->second) {
            return it->second.get();
        }
    }

    if (m_universalBackgroundBitmap) {
        return m_universalBackgroundBitmap.get();
    }

    return nullptr;
}

bool CExplorerBHO::DrawFolderBackground(HWND hwnd, HDC dc) const {
    if (!hwnd || !dc || !m_folderBackgroundsEnabled || !m_gdiplusInitialized) {
        return false;
    }

    Gdiplus::Bitmap* background = ResolveCurrentFolderBackground();
    if (!background) {
        return false;
    }

    const UINT width = background->GetWidth();
    const UINT height = background->GetHeight();
    if (width == 0 || height == 0) {
        return false;
    }

    if (width > static_cast<UINT>(std::numeric_limits<INT>::max()) ||
        height > static_cast<UINT>(std::numeric_limits<INT>::max())) {
        return false;
    }

    RECT client{};
    if (!GetClientRect(hwnd, &client) || client.right <= client.left || client.bottom <= client.top) {
        return false;
    }

    Gdiplus::Graphics graphics(dc);
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        return false;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);

    const Gdiplus::Rect destRect(client.left, client.top, client.right - client.left, client.bottom - client.top);
    const Gdiplus::Status status =
        graphics.DrawImage(background, destRect, 0, 0, static_cast<INT>(width),
                           static_cast<INT>(height), Gdiplus::UnitPixel);
    return status == Gdiplus::Ok;
}

void CExplorerBHO::UpdateCurrentFolderBackground() {
    std::wstring newKey;

    if (m_folderBackgroundsEnabled && m_shellBrowser) {
        UniquePidl current = GetCurrentFolderPidL(m_shellBrowser, m_webBrowser);
        if (current) {
            PWSTR path = nullptr;
            if (SUCCEEDED(SHGetNameFromIDList(current.get(), SIGDN_FILESYSPATH, &path)) && path && path[0] != L'\0') {
                newKey = NormalizeBackgroundKey(path);
            }
            if (path) {
                CoTaskMemFree(path);
            }
        }
    }

    if (newKey == m_currentFolderKey) {
        return;
    }

    m_currentFolderKey = std::move(newKey);
    InvalidateFolderBackgroundTargets();
}

void CExplorerBHO::InvalidateFolderBackgroundTargets() const {
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, TRUE);
    }
    if (m_shellViewWindow && IsWindow(m_shellViewWindow)) {
        InvalidateRect(m_shellViewWindow, nullptr, TRUE);
    }
    if (m_frameWindow && IsWindow(m_frameWindow)) {
        InvalidateRect(m_frameWindow, nullptr, TRUE);
    }
}

bool CExplorerBHO::HandleExplorerViewMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                             LRESULT* result) {
    if (!result) {
        return false;
    }

    const bool isListView = (hwnd == m_listView);
    const bool isDirectUi = MatchesClass(hwnd, L"DirectUIHWND");
    const bool handlesBackground = isListView || isDirectUi;

    const UINT optionsChangedMessage = GetOptionsChangedMessage();
    if (optionsChangedMessage != 0 && msg == optionsChangedMessage) {
        UpdateBreadcrumbSubclass();
        if (m_breadcrumbToolbar && m_breadcrumbSubclassInstalled && IsWindow(m_breadcrumbToolbar)) {
            InvalidateRect(m_breadcrumbToolbar, nullptr, TRUE);
        }
        UpdateCurrentFolderBackground();
        InvalidateFolderBackgroundTargets();
        *result = 0;
        return true;
    }

    switch (msg) {
        case WM_ERASEBKGND: {
            if (handlesBackground) {
                HDC dc = reinterpret_cast<HDC>(wParam);
                if (dc && DrawFolderBackground(hwnd, dc)) {
                    *result = 1;
                    return true;
                }
            }
            break;
        }
        case WM_PAINT: {
            if (handlesBackground) {
                if (wParam) {
                    DrawFolderBackground(hwnd, reinterpret_cast<HDC>(wParam));
                } else {
                    RECT update{};
                    if (GetUpdateRect(hwnd, &update, FALSE)) {
                        HDC dc = GetDCEx(hwnd, nullptr, DCX_CACHE | DCX_CLIPCHILDREN | DCX_CLIPSIBLINGS | DCX_WINDOW);
                        if (dc) {
                            DrawFolderBackground(hwnd, dc);
                            ReleaseDC(hwnd, dc);
                        }
                    }
                }
            }
            break;
        }
        case WM_THEMECHANGED:
        case WM_SIZE: {
            if (handlesBackground) {
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        }
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

    if (SetWindowSubclass(toolbar, &CExplorerBHO::BreadcrumbSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_breadcrumbToolbar = toolbar;
        m_breadcrumbSubclassInstalled = true;
        m_loggedBreadcrumbToolbarMissing = false;
        LogMessage(LogLevel::Info, L"Installed breadcrumb gradient subclass on hwnd=%p", toolbar);
        InvalidateRect(toolbar, nullptr, TRUE);
        UpdateAddressEditSubclass();
        return true;
    }

    LogLastError(L"SetWindowSubclass(breadcrumb toolbar)", GetLastError());
    return false;
}

bool CExplorerBHO::InstallProgressSubclass(HWND progressWindow) {
    if (!progressWindow || !IsWindow(progressWindow)) {
        return false;
    }

    if (SetWindowSubclass(progressWindow, &CExplorerBHO::ProgressSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_progressWindow = progressWindow;
        m_progressSubclassInstalled = true;
        if (!EnsureProgressGradientResources()) {
            LogMessage(LogLevel::Warning,
                       L"Progress gradient resources unavailable; falling back to on-demand rendering");
        }
        LogMessage(LogLevel::Info, L"Installed progress gradient subclass on hwnd=%p", progressWindow);
        return true;
    }

    LogLastError(L"SetWindowSubclass(progress window)", GetLastError());
    return false;
}

void CExplorerBHO::RemoveBreadcrumbSubclass() {
    if (m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        if (IsWindow(m_breadcrumbToolbar)) {
            RemoveWindowSubclass(m_breadcrumbToolbar, &CExplorerBHO::BreadcrumbSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_breadcrumbToolbar, nullptr, TRUE);
        }
    }
    m_breadcrumbToolbar = nullptr;
    m_breadcrumbSubclassInstalled = false;
    if (m_breadcrumbLogState == BreadcrumbLogState::Searching) {
        m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    }
    m_loggedBreadcrumbToolbarMissing = false;

    RemoveAddressEditSubclass();
    RemoveProgressSubclass();
}

void CExplorerBHO::RemoveProgressSubclass() {
    if (m_progressWindow && m_progressSubclassInstalled) {
        if (IsWindow(m_progressWindow)) {
            RemoveWindowSubclass(m_progressWindow, &CExplorerBHO::ProgressSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_progressWindow, nullptr, TRUE);
        }
    }
    m_progressWindow = nullptr;
    m_progressSubclassInstalled = false;
    DestroyProgressGradientResources();
}

bool CExplorerBHO::EnsureProgressGradientResources() {
    if (!m_useCustomProgressGradientColors) {
        return false;
    }

    if (m_progressGradientBitmap && m_progressGradientBitmapStartColor == m_progressGradientStartColor &&
        m_progressGradientBitmapEndColor == m_progressGradientEndColor && m_progressGradientBits) {
        return true;
    }

    DestroyProgressGradientResources();

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = kProgressGradientSampleWidth;
    info.bmiHeader.biHeight = -1;  // top-down
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap || !bits) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return false;
    }

    auto* pixels = static_cast<DWORD*>(bits);
    const BYTE startRed = GetRValue(m_progressGradientStartColor);
    const BYTE startGreen = GetGValue(m_progressGradientStartColor);
    const BYTE startBlue = GetBValue(m_progressGradientStartColor);
    const BYTE endRed = GetRValue(m_progressGradientEndColor);
    const BYTE endGreen = GetGValue(m_progressGradientEndColor);
    const BYTE endBlue = GetBValue(m_progressGradientEndColor);

    for (int x = 0; x < kProgressGradientSampleWidth; ++x) {
        const double t = (kProgressGradientSampleWidth > 1)
                             ? static_cast<double>(x) / static_cast<double>(kProgressGradientSampleWidth - 1)
                             : 0.0;
        const BYTE red = static_cast<BYTE>(std::clamp<int>(
            static_cast<int>(std::lround(static_cast<double>(startRed) +
                                         (static_cast<double>(endRed) - static_cast<double>(startRed)) * t)),
            0, 255));
        const BYTE green = static_cast<BYTE>(std::clamp<int>(
            static_cast<int>(std::lround(static_cast<double>(startGreen) +
                                         (static_cast<double>(endGreen) - static_cast<double>(startGreen)) * t)),
            0, 255));
        const BYTE blue = static_cast<BYTE>(std::clamp<int>(
            static_cast<int>(std::lround(static_cast<double>(startBlue) +
                                         (static_cast<double>(endBlue) - static_cast<double>(startBlue)) * t)),
            0, 255));
        pixels[x] = (static_cast<DWORD>(blue)) | (static_cast<DWORD>(green) << 8) | (static_cast<DWORD>(red) << 16) |
                    0xFF000000;
    }

    m_progressGradientBitmap = bitmap;
    m_progressGradientBits = bits;
    m_progressGradientInfo = info;
    m_progressGradientBitmapStartColor = m_progressGradientStartColor;
    m_progressGradientBitmapEndColor = m_progressGradientEndColor;
    return true;
}

void CExplorerBHO::DestroyProgressGradientResources() {
    if (m_progressGradientBitmap) {
        DeleteObject(m_progressGradientBitmap);
        m_progressGradientBitmap = nullptr;
    }
    m_progressGradientBits = nullptr;
    m_progressGradientInfo = BITMAPINFO{};
    m_progressGradientBitmapStartColor = 0;
    m_progressGradientBitmapEndColor = 0;
}

bool CExplorerBHO::InstallAddressEditSubclass(HWND editWindow) {
    if (!editWindow || !IsWindow(editWindow)) {
        return false;
    }

    if (SetWindowSubclass(editWindow, &CExplorerBHO::AddressEditSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_addressEditWindow = editWindow;
        m_addressEditSubclassInstalled = true;
        LogMessage(LogLevel::Info, L"Installed address edit gradient subclass on hwnd=%p", editWindow);
        return true;
    }

    LogLastError(L"SetWindowSubclass(address edit)", GetLastError());
    return false;
}

void CExplorerBHO::RemoveAddressEditSubclass() {
    if (m_addressEditWindow && m_addressEditSubclassInstalled) {
        if (IsWindow(m_addressEditWindow)) {
            RemoveWindowSubclass(m_addressEditWindow, &CExplorerBHO::AddressEditSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_addressEditWindow, nullptr, TRUE);
        }
    }
    m_addressEditWindow = nullptr;
    m_addressEditSubclassInstalled = false;
}

void CExplorerBHO::RequestAddressEditRedraw(HWND hwnd) const {
    if (!hwnd || !IsWindow(hwnd)) {
        return;
    }

    if (!m_breadcrumbFontGradientEnabled || !m_useCustomBreadcrumbFontColors) {
        return;
    }

    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
}

void CExplorerBHO::UpdateAddressEditSubclass() {
    if (!m_breadcrumbFontGradientEnabled || !m_useCustomBreadcrumbFontColors || !m_gdiplusInitialized) {
        RemoveAddressEditSubclass();
        return;
    }

    HWND edit = FindAddressEditControl();
    if (!edit) {
        RemoveAddressEditSubclass();
        return;
    }

    if (m_addressEditSubclassInstalled && edit == m_addressEditWindow && IsWindow(edit)) {
        InvalidateRect(edit, nullptr, TRUE);
        return;
    }

    RemoveAddressEditSubclass();
    if (InstallAddressEditSubclass(edit)) {
        InvalidateRect(edit, nullptr, TRUE);
    }
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
    m_breadcrumbFontBrightness = std::clamp(options.breadcrumbFontBrightness, 0, 100);
    m_useCustomBreadcrumbGradientColors = options.useCustomBreadcrumbGradientColors;
    m_breadcrumbGradientStartColor = options.breadcrumbGradientStartColor;
    m_breadcrumbGradientEndColor = options.breadcrumbGradientEndColor;
    m_useCustomBreadcrumbFontColors = options.useCustomBreadcrumbFontColors;
    m_breadcrumbFontGradientStartColor = options.breadcrumbFontGradientStartColor;
    m_breadcrumbFontGradientEndColor = options.breadcrumbFontGradientEndColor;
    m_useCustomProgressGradientColors = options.useCustomProgressBarGradientColors;
    m_progressGradientStartColor = options.progressBarGradientStartColor;
    m_progressGradientEndColor = options.progressBarGradientEndColor;

    ReloadFolderBackgrounds(options);
    UpdateCurrentFolderBackground();

    UpdateProgressSubclass();

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
        UpdateProgressSubclass();
        UpdateAddressEditSubclass();
        return;
    }

    InstallBreadcrumbSubclass(toolbar);
    UpdateProgressSubclass();
    UpdateAddressEditSubclass();
}

void CExplorerBHO::UpdateProgressSubclass() {
    if (!m_useCustomProgressGradientColors) {
        if (m_progressSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Progress gradients disabled; removing subclass");
        }
        RemoveProgressSubclass();
        return;
    }

    HWND progress = FindProgressWindow();
    if (!progress) {
        if (m_progressSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Progress window not found; removing subclass");
        }
        RemoveProgressSubclass();
        return;
    }

    if (m_progressSubclassInstalled && progress == m_progressWindow) {
        InvalidateRect(progress, nullptr, TRUE);
        return;
    }

    RemoveProgressSubclass();
    if (InstallProgressSubclass(progress)) {
        InvalidateRect(progress, nullptr, TRUE);
    }
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

    HRESULT backgroundResult = DrawThemeParentBackground(hwnd, drawDc, &client);
    if (FAILED(backgroundResult)) {
        HBRUSH brush = reinterpret_cast<HBRUSH>(GetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND));
        if (!brush) {
            brush = GetSysColorBrush(COLOR_WINDOW);
        }
        FillRect(drawDc, &client, brush);
    }

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

    auto drawDropdownArrow = [&](const RECT& buttonRect, bool hot, bool pressed, BYTE textAlphaValue,
                                 const Gdiplus::Color& brightFontEndColor, const Gdiplus::Color& arrowTextStart,
                                 const Gdiplus::Color& arrowTextEnd, bool fontGradientEnabled,
                                 bool backgroundGradientEnabled, bool backgroundGradientVisible, BYTE gradientAlpha,
                                 COLORREF highlightColorRef) {
        const float arrowWidth = 6.0f;
        const float arrowHeight = 4.0f;
        const float rectWidth = static_cast<float>(buttonRect.right - buttonRect.left);
        const float rectHeight = static_cast<float>(buttonRect.bottom - buttonRect.top);
        const float centerX = static_cast<float>(buttonRect.left) + rectWidth - 9.0f;
        const float centerY = static_cast<float>(buttonRect.top) + rectHeight / 2.0f;

        if (hot || pressed) {
            const float highlightWidth = arrowWidth + 6.0f;
            const float highlightHeight = rectHeight > 4.0f ? (rectHeight - 4.0f) : 4.0f;
            Gdiplus::RectF highlightRect(centerX - highlightWidth / 2.0f,
                                         static_cast<Gdiplus::REAL>(buttonRect.top + 2),
                                         highlightWidth,
                                         highlightHeight);
            const BYTE highlightAlpha = static_cast<BYTE>(pressed ? 160 : 130);
            const Gdiplus::Color highlightBase(highlightAlpha, brightFontEndColor.GetR(),
                                               brightFontEndColor.GetG(), brightFontEndColor.GetB());
            Gdiplus::Color highlightColor =
                BrightenBreadcrumbColor(highlightBase, hot, pressed, highlightColorRef);
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
        const bool useArrowGradient = fontGradientEnabled || backgroundGradientEnabled;
        BYTE arrowAlphaBase = textAlphaValue;
        if (backgroundGradientVisible && gradientAlpha > arrowAlphaBase) {
            arrowAlphaBase = gradientAlpha;
        }
        const int arrowBoost = pressed ? 60 : (hot ? 35 : 15);
        const int boostedAlpha = static_cast<int>(arrowAlphaBase) + arrowBoost;
        const BYTE arrowAlpha = static_cast<BYTE>(boostedAlpha > 255 ? 255 : boostedAlpha);
        const Gdiplus::Color arrowStartColor(arrowAlpha, arrowTextStart.GetR(), arrowTextStart.GetG(),
                                             arrowTextStart.GetB());
        const Gdiplus::Color arrowEndColor(arrowAlpha, arrowTextEnd.GetR(), arrowTextEnd.GetG(),
                                           arrowTextEnd.GetB());
        if (useArrowGradient) {
            Gdiplus::LinearGradientBrush arrowBrush(arrowRect, arrowStartColor, arrowEndColor,
                                                    Gdiplus::LinearGradientModeHorizontal);
            arrowBrush.SetGammaCorrection(TRUE);
            graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
        } else {
            const BYTE arrowRed = AverageColorChannel(arrowStartColor.GetR(), arrowEndColor.GetR());
            const BYTE arrowGreen = AverageColorChannel(arrowStartColor.GetG(), arrowEndColor.GetG());
            const BYTE arrowBlue = AverageColorChannel(arrowStartColor.GetB(), arrowEndColor.GetB());
            Gdiplus::SolidBrush arrowBrush(Gdiplus::Color(arrowAlpha, arrowRed, arrowGreen, arrowBlue));
            graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
        }
    };

    HTHEME theme = nullptr;
    if (IsAppThemed() && IsThemeActive()) {
        theme = OpenThemeData(hwnd, L"BreadcrumbBar");
        if (!theme) {
            theme = OpenThemeData(hwnd, L"Toolbar");
        }
    }
    struct ThemeCloser {
        HTHEME handle;
        ~ThemeCloser() {
            if (handle) {
                CloseThemeData(handle);
            }
        }
    } themeCloser{theme};

    const COLORREF highlightBackgroundColor =
        theme ? GetThemeSysColor(theme, COLOR_HIGHLIGHT) : GetSysColor(COLOR_HIGHLIGHT);

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
    const int fontBrightness = std::clamp(m_breadcrumbFontBrightness, 0, 100);
    const BYTE textAlphaBase = 255;

    bool buttonIsPressed = false;
    bool buttonIsHot = false;
    bool buttonHasDropdown = false;
    bool buttonUseFontGradient = false;
    BYTE buttonTextAlpha = 0;
    BYTE buttonScaledAlpha = 0;
    bool buttonBackgroundGradientVisible = false;
    Gdiplus::Color buttonBrightFontEnd;
    Gdiplus::Color buttonTextPaintStart;
    Gdiplus::Color buttonTextPaintEnd;
    RECT buttonRect{};

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

    auto fetchBreadcrumbText = [&](const TBBUTTON& button) -> std::wstring {
        const UINT commandId = static_cast<UINT>(button.idCommand);

        LRESULT textLength = SendMessage(hwnd, TB_GETBUTTONTEXTW, commandId, 0);
        if (textLength > 0) {
            std::wstring text(static_cast<size_t>(textLength) + 1, L'\0');
            LRESULT copied =
                SendMessage(hwnd, TB_GETBUTTONTEXTW, commandId, reinterpret_cast<LPARAM>(text.data()));
            if (copied > 0) {
                text.resize(static_cast<size_t>(copied));
                return text;
            }
        }

        // Some breadcrumb configurations clear the toolbar's stored text. In those cases, query the
        // button information directly so we can render the gradient text ourselves.
        constexpr size_t kMaxBreadcrumbText = 512;
        std::wstring fallback(kMaxBreadcrumbText, L'\0');
        TBBUTTONINFOW info{};
        info.cbSize = sizeof(info);
        info.dwMask = TBIF_TEXT;
        info.pszText = fallback.data();
        info.cchText = static_cast<int>(fallback.size());
        if (SendMessage(hwnd, TB_GETBUTTONINFOW, commandId, reinterpret_cast<LPARAM>(&info))) {
            fallback.resize(std::wcslen(fallback.c_str()));
            if (!fallback.empty()) {
                return fallback;
            }
        }

        return std::wstring();
    };

    const int buttonCount = static_cast<int>(SendMessage(hwnd, TB_BUTTONCOUNT, 0, 0));
    const LRESULT hotItemIndex = SendMessage(hwnd, TB_GETHOTITEM, 0, 0);
    int gradientStartX = 0;
    int gradientEndX = 0;
    if (m_useCustomBreadcrumbFontColors) {
        RECT toolbarRect{};
        if (GetClientRect(hwnd, &toolbarRect)) {
            gradientStartX = toolbarRect.left;
            gradientEndX = toolbarRect.right;
        }
        int detectedLeft = std::numeric_limits<int>::max();
        int detectedRight = std::numeric_limits<int>::min();
        for (int i = 0; i < buttonCount; ++i) {
            TBBUTTON gradientButton{};
            if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&gradientButton))) {
                continue;
            }
            if ((gradientButton.fsStyle & TBSTYLE_SEP) != 0 || (gradientButton.fsState & TBSTATE_HIDDEN) != 0) {
                continue;
            }
            RECT gradientRect{};
            if (!SendMessage(hwnd, TB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&gradientRect))) {
                continue;
            }
            detectedLeft = std::min(detectedLeft, static_cast<int>(gradientRect.left));
            detectedRight = std::max(detectedRight, static_cast<int>(gradientRect.right));
        }
        if (detectedLeft < detectedRight) {
            gradientStartX = detectedLeft;
            gradientEndX = detectedRight;
        }
    }

    auto sampleFontGradientAtX = [&](int x) -> COLORREF {
        if (!m_useCustomBreadcrumbFontColors) {
            return m_breadcrumbFontGradientStartColor;
        }
        if (gradientEndX <= gradientStartX) {
            return x <= gradientStartX ? m_breadcrumbFontGradientStartColor : m_breadcrumbFontGradientEndColor;
        }
        const int clamped = std::clamp(x, gradientStartX, gradientEndX);
        const double position = static_cast<double>(clamped - gradientStartX) /
                                static_cast<double>(gradientEndX - gradientStartX);
        auto interpolateChannel = [&](BYTE start, BYTE end) -> BYTE {
            const double value = static_cast<double>(start) +
                                 (static_cast<double>(end) - static_cast<double>(start)) * position;
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(value)), 0, 255));
        };
        return RGB(interpolateChannel(GetRValue(m_breadcrumbFontGradientStartColor),
                                      GetRValue(m_breadcrumbFontGradientEndColor)),
                   interpolateChannel(GetGValue(m_breadcrumbFontGradientStartColor),
                                      GetGValue(m_breadcrumbFontGradientEndColor)),
                   interpolateChannel(GetBValue(m_breadcrumbFontGradientStartColor),
                                      GetBValue(m_breadcrumbFontGradientEndColor)));
    };
    int colorIndex = 0;
    for (int i = 0; i < buttonCount; ++i) {
        TBBUTTON button{};
        if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&button))) {
            continue;
        }
        if ((button.fsStyle & TBSTYLE_SEP) != 0 || (button.fsState & TBSTATE_HIDDEN) != 0) {
            continue;
        }

        if (!SendMessage(hwnd, TB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&buttonRect))) {
            continue;
        }

        buttonIsPressed = (button.fsState & TBSTATE_PRESSED) != 0;
        buttonIsHot = !buttonIsPressed && ((button.fsState & TBSTATE_HOT) != 0 ||
                                           (hotItemIndex >= 0 && i == static_cast<int>(hotItemIndex)));
        buttonHasDropdown = (button.fsStyle & BTNS_DROPDOWN) != 0;
        const bool hasIcon = imageList && imageWidth > 0 && imageHeight > 0 && button.iBitmap >= 0 &&
                              button.iBitmap != I_IMAGENONE;
        buttonUseFontGradient = m_breadcrumbFontGradientEnabled;

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

        auto darkenChannel = [](BYTE channel) -> BYTE {
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(channel) * 35 / 100, 0, 255));
        };
        auto transformBackgroundChannel = [&](BYTE channel) -> BYTE {
            return m_useCustomBreadcrumbGradientColors ? channel : darkenChannel(channel);
        };
        auto applyBrightness = [&](BYTE channel) -> BYTE {
            const int boosted = channel + ((255 - channel) * fontBrightness) / 100;
            return static_cast<BYTE>(std::clamp<int>(boosted, 0, 255));
        };

        Gdiplus::RectF rectF(static_cast<Gdiplus::REAL>(buttonRect.left),
                             static_cast<Gdiplus::REAL>(buttonRect.top),
                             static_cast<Gdiplus::REAL>(buttonRect.right - buttonRect.left),
                             static_cast<Gdiplus::REAL>(buttonRect.bottom - buttonRect.top));

        BYTE baseAlpha = 200;
        if (buttonIsPressed) {
            baseAlpha = 235;
        } else if (buttonIsHot) {
            baseAlpha = 220;
        }

        buttonScaledAlpha = static_cast<BYTE>(std::clamp<int>(baseAlpha * gradientOpacityPercent / 100, 0, 255));
        buttonBackgroundGradientVisible = (m_breadcrumbGradientEnabled && buttonScaledAlpha > 0);

        Gdiplus::Color backgroundGradientStartColor;
        Gdiplus::Color backgroundGradientEndColor;
        bool hasBackgroundGradientColors = false;
        Gdiplus::Color backgroundSolidColor;
        bool hasBackgroundSolidColor = false;

        if (buttonBackgroundGradientVisible) {
            backgroundGradientStartColor = BrightenBreadcrumbColor(
                Gdiplus::Color(buttonScaledAlpha, transformBackgroundChannel(GetRValue(startRgb)),
                               transformBackgroundChannel(GetGValue(startRgb)),
                               transformBackgroundChannel(GetBValue(startRgb))),
                buttonIsHot, buttonIsPressed, highlightBackgroundColor);
            backgroundGradientEndColor = BrightenBreadcrumbColor(
                Gdiplus::Color(buttonScaledAlpha, transformBackgroundChannel(GetRValue(endRgb)),
                               transformBackgroundChannel(GetGValue(endRgb)),
                               transformBackgroundChannel(GetBValue(endRgb))),
                buttonIsHot, buttonIsPressed, highlightBackgroundColor);
            hasBackgroundGradientColors = true;
            Gdiplus::LinearGradientBrush backgroundBrush(rectF, backgroundGradientStartColor,
                                                         backgroundGradientEndColor, Gdiplus::LinearGradientModeHorizontal);
            backgroundBrush.SetGammaCorrection(TRUE);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
            graphics.FillRectangle(&backgroundBrush, rectF);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
        } else {
            RECT sampleRect = buttonRect;
            const COLORREF averageBackground = SampleAverageColor(drawDc, sampleRect);
            backgroundSolidColor = Gdiplus::Color(255, GetRValue(averageBackground), GetGValue(averageBackground),
                                                  GetBValue(averageBackground));
            hasBackgroundSolidColor = true;
            if (buttonIsHot || buttonIsPressed) {
                graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                const BYTE overlayAlpha = static_cast<BYTE>(buttonIsPressed ? 140 : 100);
                Gdiplus::Color overlayColor(overlayAlpha, GetRValue(highlightBackgroundColor),
                                            GetGValue(highlightBackgroundColor), GetBValue(highlightBackgroundColor));
                Gdiplus::SolidBrush overlayBrush(overlayColor);
                graphics.FillRectangle(&overlayBrush, rectF);
                auto blendOverlayChannel = [&](BYTE base, BYTE overlay) -> BYTE {
                    const int blended = static_cast<int>(std::lround(base + (overlay - base) * (overlayAlpha / 255.0)));
                    return static_cast<BYTE>(std::clamp(blended, 0, 255));
                };
                backgroundSolidColor = Gdiplus::Color(
                    255, blendOverlayChannel(backgroundSolidColor.GetR(), overlayColor.GetR()),
                    blendOverlayChannel(backgroundSolidColor.GetG(), overlayColor.GetG()),
                    blendOverlayChannel(backgroundSolidColor.GetB(), overlayColor.GetB()));
            }
        }

        if (hasIcon) {
            const int iconX = buttonRect.left + 4;
            const LONG verticalSpace = ((buttonRect.bottom - buttonRect.top) - imageHeight) / 2;
            const LONG iconYOffset = std::max<LONG>(0, verticalSpace);
            const int iconY = static_cast<int>(buttonRect.top + iconYOffset);
            ImageList_Draw(imageList, button.iBitmap, drawDc, iconX, iconY, ILD_TRANSPARENT);
        }

        buttonTextAlpha = textAlphaBase;
        if (buttonIsPressed) {
            buttonTextAlpha = static_cast<BYTE>(std::min<int>(255, buttonTextAlpha + 60));
        } else if (buttonIsHot) {
            buttonTextAlpha = static_cast<BYTE>(std::min<int>(255, buttonTextAlpha + 35));
        }
        COLORREF buttonFontStartRgb = startRgb;
        COLORREF buttonFontEndRgb = endRgb;
        if (m_useCustomBreadcrumbFontColors) {
            buttonFontStartRgb = sampleFontGradientAtX(buttonRect.left);
            buttonFontEndRgb = sampleFontGradientAtX(buttonRect.right);
        }

        auto computeBrightFontColor = [&](COLORREF rgb) -> Gdiplus::Color {
            const BYTE adjustedRed = applyBrightness(GetRValue(rgb));
            const BYTE adjustedGreen = applyBrightness(GetGValue(rgb));
            const BYTE adjustedBlue = applyBrightness(GetBValue(rgb));
            return BrightenBreadcrumbColor(
                Gdiplus::Color(buttonTextAlpha, adjustedRed, adjustedGreen, adjustedBlue), buttonIsHot,
                buttonIsPressed, highlightBackgroundColor);
        };

        const Gdiplus::Color buttonBrightFontStart = computeBrightFontColor(buttonFontStartRgb);
        buttonBrightFontEnd = computeBrightFontColor(buttonFontEndRgb);

        auto computeOpaqueFontColor = [&](const Gdiplus::Color& fontColor, bool useStart) {
            if (buttonTextAlpha >= 255) {
                return Gdiplus::Color(255, fontColor.GetR(), fontColor.GetG(), fontColor.GetB());
            }
            const double opacity = static_cast<double>(buttonTextAlpha) / 255.0;
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
                backgroundRed = GetRValue(highlightBackgroundColor);
                backgroundGreen = GetGValue(highlightBackgroundColor);
                backgroundBlue = GetBValue(highlightBackgroundColor);
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

        buttonTextPaintStart = computeOpaqueFontColor(buttonBrightFontStart, true);
        buttonTextPaintEnd = computeOpaqueFontColor(buttonBrightFontEnd, false);

        constexpr int kTextPadding = 8;
        const int iconReserve = hasIcon ? (imageWidth + 6) : 0;
        const int dropdownReserve = buttonHasDropdown ? 12 : 0;
        const int availableTextWidth = (buttonRect.right - buttonRect.left) - iconReserve - dropdownReserve -
                                       (kTextPadding * 2);
        const bool iconOnlyButton = hasIcon && availableTextWidth <= 4 && (button.fsStyle & BTNS_SHOWTEXT) == 0;

        std::wstring text;
        if (!iconOnlyButton) {
            text = fetchBreadcrumbText(button);
            if (!text.empty()) {
                const int iconAreaLeft = buttonRect.left + iconReserve;
                const int textBaseLeft = iconAreaLeft + kTextPadding;
                RECT textRect = buttonRect;
                textRect.left = std::max(iconAreaLeft, textBaseLeft - 1);
                textRect.right -= kTextPadding;
                if (buttonHasDropdown) {
                    textRect.right -= dropdownReserve;
                }

                if (textRect.right > textRect.left) {
                    Gdiplus::RectF textRectF(static_cast<Gdiplus::REAL>(textRect.left),
                                             static_cast<Gdiplus::REAL>(textRect.top),
                                             static_cast<Gdiplus::REAL>(textRect.right - textRect.left),
                                             static_cast<Gdiplus::REAL>(textRect.bottom - textRect.top));

                    COLORREF textFontStartRgb = buttonFontStartRgb;
                    COLORREF textFontEndRgb = buttonFontEndRgb;
                    if (m_useCustomBreadcrumbFontColors) {
                        textFontStartRgb = sampleFontGradientAtX(textRect.left);
                        textFontEndRgb = sampleFontGradientAtX(textRect.right);
                    }
                    const Gdiplus::Color brightFontStart = computeBrightFontColor(textFontStartRgb);
                    const Gdiplus::Color textBrightFontEnd = computeBrightFontColor(textFontEndRgb);
                    buttonTextPaintStart = computeOpaqueFontColor(brightFontStart, true);
                    buttonTextPaintEnd = computeOpaqueFontColor(textBrightFontEnd, false);

                    if (buttonTextAlpha > 0) {
                        if (buttonUseFontGradient) {
                            const auto previousHint = graphics.GetTextRenderingHint();
                            const auto previousMode = graphics.GetCompositingMode();
                            const auto previousPixelOffset = graphics.GetPixelOffsetMode();
                            const auto previousSmoothing = graphics.GetSmoothingMode();

                            graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintAntiAliasGridFit);
                            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                            graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
                            graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);

                            Gdiplus::LinearGradientBrush textBrush(
                                textRectF,
                                buttonTextPaintStart,
                                buttonTextPaintEnd,
                                Gdiplus::LinearGradientModeHorizontal);
                            textBrush.SetGammaCorrection(TRUE);

                            bool renderedWithPath = false;
                            Gdiplus::FontFamily fontFamily;
                            if (font.GetFamily(&fontFamily) == Gdiplus::Ok) {
                                Gdiplus::GraphicsPath textPath;
                                if (textPath.AddString(text.c_str(), static_cast<INT>(text.size()), &fontFamily,
                                                       font.GetStyle(), font.GetSize(), textRectF, &format) ==
                                    Gdiplus::Ok) {
                                    graphics.FillPath(&textBrush, &textPath);
                                    renderedWithPath = true;
                                }
                            }

                            if (!renderedWithPath) {
                                graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                    &textBrush);
                            }

                            graphics.SetSmoothingMode(previousSmoothing);
                            graphics.SetPixelOffsetMode(previousPixelOffset);
                            graphics.SetCompositingMode(previousMode);
                            graphics.SetTextRenderingHint(previousHint);
                        } else {
                            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
                            const BYTE avgRed = AverageColorChannel(brightFontStart.GetR(), textBrightFontEnd.GetR());
                            const BYTE avgGreen = AverageColorChannel(brightFontStart.GetG(), textBrightFontEnd.GetG());
                            const BYTE avgBlue = AverageColorChannel(brightFontStart.GetB(), textBrightFontEnd.GetB());
                            Gdiplus::Color solidColor = computeOpaqueFontColor(
                                Gdiplus::Color(buttonTextAlpha, avgRed, avgGreen, avgBlue), true);
                            Gdiplus::SolidBrush textBrush(solidColor);
                            graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                &textBrush);
                        }
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                    }
                }
            }
        }

        Gdiplus::Color arrowTextStart = buttonTextPaintStart;
        Gdiplus::Color arrowTextEnd = buttonTextPaintEnd;
        if (buttonHasDropdown && m_useCustomBreadcrumbFontColors) {
            const int arrowLeft = buttonRect.right - 12;
            const int arrowRight = buttonRect.right - 6;
            const Gdiplus::Color arrowBrightStart = computeBrightFontColor(sampleFontGradientAtX(arrowLeft));
            const Gdiplus::Color arrowBrightEnd = computeBrightFontColor(sampleFontGradientAtX(arrowRight));
            arrowTextStart = computeOpaqueFontColor(arrowBrightStart, true);
            arrowTextEnd = computeOpaqueFontColor(arrowBrightEnd, false);
        }

        if (buttonHasDropdown) {
            drawDropdownArrow(buttonRect, buttonIsHot, buttonIsPressed, buttonTextAlpha, buttonBrightFontEnd,
                              arrowTextStart, arrowTextEnd, buttonUseFontGradient, m_breadcrumbGradientEnabled,
                              buttonBackgroundGradientVisible, buttonScaledAlpha, highlightBackgroundColor);
        }
    }

    if (buffer) {
        EndBufferedPaint(buffer, TRUE);
    }
    EndPaint(hwnd, &ps);
    return true;
}

bool CExplorerBHO::HandleProgressPaint(HWND hwnd) {
    if (!m_useCustomProgressGradientColors) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc) {
        return false;
    }

    RECT client{};
    if (!GetClientRect(hwnd, &client)) {
        EndPaint(hwnd, &ps);
        return true;
    }

    RECT inner = client;
    DrawEdge(dc, &inner, EDGE_SUNKEN, BF_RECT | BF_ADJUST);

    FillRect(dc, &inner, GetSysColorBrush(COLOR_WINDOW));

    PBRANGE range{};
    SendMessageW(hwnd, PBM_GETRANGE, TRUE, reinterpret_cast<LPARAM>(&range));
    if (range.iHigh <= range.iLow) {
        range.iLow = 0;
        range.iHigh = 100;
    }

    LRESULT position = SendMessageW(hwnd, PBM_GETPOS, 0, 0);
    const int span = range.iHigh - range.iLow;
    double fraction = 0.0;
    if (span > 0) {
        fraction = static_cast<double>(position - range.iLow) / static_cast<double>(span);
    }
    fraction = std::clamp(fraction, 0.0, 1.0);

    const LONG width = inner.right - inner.left;
    if (fraction > 0.0 && width > 0) {
        const LONG progressWidth = static_cast<LONG>(std::lround(fraction * static_cast<double>(width)));
        if (progressWidth > 0) {
            RECT fillRect = inner;
            fillRect.right = std::min(fillRect.left + progressWidth, inner.right);
            const LONG fillWidth = fillRect.right - fillRect.left;
            const LONG fillHeight = fillRect.bottom - fillRect.top;
            if (fillWidth > 0 && fillHeight > 0) {
                bool rendered = false;
                if (EnsureProgressGradientResources() && m_progressGradientBits &&
                    m_progressGradientInfo.bmiHeader.biWidth > 0) {
                    const int previousMode = SetStretchBltMode(dc, HALFTONE);
                    POINT origin{};
                    if (previousMode != 0) {
                        SetBrushOrgEx(dc, 0, 0, &origin);
                    }
                    const int srcWidth = m_progressGradientInfo.bmiHeader.biWidth;
                    const int srcHeight = (m_progressGradientInfo.bmiHeader.biHeight < 0)
                                              ? -m_progressGradientInfo.bmiHeader.biHeight
                                              : m_progressGradientInfo.bmiHeader.biHeight;
                    const int result = StretchDIBits(dc, fillRect.left, fillRect.top, fillWidth, fillHeight, 0, 0,
                                                     srcWidth, srcHeight, m_progressGradientBits,
                                                     &m_progressGradientInfo, DIB_RGB_COLORS, SRCCOPY);
                    if (previousMode != 0) {
                        SetBrushOrgEx(dc, origin.x, origin.y, nullptr);
                        SetStretchBltMode(dc, previousMode);
                    }
                    rendered = (result != GDI_ERROR);
                }

                if (!rendered) {
                    TRIVERTEX vertex[2] = {};
                    vertex[0].x = fillRect.left;
                    vertex[0].y = fillRect.top;
                    vertex[0].Red = static_cast<COLOR16>(GetRValue(m_progressGradientStartColor) << 8);
                    vertex[0].Green = static_cast<COLOR16>(GetGValue(m_progressGradientStartColor) << 8);
                    vertex[0].Blue = static_cast<COLOR16>(GetBValue(m_progressGradientStartColor) << 8);
                    vertex[0].Alpha = 0xFFFF;

                    vertex[1].x = fillRect.right;
                    vertex[1].y = fillRect.bottom;
                    vertex[1].Red = static_cast<COLOR16>(GetRValue(m_progressGradientEndColor) << 8);
                    vertex[1].Green = static_cast<COLOR16>(GetGValue(m_progressGradientEndColor) << 8);
                    vertex[1].Blue = static_cast<COLOR16>(GetBValue(m_progressGradientEndColor) << 8);
                    vertex[1].Alpha = 0xFFFF;

                    GRADIENT_RECT gradientRect{0, 1};
                    GradientFill(dc, vertex, 2, &gradientRect, 1, GRADIENT_FILL_RECT_H);
                }
            }
        }
    }

    EndPaint(hwnd, &ps);
    return true;
}

bool CExplorerBHO::HandleAddressEditPaint(HWND hwnd) {
    if (!m_breadcrumbFontGradientEnabled || !m_useCustomBreadcrumbFontColors) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc) {
        return false;
    }

    RECT client{};
    if (!GetClientRect(hwnd, &client)) {
        EndPaint(hwnd, &ps);
        return true;
    }

    const BOOL caretHidden = HideCaret(hwnd);

    COLORREF backgroundColor = GetBkColor(dc);
    if (backgroundColor == CLR_INVALID) {
        backgroundColor = GetSysColor(COLOR_WINDOW);
    }
    const COLORREF originalTextColor = SetTextColor(dc, backgroundColor);
    const int originalBkMode = SetBkMode(dc, OPAQUE);
    SendMessageW(hwnd, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc), PRF_CLIENT);
    SetTextColor(dc, originalTextColor);
    SetBkMode(dc, originalBkMode);

    std::wstring text;
    const int length = GetWindowTextLengthW(hwnd);
    if (length > 0) {
        text.resize(length);
        int copied = GetWindowTextW(hwnd, text.data(), length + 1);
        if (copied < 0) {
            copied = 0;
        }
        text.resize(static_cast<size_t>(std::clamp(copied, 0, length)));
    }

    RECT formatRect = client;
    SendMessageW(hwnd, EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formatRect));
    if (formatRect.right <= formatRect.left) {
        formatRect = client;
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    if (!font) {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    HFONT oldFont = nullptr;
    if (font) {
        oldFont = static_cast<HFONT>(SelectObject(dc, font));
    }

    const COLORREF gradientStart = m_breadcrumbFontGradientStartColor;
    const COLORREF gradientEnd = m_breadcrumbFontGradientEndColor;
    const int brightness = std::clamp(m_breadcrumbFontBrightness, 0, 100);

    const COLORREF previousTextColor = GetTextColor(dc);
    const int previousBkMode = SetBkMode(dc, TRANSPARENT);

    auto applyBrightness = [&](BYTE channel) -> BYTE {
        const int boosted = channel + ((255 - channel) * brightness) / 100;
        return static_cast<BYTE>(std::clamp(boosted, 0, 255));
    };

    auto brightenColor = [&](COLORREF color) -> Gdiplus::Color {
        return Gdiplus::Color(255, applyBrightness(GetRValue(color)), applyBrightness(GetGValue(color)),
                              applyBrightness(GetBValue(color)));
    };

    bool gradientDrawn = false;
    if (!text.empty() && formatRect.right > formatRect.left && formatRect.bottom > formatRect.top) {
        Gdiplus::Graphics graphics(dc);
        if (graphics.GetLastStatus() == Gdiplus::Ok) {
            graphics.SetPageUnit(Gdiplus::UnitPixel);
            graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
            graphics.SetSmoothingMode(Gdiplus::SmoothingModeNone);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);

            Gdiplus::Font gradientFont(dc, font);
            if (gradientFont.GetLastStatus() == Gdiplus::Ok) {
                const Gdiplus::RectF layoutRect(
                    static_cast<Gdiplus::REAL>(formatRect.left), static_cast<Gdiplus::REAL>(formatRect.top),
                    static_cast<Gdiplus::REAL>(std::max<LONG>(1, formatRect.right - formatRect.left)),
                    static_cast<Gdiplus::REAL>(std::max<LONG>(1, formatRect.bottom - formatRect.top)));

                Gdiplus::REAL gradientStartX = static_cast<Gdiplus::REAL>(formatRect.left);
                Gdiplus::REAL gradientEndX = static_cast<Gdiplus::REAL>(formatRect.right);
                if (gradientEndX <= gradientStartX) {
                    gradientEndX = gradientStartX + 1.0f;
                }

                Gdiplus::LinearGradientBrush brush(Gdiplus::PointF(gradientStartX, 0.0f),
                                                   Gdiplus::PointF(gradientEndX, 0.0f), brightenColor(gradientStart),
                                                   brightenColor(gradientEnd));
                if (brush.GetLastStatus() == Gdiplus::Ok) {
                    brush.SetWrapMode(Gdiplus::WrapModeClamp);

                    Gdiplus::StringFormat format(Gdiplus::StringFormat::GenericDefault());
                    format.SetFormatFlags(format.GetFormatFlags() | Gdiplus::StringFormatFlagsNoWrap);
                    format.SetTrimming(Gdiplus::StringTrimmingNone);

                    graphics.SetClip(layoutRect);
                    if (graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &gradientFont, layoutRect, &format,
                                            &brush) == Gdiplus::Ok) {
                        gradientDrawn = true;
                    }
                    graphics.ResetClip();
                }
            }
        }
    }

    if (!gradientDrawn) {
        auto interpolateChannel = [&](BYTE start, BYTE end, double position) -> BYTE {
            const double value = static_cast<double>(start) +
                                 (static_cast<double>(end) - static_cast<double>(start)) * position;
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(value)), 0, 255));
        };

        const double gradientWidth = static_cast<double>(std::max<LONG>(1, formatRect.right - formatRect.left));

        for (int i = 0; i < static_cast<int>(text.size()); ++i) {
            LRESULT pos = SendMessageW(hwnd, EM_POSFROMCHAR, i, 0);
            if (pos == -1) {
                continue;
            }
            const int charX = static_cast<SHORT>(LOWORD(static_cast<DWORD_PTR>(pos)));
            const int charY = static_cast<SHORT>(HIWORD(static_cast<DWORD_PTR>(pos)));

            LRESULT nextPos = SendMessageW(hwnd, EM_POSFROMCHAR, i + 1, 0);
            int nextX = (nextPos == -1) ? charX : static_cast<SHORT>(LOWORD(static_cast<DWORD_PTR>(nextPos)));
            if (nextPos == -1 || nextX <= charX) {
                SIZE extent{};
                if (GetTextExtentPoint32W(dc, &text[i], 1, &extent)) {
                    nextX = charX + extent.cx;
                } else {
                    nextX = charX + 1;
                }
            }

            const int charWidth = std::max(1, nextX - charX);
            double centerX = static_cast<double>(charX) + static_cast<double>(charWidth) / 2.0;
            double position = (centerX - static_cast<double>(formatRect.left)) / gradientWidth;
            position = std::clamp(position, 0.0, 1.0);

            const BYTE red = applyBrightness(
                interpolateChannel(GetRValue(gradientStart), GetRValue(gradientEnd), position));
            const BYTE green = applyBrightness(
                interpolateChannel(GetGValue(gradientStart), GetGValue(gradientEnd), position));
            const BYTE blue = applyBrightness(
                interpolateChannel(GetBValue(gradientStart), GetBValue(gradientEnd), position));

            SetTextColor(dc, RGB(red, green, blue));
            ExtTextOutW(dc, charX, charY, ETO_CLIPPED, &formatRect, &text[i], 1, nullptr);
        }
    }

    SetBkMode(dc, previousBkMode);
    SetTextColor(dc, previousTextColor);

    if (oldFont) {
        SelectObject(dc, oldFont);
    }

    if (caretHidden) {
        ShowCaret(hwnd);
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

        wchar_t classBuffer[64]{};
        if (!className) {
            if (GetClassNameW(hwnd, classBuffer, ARRAYSIZE(classBuffer)) > 0) {
                className = classBuffer;
            }
        }

        if (!className) {
            return CallNextHookEx(nullptr, code, wParam, lParam);
        }

        const bool isToolbar = (_wcsicmp(className, TOOLBARCLASSNAMEW) == 0);
        const bool isCombo = (_wcsicmp(className, L"ComboBoxEx32") == 0);
        const bool isEdit = (_wcsicmp(className, L"Edit") == 0);
        if (!isToolbar && !isCombo && !isEdit) {
            return CallNextHookEx(nullptr, code, wParam, lParam);
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
                if (!observer || !observer->m_gdiplusInitialized) {
                    continue;
                }

                if (isToolbar) {
                    if (!observer->m_breadcrumbGradientEnabled && !observer->m_breadcrumbFontGradientEnabled) {
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
                    continue;
                }

                if (!observer->m_breadcrumbFontGradientEnabled || !observer->m_useCustomBreadcrumbFontColors) {
                    continue;
                }

                HWND ancestryCheck = hwnd;
                if (create && create->lpcs && create->lpcs->hwndParent) {
                    ancestryCheck = create->lpcs->hwndParent;
                }
                if (!observer->IsBreadcrumbToolbarAncestor(ancestryCheck) &&
                    !observer->IsBreadcrumbToolbarAncestor(hwnd)) {
                    continue;
                }
                if (!observer->IsWindowOwnedByThisExplorer(hwnd)) {
                    continue;
                }

                observer->UpdateAddressEditSubclass();
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

LRESULT CALLBACK CExplorerBHO::ProgressSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                    UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_PAINT:
            if (self->HandleProgressPaint(hwnd)) {
                return 0;
            }
            break;
        case WM_ERASEBKGND:
            if (self->m_useCustomProgressGradientColors) {
                return 1;
            }
            break;
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
            if (self->m_useCustomProgressGradientColors) {
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        case WM_NCDESTROY:
            self->RemoveProgressSubclass();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::AddressEditSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                       UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_PAINT:
            if (self->HandleAddressEditPaint(hwnd)) {
                return 0;
            }
            break;
        case WM_SETTEXT:
        case EM_REPLACESEL:
        case EM_SETSEL:
        case WM_CUT:
        case WM_PASTE:
        case WM_UNDO:
        case WM_CLEAR:
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
        case WM_SETFONT:
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_CHAR:
        case WM_KEYDOWN:
        case WM_KEYUP:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_MBUTTONDOWN:
        case WM_MBUTTONUP:
        case WM_MOUSEWHEEL:
        case WM_MOUSEMOVE: {
            LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            if (msg == WM_MOUSEMOVE && (wParam & (MK_LBUTTON | MK_MBUTTON | MK_RBUTTON)) == 0) {
                return result;
            }
            self->RequestAddressEditRedraw(hwnd);
            return result;
        }
        case WM_NCDESTROY:
            self->RemoveAddressEditSubclass();
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

        RemoveWindowSubclass(hwnd, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(self));
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

}  // namespace shelltabs

