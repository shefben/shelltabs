#include "TabBand.h"

#include <algorithm>
#include <atomic>
#include <deque>
#include <memory>
#include <cstring>
#include <cwchar>
#include <mutex>
#include <string>
#include <utility>
#include <vector>
#include <stop_token>

#include "SessionStore.h"
#include "CrashRecoveryLogger.h"
#include <objbase.h>

#include <CommCtrl.h>
#include <ShlObj.h>
#include "TaskbarTabProvider.h"
#include <shellapi.h> // For ShellExecuteW
#include <Shlwapi.h>
#include <shlguid.h>
#include <shobjidl_core.h>

#ifndef SBSP_EXPLORE
#define SBSP_EXPLORE 0x00000004
#endif

#include "Guids.h"
#include "GroupStore.h"
#include "OptionsDialog.h"
#include "OptionsStore.h"
#include "Logging.h"
#include "Module.h"
#include "PreviewCache.h"
#include "ShellTabsMessages.h"
#include "TabBandWindow.h"
#include "Utilities.h"
#include "IconCache.h"
#include "FtpPidl.h"

namespace shelltabs {

using Microsoft::WRL::ComPtr;

namespace {

enum class WindowSeedType {
    StandaloneTab,
    Group,
};

struct PendingWindowSeed {
    WindowSeedType type = WindowSeedType::StandaloneTab;
    TabGroup group;
};

std::mutex& PendingWindowSeedMutex() {
    static auto* mutex = new std::mutex();
    return *mutex;
}

std::deque<std::shared_ptr<PendingWindowSeed>>& PendingWindowSeedQueue() {
    static auto* queue = new std::deque<std::shared_ptr<PendingWindowSeed>>();
    return *queue;
}

void EnqueuePendingWindowSeed(const std::shared_ptr<PendingWindowSeed>& seed) {
    if (!seed) {
        return;
    }
    auto& mutex = PendingWindowSeedMutex();
    auto& queue = PendingWindowSeedQueue();
    std::scoped_lock lock(mutex);
    queue.push_back(seed);
}

std::shared_ptr<PendingWindowSeed> DequeuePendingWindowSeed() {
    auto& mutex = PendingWindowSeedMutex();
    auto& queue = PendingWindowSeedQueue();
    std::scoped_lock lock(mutex);
    if (queue.empty()) {
        return {};
    }
    auto seed = std::move(queue.front());
    queue.pop_front();
    return seed;
}

std::wstring TrimWhitespace(const std::wstring& value) {
    const size_t first = value.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos) {
        return {};
    }
    const size_t last = value.find_last_not_of(L" \t\r\n");
    return value.substr(first, last - first + 1);
}

UniquePidl CreateThisPcPidl() {
    UniquePidl pidl = ParseExplorerUrl(L"shell:MyComputerFolder");
    if (pidl) {
        return pidl;
    }

    LPITEMIDLIST raw = nullptr;
    if (SUCCEEDED(SHGetSpecialFolderLocation(nullptr, CSIDL_DRIVES, &raw)) && raw) {
        return UniquePidl(raw);
    }
    return nullptr;
}

bool EnsureFtpNamespaceBinding(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return false;
    }
    shelltabs::FtpUrlParts parts;
    std::vector<std::wstring> segments;
    bool isDirectory = true;
    if (!shelltabs::ftp::TryParseFtpPidl(pidl, &parts, &segments, &isDirectory)) {
        return false;
    }
    ComPtr<IShellFolder> folder;
    return SUCCEEDED(SHBindToObject(nullptr, pidl, nullptr, IID_PPV_ARGS(&folder)));
}

void LogLoadFailure(const wchar_t* context, const std::wstring& details) {
    if (!details.empty()) {
        LogMessage(LogLevel::Warning, L"%ls: %ls", context, details.c_str());
    } else {
        LogMessage(LogLevel::Warning, L"%ls", context);
    }
}

bool LoadGroupStoreForContext(const wchar_t* context, GroupStore& store) {
    std::wstring errorContext;
    if (!store.Load(&errorContext)) {
        LogLoadFailure(context, errorContext);
        return false;
    }
    return true;
}

bool LoadOptionsStoreForContext(const wchar_t* context, OptionsStore& store) {
    std::wstring errorContext;
    if (!store.Load(&errorContext)) {
        LogLoadFailure(context, errorContext);
        return false;
    }
    return true;
}
}

TabBand::TabBand() : m_refCount(1), m_processedGroupStoreGeneration(0) {
    ModuleAddRef();
    LogMessage(LogLevel::Info, L"TabBand constructed (this=%p)", this);
}

TabBand::~TabBand() {
    LogMessage(LogLevel::Info, L"TabBand destroyed (this=%p)", this);

    // Ensure clean disconnection in destructor
    if (!m_isDestroying.load()) {
        DisconnectSite();
    }

    ModuleRelease();
}

IFACEMETHODIMP TabBand::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }

    if (riid == IID_IUnknown || riid == IID_IDeskBand2) {
        *object = static_cast<IDeskBand2*>(this);
    } else if (riid == IID_IDeskBand) {
        *object = static_cast<IDeskBand*>(this);
    } else if (riid == IID_IDockingWindow) {
        *object = static_cast<IDockingWindow*>(this);
    } else if (riid == IID_IOleWindow) {
        *object = static_cast<IOleWindow*>(this);
    } else if (riid == IID_IInputObject) {
        *object = static_cast<IInputObject*>(this);
    } else if (riid == IID_IObjectWithSite) {
        *object = static_cast<IObjectWithSite*>(this);
    } else if (riid == IID_IPersist) {
        *object = static_cast<IPersist*>(this);
    } else if (riid == IID_IPersistStream) {
        *object = static_cast<IPersistStream*>(this);
    } else {
        *object = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) TabBand::AddRef() {
    return static_cast<ULONG>(++m_refCount);
}

IFACEMETHODIMP_(ULONG) TabBand::Release() {
    const ULONG count = static_cast<ULONG>(--m_refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}

IFACEMETHODIMP TabBand::GetWindow(HWND* phwnd) {
    return GuardExplorerCall(
        L"TabBand::GetWindow",
        [&]() -> HRESULT {
            if (!phwnd) {
                return E_POINTER;
            }
            EnsureWindow();
            *phwnd = m_window ? m_window->GetHwnd() : nullptr;
            return *phwnd ? S_OK : E_FAIL;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::ContextSensitiveHelp(BOOL enterMode) {
    return GuardExplorerCall(
        L"TabBand::ContextSensitiveHelp",
        [&]() -> HRESULT {
            Microsoft::WRL::ComPtr<IOleWindow> site = m_siteOleWindow;
            if (!site && m_dockingSite) {
                m_dockingSite.As(&site);
            }
            if (!site && m_dockingFrame) {
                m_dockingFrame.As(&site);
            }
            if (!site && m_site) {
                m_site.As(&site);
            }

            if (site) {
                const HRESULT hr = site->ContextSensitiveHelp(enterMode);
                if (FAILED(hr) && hr != E_NOTIMPL) {
                    return hr;
                }
            }
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::ShowDW(BOOL fShow) {
    return GuardExplorerCall(
        L"TabBand::ShowDW",
        [&]() -> HRESULT {
            EnsureWindow();
            if (m_window) {
                m_window->Show(fShow != FALSE);
            }
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::CloseDW(DWORD) {
    return GuardExplorerCall(
        L"TabBand::CloseDW",
        [&]() -> HRESULT {
            if (m_window) {
                m_window->Show(false);
            }
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::ResizeBorderDW(const RECT* prcBorder, IUnknown* punkToolbarSite, BOOL fReserved) {
    return GuardExplorerCall(
        L"TabBand::ResizeBorderDW",
        [&]() -> HRESULT {
            Microsoft::WRL::ComPtr<IDockingWindow> dockingWindow;
            if (m_dockingFrame) {
                m_dockingFrame.As(&dockingWindow);
            }
            if (!dockingWindow && punkToolbarSite) {
                punkToolbarSite->QueryInterface(IID_PPV_ARGS(&dockingWindow));
            }
            if (!dockingWindow && m_dockingSite) {
                m_dockingSite.As(&dockingWindow);
            }
            if (!dockingWindow && m_site) {
                m_site.As(&dockingWindow);
            }

            if (dockingWindow) {
                IUnknown* siteForCall = punkToolbarSite;
                if (!siteForCall && m_site) {
                    siteForCall = m_site.Get();
                }
                const HRESULT hr = dockingWindow->ResizeBorderDW(prcBorder, siteForCall, fReserved);
                if (FAILED(hr) && hr != E_NOTIMPL) {
                    return hr;
                }
            }
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi) {
    return GuardExplorerCall(
        L"TabBand::GetBandInfo",
        [&]() -> HRESULT {
            if (!pdbi) {
                return E_POINTER;
            }
            m_bandId = dwBandID;
            m_viewMode = dwViewMode;

            if (pdbi->dwMask & DBIM_MINSIZE) {
                pdbi->ptMinSize.x = 300;
                pdbi->ptMinSize.y = 28;
            }
            if (pdbi->dwMask & DBIM_MAXSIZE) {
                pdbi->ptMaxSize.x = -1;
                pdbi->ptMaxSize.y = 0;
            }
            if (pdbi->dwMask & DBIM_INTEGRAL) {
                pdbi->ptIntegral.x = 0;
                pdbi->ptIntegral.y = 1;
            }
            if (pdbi->dwMask & DBIM_ACTUAL) {
                pdbi->ptActual.x = 0;
                pdbi->ptActual.y = 30;
            }
            if (pdbi->dwMask & DBIM_TITLE) {
                pdbi->wszTitle[0] = L'\0';
            }
            if (pdbi->dwMask & DBIM_MODEFLAGS) {
                pdbi->dwModeFlags = DBIMF_VARIABLEHEIGHT | DBIMF_NORMAL | DBIMF_TOPALIGN;
            }
            if (pdbi->dwMask & DBIM_BKCOLOR) {
                pdbi->dwMask &= ~DBIM_BKCOLOR;
            }

            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::CanRenderComposited(BOOL* pfCanRenderComposited) {
    if (!pfCanRenderComposited) {
        return E_POINTER;
    }
    *pfCanRenderComposited = TRUE;
    return S_OK;
}

IFACEMETHODIMP TabBand::SetCompositionState(BOOL fCompositionEnabled) {
    m_isComposited = fCompositionEnabled != FALSE;
    return S_OK;
}

IFACEMETHODIMP TabBand::GetCompositionState(BOOL* pfCompositionEnabled) {
    if (!pfCompositionEnabled) {
        return E_POINTER;
    }
    *pfCompositionEnabled = m_isComposited ? TRUE : FALSE;
    return S_OK;
}

IFACEMETHODIMP TabBand::UIActivateIO(BOOL fActivate, LPMSG) {
    return GuardExplorerCall(
        L"TabBand::UIActivateIO",
        [&]() -> HRESULT {
            if (fActivate) {
                EnsureWindow();
                if (m_window) {
                    m_window->FocusTab();
                }
                if (m_site) {
                    m_site->OnFocusChangeIS(static_cast<IDockingWindow*>(this), TRUE);
                }
            } else if (m_site) {
                m_site->OnFocusChangeIS(static_cast<IDockingWindow*>(this), FALSE);
            }
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::HasFocusIO() {
    return GuardExplorerCall(
        L"TabBand::HasFocusIO",
        [&]() -> HRESULT {
            if (!m_window) {
                return S_FALSE;
            }
            return m_window->HasFocus() ? S_OK : S_FALSE;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::TranslateAcceleratorIO(LPMSG msg) {
    return GuardExplorerCall(L"TabBand::TranslateAcceleratorIO",
        [&]() -> HRESULT {
            if (!msg) {
                return S_FALSE;
            }

            // Handle Alt+Left and Alt+Right for navigation
            if (msg->message == WM_SYSKEYDOWN || msg->message == WM_KEYDOWN) {
                const bool altPressed = (GetKeyState(VK_MENU) & 0x8000) != 0;

                if (altPressed && msg->wParam == VK_LEFT) {
                    // Alt+Left: Navigate back
                    if (CanNavigateBack()) {
                        OnNavigateBack();
                        return S_OK;
                    }
                } else if (altPressed && msg->wParam == VK_RIGHT) {
                    // Alt+Right: Navigate forward
                    if (CanNavigateForward()) {
                        OnNavigateForward();
                        return S_OK;
                    }
                }
            }

            return S_FALSE;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::SetSite(IUnknown* pUnkSite) {
    return GuardExplorerCall(
        L"TabBand::SetSite",
        [&]() -> HRESULT {
            LogMessage(LogLevel::Info, L"TabBand::SetSite begin (this=%p, site=%p)", this, pUnkSite);
            if (pUnkSite == m_site.Get()) {
                LogMessage(LogLevel::Info, L"TabBand::SetSite site unchanged");
                return S_OK;
            }
            if (!pUnkSite) {
                LogMessage(LogLevel::Info, L"TabBand::SetSite clearing site");
                DisconnectSite();
                return S_OK;
            }

            DisconnectSite();
            // DisconnectSite sets m_isDestroying=true to prevent re-entrancy.
            // Reset it here — we are RE-initializing, not destroying.
            m_isDestroying = false;

            Microsoft::WRL::ComPtr<IInputObjectSite> site;
            HRESULT hr = pUnkSite->QueryInterface(IID_PPV_ARGS(&site));
            if (FAILED(hr)) {
                return hr;
            }
            m_site = site;
            m_siteOleWindow.Reset();
            m_dockingSite.Reset();
            m_dockingFrame.Reset();
            if (site) {
                site.As(&m_siteOleWindow);
                site.As(&m_dockingSite);
                site.As(&m_dockingFrame);
            }
            if (!m_siteOleWindow) {
                pUnkSite->QueryInterface(IID_PPV_ARGS(&m_siteOleWindow));
            }
            if (!m_dockingSite) {
                pUnkSite->QueryInterface(IID_PPV_ARGS(&m_dockingSite));
            }
            if (!m_dockingFrame) {
                pUnkSite->QueryInterface(IID_PPV_ARGS(&m_dockingFrame));
            }

            Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
            hr = pUnkSite->QueryInterface(IID_PPV_ARGS(&serviceProvider));
            if (SUCCEEDED(hr) && serviceProvider) {
                serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_shellBrowser));
                if (!m_shellBrowser) {
                    serviceProvider->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&m_shellBrowser));
                }
                serviceProvider->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&m_webBrowser));
            }

            if ((!m_shellBrowser || !m_webBrowser) && site) {
                serviceProvider.Reset();
                if (SUCCEEDED(site.As(&serviceProvider)) && serviceProvider) {
                    if (!m_shellBrowser) {
                        serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_shellBrowser));
                    }
                    if (!m_shellBrowser) {
                        serviceProvider->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&m_shellBrowser));
                    }
                    if (!m_webBrowser) {
                        serviceProvider->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&m_webBrowser));
                    }
                }
            }

            if (!m_shellBrowser || !m_webBrowser) {
                LogMessage(LogLevel::Warning, L"TabBand::SetSite missing browser interfaces");
                DisconnectSite();
                return E_FAIL;
            }

            LogMessage(LogLevel::Info, L"TabBand::SetSite resolved browser interfaces");
            LogMessage(LogLevel::Info, L"TabBand::SetSite EnsureWindow");
            EnsureWindow();
            if (!m_window) {
                LogMessage(LogLevel::Error, L"TabBand::SetSite failed to create window");
                DisconnectSite();
                return E_FAIL;
            }

            m_tabs.SetWindowId(BuildWindowId());

            LogMessage(LogLevel::Info, L"TabBand::SetSite EnsureSessionStore");
            EnsureSessionStore();

            HRESULT setSiteHr = m_window->SetSite(pUnkSite);
            if (FAILED(setSiteHr)) {
                LogMessage(LogLevel::Warning,
                           L"TabBand::SetSite TabBandWindow::SetSite failed (hr=0x%08X)",
                           static_cast<unsigned int>(setSiteHr));
            }

            m_browserEvents = std::make_unique<BrowserEvents>(this);
            if (m_browserEvents) {
                hr = m_browserEvents->Connect(m_webBrowser);
                if (FAILED(hr)) {
                    LogMessage(LogLevel::Warning, L"TabBand::SetSite BrowserEvents::Connect failed (hr=0x%08X)",
                               static_cast<unsigned int>(hr));
                    m_browserEvents.reset();
                } else {
                    LogMessage(LogLevel::Info, L"TabBand::SetSite BrowserEvents connected");
                }
            }

            LogMessage(LogLevel::Info, L"TabBand::SetSite InitializeTabs");
            InitializeTabs();
            LogMessage(LogLevel::Info, L"TabBand::SetSite UpdateTabsUI (initial)");
            UpdateTabsUI();

            // Initialize options
            EnsureOptionsLoaded();

            // If reuse-existing-window is enabled and another ShellTabs window
            // already exists, mark this window for redirect on first navigation.
            if (m_options.reuseExistingWindow && TabManager::ActiveWindowCount() > 1) {
                m_pendingWindowRedirect = true;
            }

            LogMessage(LogLevel::Info, L"TabBand::SetSite completed successfully");

            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::GetSite(REFIID riid, void** ppvSite) {
    return GuardExplorerCall(
        L"TabBand::GetSite",
        [&]() -> HRESULT {
            if (!ppvSite) {
                return E_POINTER;
            }
            if (!m_site) {
                *ppvSite = nullptr;
                return E_FAIL;
            }
            return m_site->QueryInterface(riid, ppvSite);
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP TabBand::GetClassID(CLSID* pClassID) {
    if (!pClassID) {
        return E_POINTER;
    }
    *pClassID = CLSID_ShellTabsBand;
    return S_OK;
}

IFACEMETHODIMP TabBand::IsDirty() {
    return S_FALSE;
}

IFACEMETHODIMP TabBand::Load(IStream*) {
    return S_OK;
}

IFACEMETHODIMP TabBand::Save(IStream*, BOOL) {
    return S_OK;
}

IFACEMETHODIMP TabBand::GetSizeMax(ULARGE_INTEGER* pcbSize) {
    if (pcbSize) {
        pcbSize->QuadPart = 0;
    }
    return S_OK;
}

void TabBand::OnBrowserNavigate() {
    EnsureTabForCurrentFolder();
    UpdateTabsUI();
    CaptureActiveTabPreview();

    const bool wasInternal = m_internalNavigation;

    // Restore scroll position for the tab we just navigated to
    // This is only relevant for internal navigations (tab switches)
    if (wasInternal) {
        const TabLocation selected = m_tabs.SelectedLocation();
        if (!RestoreCurrentTabScrollPosition()) {
            // The shell view's items may not be enumerated yet — retry on a timer.
            ScheduleScrollRestoreRetries(selected);
        }
    }

    m_internalNavigation = false;

    // For tab switches the tab list itself didn't change — let the periodic
    // session-flush timer pick up the selection update so that switching
    // doesn't pay a synchronous WRITE_THROUGH disk I/O cost.
    //
    // For external navigations (user typed in address bar, hit Back/Forward,
    // followed a link in a folder, etc.) the tab's PIDL may have changed and
    // we want that on disk now.
    if (!wasInternal) {
        SaveSession();
    }
}

void TabBand::OnBrowserQuit() {
    DisconnectSite();
}

bool TabBand::OnBrowserNewWindow(const std::wstring& targetUrl) {
    return HandleNewWindowRequest(targetUrl);
}

bool TabBand::OnCtrlBeforeNavigate(const std::wstring& url) {
    if (m_internalNavigation || url.empty()) {
        return false;
    }

    UniquePidl pidl = ParseExplorerUrl(url);
    if (!pidl) {
        return false;
    }

    std::wstring name = GetDisplayName(pidl.get());
    if (name.empty()) {
        name = L"Tab";
    }
    std::wstring tooltip = GetParsingName(pidl.get());
    if (tooltip.empty()) {
        tooltip = name;
    }

    const TabLocation current = m_tabs.SelectedLocation();
    const int groupIndex = current.groupIndex >= 0 ? current.groupIndex : -1;
    TabLocation location = m_tabs.Add(std::move(pidl), name, tooltip, true, groupIndex);
    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist new tab
    if (location.IsValid()) {
        QueueNavigateTo(location);
    }
    return true;
}

void TabBand::OnTabSelected(TabLocation location) {
    LogMessage(LogLevel::Info, L"OnTabSelected location=(group=%d,tab=%d)",
               location.groupIndex, location.tabIndex);
    const auto current = m_tabs.SelectedLocation();
    if (current.groupIndex == location.groupIndex && current.tabIndex == location.tabIndex) {
        LogMessage(LogLevel::Info, L"OnTabSelected: already selected, ignoring");
        return;
    }
    LogMessage(LogLevel::Info, L"OnTabSelected: navigating from (group=%d,tab=%d) to (group=%d,tab=%d)",
               current.groupIndex, current.tabIndex, location.groupIndex, location.tabIndex);
    NavigateToTab(location);
}

void TabBand::OnNewTabRequested(int targetGroup) {
    EnsureOptionsLoaded();

    const TabLocation current = m_tabs.SelectedLocation();
    if (targetGroup < 0) {
        targetGroup = current.groupIndex >= 0 ? current.groupIndex : 0;
    }

    auto finalize = [&](UniquePidl pidl, std::wstring name, std::wstring tooltip) {
        if (!pidl) {
            return;
        }
        if (name.empty()) {
            name = GetDisplayName(pidl.get());
        }
        if (name.empty()) {
            name = L"Tab";
        }
        if (tooltip.empty()) {
            tooltip = GetParsingName(pidl.get());
        }
        if (tooltip.empty()) {
            tooltip = name;
        }

        TabLocation location = m_tabs.Add(std::move(pidl), name, tooltip, true, targetGroup);
        UpdateTabsUI();
        SyncAllSavedGroups();
        SaveSession();  // Immediately persist new tab
        if (location.IsValid()) {
            NavigateToTab(location);
        }
    };

    switch (m_options.newTabTemplate) {
        case NewTabTemplate::kDuplicateCurrent: {
            UniquePidl pidl;
            std::wstring name;
            std::wstring tooltip;
            if (const auto* tab = m_tabs.Get(current)) {
                if (tab->pidl) {
                    pidl = ClonePidl(tab->pidl.get());
                }
                name = tab->name;
                tooltip = tab->tooltip;
            }
            if (!pidl) {
                pidl = QueryCurrentFolder();
                if (pidl) {
                    name = GetDisplayName(pidl.get());
                    tooltip = GetParsingName(pidl.get());
                }
            }
            finalize(std::move(pidl), std::move(name), std::move(tooltip));
            return;
        }
        case NewTabTemplate::kThisPc: {
            UniquePidl pidl = CreateThisPcPidl();
            if (!pidl) {
                return;
            }
            std::wstring name = GetDisplayName(pidl.get());
            if (name.empty()) {
                name = L"This PC";
            }
            std::wstring tooltip = GetParsingName(pidl.get());
            finalize(std::move(pidl), std::move(name), std::move(tooltip));
            return;
        }
        case NewTabTemplate::kCustomPath: {
            const std::wstring rawPath = TrimWhitespace(m_options.newTabCustomPath);
            if (rawPath.empty()) {
                return;
            }
            UniquePidl pidl = ParseExplorerUrl(rawPath);
            if (!pidl) {
                pidl = ParseDisplayName(rawPath);
            }
            if (!pidl) {
                return;
            }
            std::wstring name = GetDisplayName(pidl.get());
            if (name.empty()) {
                name = rawPath;
            }
            std::wstring tooltip = GetParsingName(pidl.get());
            finalize(std::move(pidl), std::move(name), std::move(tooltip));
            return;
        }
        case NewTabTemplate::kSavedGroup: {
            const std::wstring target = TrimWhitespace(m_options.newTabSavedGroup);
            if (target.empty()) {
                return;
            }
            auto& store = GroupStore::Instance();
            if (!LoadGroupStoreForContext(L"TabBand::OnNewTabRequested failed to load saved groups", store)) {
                return;
            }
            const SavedGroup* saved = store.Find(target);
            if (!saved) {
                return;
            }
            UniquePidl pidl;
            std::wstring name;
            std::wstring tooltip;
            for (const auto& path : saved->tabPaths) {
                const std::wstring trimmedPath = TrimWhitespace(path);
                if (trimmedPath.empty()) {
                    continue;
                }
                UniquePidl candidate = ParseDisplayName(trimmedPath);
                if (!candidate) {
                    candidate = ParseExplorerUrl(trimmedPath);
                }
                if (!candidate) {
                    continue;
                }
                tooltip = GetParsingName(candidate.get());
                name = GetDisplayName(candidate.get());
                if (name.empty()) {
                    name = trimmedPath;
                }
                pidl = std::move(candidate);
                break;
            }
            finalize(std::move(pidl), std::move(name), std::move(tooltip));
            return;
        }
        default:
            break;
    }
}

void TabBand::OnCloseTabRequested(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool wasSelected = (selected.groupIndex == location.groupIndex && selected.tabIndex == location.tabIndex);

    const TabGroup* groupBefore = m_tabs.GetGroup(location.groupIndex);
    if (!groupBefore) {
        return;
    }

    ClosedTabSet closedSet;
    closedSet.groupIndex = location.groupIndex;
    closedSet.selectionOriginalIndex = location.tabIndex;
    closedSet.groupInfo = CaptureGroupMetadata(*groupBefore);
    if (groupBefore->tabs.size() == 1) {
        closedSet.groupRemoved = true;
    }

    std::wstring removedGroupId;
    if (!groupBefore->savedGroupId.empty() && groupBefore->tabs.size() == 1) {
        removedGroupId = groupBefore->savedGroupId;
    }

    auto removed = m_tabs.TakeTab(location);
    if (!removed) {
        return;
    }
    CancelPendingPreviewForTab(*removed);
    EnsureTabPath(*removed);
    closedSet.entries.push_back({location.tabIndex, std::move(*removed)});
    PushClosedSet(std::move(closedSet));

    if (m_tabs.TotalTabCount() == 0) {
        SaveSession();
        CloseFrameWindowAsync();
        return;
    }

    UpdateTabsUI();
    SyncAllSavedGroups();
    if (!removedGroupId.empty()) {
        GroupStore::Instance().UpdateTabs(removedGroupId, {});
    }

    if (wasSelected) {
        const TabLocation newSelection = m_tabs.SelectedLocation();
        if (newSelection.IsValid()) {
            NavigateToTab(newSelection);
        }
    }

    SaveSession();
}

void TabBand::OnCloseOtherTabsRequested(TabLocation location) {
    if (!location.IsValid() || !CanCloseOtherTabs(location)) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool targetWasSelected = (selected.groupIndex == location.groupIndex &&
                                    selected.tabIndex == location.tabIndex);

    const TabGroup* groupBefore = m_tabs.GetGroup(location.groupIndex);
    if (!groupBefore) {
        return;
    }

    const TabInfo* targetTab = m_tabs.Get(location);
    if (!targetTab) {
        return;
    }
    PCIDLIST_ABSOLUTE anchorPid = targetTab->pidl.get();

    ClosedTabSet closedSet;
    closedSet.groupIndex = location.groupIndex;
    closedSet.groupInfo = CaptureGroupMetadata(*groupBefore);

    for (int index = static_cast<int>(groupBefore->tabs.size()) - 1; index >= 0; --index) {
        if (index == location.tabIndex) {
            continue;
        }
        auto removed = m_tabs.TakeTab({location.groupIndex, index});
        if (!removed) {
            continue;
        }
        CancelPendingPreviewForTab(*removed);
        EnsureTabPath(*removed);
        closedSet.entries.push_back({index, std::move(*removed)});
    }

    if (closedSet.entries.empty()) {
        return;
    }
    closedSet.selectionOriginalIndex = closedSet.entries.back().originalIndex;
    PushClosedSet(std::move(closedSet));

    TabLocation anchorLocation;
    if (anchorPid) {
        anchorLocation = m_tabs.Find(anchorPid);
        if (anchorLocation.IsValid() && targetWasSelected) {
            m_tabs.SetSelectedLocation(anchorLocation);
        }
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (targetWasSelected && anchorLocation.IsValid()) {
        NavigateToTab(anchorLocation);
    }

    SaveSession();
}

void TabBand::OnCloseTabsToRightRequested(TabLocation location) {
    if (!location.IsValid() || !CanCloseTabsToRight(location)) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool targetWasSelected = (selected.groupIndex == location.groupIndex &&
                                    selected.tabIndex == location.tabIndex);

    const TabGroup* groupBefore = m_tabs.GetGroup(location.groupIndex);
    if (!groupBefore) {
        return;
    }
    const TabInfo* targetTab = m_tabs.Get(location);
    if (!targetTab) {
        return;
    }
    PCIDLIST_ABSOLUTE anchorPid = targetTab->pidl.get();

    ClosedTabSet closedSet;
    closedSet.groupIndex = location.groupIndex;
    closedSet.groupInfo = CaptureGroupMetadata(*groupBefore);

    for (int index = static_cast<int>(groupBefore->tabs.size()) - 1; index > location.tabIndex; --index) {
        auto removed = m_tabs.TakeTab({location.groupIndex, index});
        if (!removed) {
            continue;
        }
        CancelPendingPreviewForTab(*removed);
        EnsureTabPath(*removed);
        closedSet.entries.push_back({index, std::move(*removed)});
    }

    if (closedSet.entries.empty()) {
        return;
    }
    closedSet.selectionOriginalIndex = closedSet.entries.back().originalIndex;
    PushClosedSet(std::move(closedSet));

    TabLocation anchorLocation;
    if (anchorPid) {
        anchorLocation = m_tabs.Find(anchorPid);
        if (anchorLocation.IsValid() && targetWasSelected) {
            m_tabs.SetSelectedLocation(anchorLocation);
        }
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (targetWasSelected && anchorLocation.IsValid()) {
        NavigateToTab(anchorLocation);
    }

    SaveSession();
}

void TabBand::OnCloseTabsToLeftRequested(TabLocation location) {
    if (!location.IsValid() || !CanCloseTabsToLeft(location)) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool targetWasSelected = (selected.groupIndex == location.groupIndex &&
                                    selected.tabIndex == location.tabIndex);

    const TabGroup* groupBefore = m_tabs.GetGroup(location.groupIndex);
    if (!groupBefore) {
        return;
    }
    const TabInfo* targetTab = m_tabs.Get(location);
    if (!targetTab) {
        return;
    }
    PCIDLIST_ABSOLUTE anchorPid = targetTab->pidl.get();

    ClosedTabSet closedSet;
    closedSet.groupIndex = location.groupIndex;
    closedSet.groupInfo = CaptureGroupMetadata(*groupBefore);

    for (int index = location.tabIndex - 1; index >= 0; --index) {
        auto removed = m_tabs.TakeTab({location.groupIndex, index});
        if (!removed) {
            continue;
        }
        CancelPendingPreviewForTab(*removed);
        EnsureTabPath(*removed);
        closedSet.entries.push_back({index, std::move(*removed)});
    }

    if (closedSet.entries.empty()) {
        return;
    }
    closedSet.selectionOriginalIndex = closedSet.entries.back().originalIndex;
    PushClosedSet(std::move(closedSet));

    TabLocation anchorLocation;
    if (anchorPid) {
        anchorLocation = m_tabs.Find(anchorPid);
        if (anchorLocation.IsValid() && targetWasSelected) {
            m_tabs.SetSelectedLocation(anchorLocation);
        }
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (targetWasSelected && anchorLocation.IsValid()) {
        NavigateToTab(anchorLocation);
    }

    SaveSession();
}

void TabBand::OnReopenClosedTabRequested(size_t index) {
    if (index >= m_closedTabHistory.size()) {
        return;
    }

    auto it = m_closedTabHistory.rbegin() + index;
    ClosedTabSet set = std::move(*it);
    m_closedTabHistory.erase(it.base() - 1);

    if (set.entries.empty()) {
        return;
    }

    int targetGroupIndex = set.groupIndex;
    if (targetGroupIndex < 0) {
        targetGroupIndex = 0;
    }

    std::vector<ClosedTabEntry> entries = std::move(set.entries);
    std::sort(entries.begin(), entries.end(), [](const ClosedTabEntry& a, const ClosedTabEntry& b) {
        return a.originalIndex < b.originalIndex;
    });

    TabLocation selected{};
    bool haveSelection = false;

    bool createGroup = set.groupRemoved;
    if (!createGroup) {
        if (!m_tabs.GetGroup(targetGroupIndex)) {
            createGroup = true;
        }
    }

    if (createGroup) {
        TabGroup group;
        if (set.groupInfo) {
            group.name = set.groupInfo->name;
            group.collapsed = set.groupInfo->collapsed;
            group.headerVisible = set.groupInfo->headerVisible;
            group.hasCustomOutline = set.groupInfo->hasOutline;
            group.outlineColor = set.groupInfo->outlineColor;
            group.outlineStyle = set.groupInfo->outlineStyle;
            group.savedGroupId = set.groupInfo->savedGroupId;
        }
        for (auto& entry : entries) {
            group.tabs.emplace_back(std::move(entry.tab));
        }
        const int insertedIndex = m_tabs.InsertGroup(std::move(group), targetGroupIndex);
        targetGroupIndex = insertedIndex;
        if (!entries.empty()) {
            int selectionOriginal = set.selectionOriginalIndex;
            int selectIndex = static_cast<int>(entries.size()) - 1;
            if (selectionOriginal >= 0) {
                for (size_t i = 0; i < entries.size(); ++i) {
                    if (entries[i].originalIndex == selectionOriginal) {
                        selectIndex = static_cast<int>(i);
                        break;
                    }
                }
            }
            selected = {targetGroupIndex, selectIndex};
            haveSelection = true;
        }
    } else {
        std::vector<std::pair<int, TabLocation>> insertedLocations;
        insertedLocations.reserve(entries.size());
        int insertedCount = 0;
        for (auto& entry : entries) {
            TabLocation loc = m_tabs.InsertTab(std::move(entry.tab), targetGroupIndex,
                                               entry.originalIndex + insertedCount, false);
            insertedLocations.emplace_back(entry.originalIndex, loc);
            ++insertedCount;
        }

        int selectionOriginal = set.selectionOriginalIndex;
        if (selectionOriginal >= 0) {
            for (const auto& pair : insertedLocations) {
                if (pair.first == selectionOriginal) {
                    selected = pair.second;
                    haveSelection = true;
                    break;
                }
            }
        }
        if (!haveSelection && !insertedLocations.empty()) {
            selected = insertedLocations.back().second;
            haveSelection = true;
        }
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (haveSelection && selected.IsValid()) {
        m_tabs.SetSelectedLocation(selected);
        NavigateToTab(selected);
    }

    SaveSession();
}

void TabBand::OnReopenClosedTabRequested() {
    OnReopenClosedTabRequested(0);
}

bool TabBand::CanCloseOtherTabs(TabLocation location) const {
    const TabGroup* group = m_tabs.GetGroup(location.groupIndex);
    if (!group) {
        return false;
    }
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group->tabs.size())) {
        return false;
    }
    return group->tabs.size() > 1;
}

bool TabBand::CanCloseTabsToRight(TabLocation location) const {
    const TabGroup* group = m_tabs.GetGroup(location.groupIndex);
    if (!group) {
        return false;
    }
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group->tabs.size())) {
        return false;
    }
    return location.tabIndex < static_cast<int>(group->tabs.size()) - 1;
}

bool TabBand::CanCloseTabsToLeft(TabLocation location) const {
    const TabGroup* group = m_tabs.GetGroup(location.groupIndex);
    if (!group) {
        return false;
    }
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group->tabs.size())) {
        return false;
    }
    return location.tabIndex > 0;
}

bool TabBand::CanReopenClosedTabs() const {
    return !m_closedTabHistory.empty();
}

std::wstring TabBand::GetReopenClosedLabel() const {
    if (m_closedTabHistory.empty()) {
        return L"Reopen Last Closed Tab or Island";
    }
    const auto& last = m_closedTabHistory.back();
    if (last.groupRemoved && last.entries.size() > 1) {
        return L"Reopen Last Closed Island (" + std::to_wstring(last.entries.size()) + L" tabs)";
    }
    return L"Reopen Last Closed Tab";
}

bool TabBand::CanNavigateBack() const {
    const TabLocation selected = m_tabs.SelectedLocation();
    return m_tabs.CanNavigateBack(selected);
}

bool TabBand::CanNavigateForward() const {
    const TabLocation selected = m_tabs.SelectedLocation();
    return m_tabs.CanNavigateForward(selected);
}

void TabBand::OnHideTabRequested(TabLocation location) {
    m_tabs.HideTab(location);
    UpdateTabsUI();
}

void TabBand::OnUnhideTabRequested(TabLocation location) {
    m_tabs.UnhideTab(location);
    UpdateTabsUI();
    if (location.IsValid()) {
        NavigateToTab(location);
    }
}

void TabBand::OnDetachTabRequested(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }
    const auto* tab = m_tabs.Get(location);
    if (!tab) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool wasSelected = (selected.groupIndex == location.groupIndex && selected.tabIndex == location.tabIndex);

    std::wstring removedGroupId;
    if (const auto* group = m_tabs.GetGroup(location.groupIndex)) {
        if (!group->savedGroupId.empty() && group->tabs.size() == 1) {
            removedGroupId = group->savedGroupId;
        }
    }

    auto seed = std::make_shared<PendingWindowSeed>();
    seed->type = WindowSeedType::StandaloneTab;
    EnqueuePendingWindowSeed(seed);
    OpenTabInNewWindow(*tab);
    CancelPendingPreviewForTab(*tab);
    m_tabs.Remove(location);
    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }
    UpdateTabsUI();
    SyncAllSavedGroups();
    if (!removedGroupId.empty()) {
        GroupStore::Instance().UpdateTabs(removedGroupId, {});
    }

    if (wasSelected) {
        const TabLocation newSelection = m_tabs.SelectedLocation();
        if (newSelection.IsValid()) {
            NavigateToTab(newSelection);
        }
    }

    SaveSession();  // Immediately persist tab removal
}

void TabBand::OnCloneTabRequested(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }

    const auto* tab = m_tabs.Get(location);
    if (!tab || !tab->pidl) {
        return;
    }

    UniquePidl clone = ClonePidl(tab->pidl.get());
    if (!clone) {
        return;
    }

    std::wstring name = tab->name;
    if (name.empty()) {
        name = GetDisplayName(tab->pidl.get());
    }
    if (name.empty()) {
        name = L"Tab";
    }
    std::wstring tooltip = tab->tooltip.empty() ? name : tab->tooltip;

    TabLocation newLocation = m_tabs.Add(std::move(clone), name, tooltip, true, location.groupIndex, tab->pinned);
    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist cloned tab
    if (newLocation.IsValid()) {
        NavigateToTab(newLocation);
    }
}

void TabBand::OnToggleTabPinned(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }
    const auto* tab = m_tabs.Get(location);
    if (!tab) {
        return;
    }
    if (!m_tabs.ToggleTabPinned(location)) {
        return;
    }
    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist pin state change
}

void TabBand::OnNavigateBack() {
    if (m_internalNavigation) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    if (!selected.IsValid()) {
        return;
    }

    if (!m_shellBrowser) {
        return;
    }

    auto entry = m_tabs.NavigateBack(selected);
    if (!entry) {
        return;
    }

    if (!entry->pidl) {
        m_tabs.NavigateForward(selected);
        return;
    }

    // Set flag to prevent recording this navigation in history
    m_internalNavigation = true;

    // Navigate to the history entry
    const HRESULT hr = m_shellBrowser->BrowseObject(entry->pidl.get(), SBSP_SAMEBROWSER | SBSP_ABSOLUTE | SBSP_WRITENOHISTORY);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TabBand::OnNavigateBack failed to navigate (hr=0x%08X)", hr);
        m_tabs.NavigateForward(selected);
        m_internalNavigation = false;
    }
}

void TabBand::OnNavigateForward() {
    if (m_internalNavigation) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    if (!selected.IsValid()) {
        return;
    }

    if (!m_shellBrowser) {
        return;
    }

    auto entry = m_tabs.NavigateForward(selected);
    if (!entry) {
        return;
    }

    if (!entry->pidl) {
        m_tabs.NavigateBack(selected);
        return;
    }

    // Set flag to prevent recording this navigation in history
    m_internalNavigation = true;

    // Navigate to the history entry
    const HRESULT hr = m_shellBrowser->BrowseObject(entry->pidl.get(), SBSP_SAMEBROWSER | SBSP_ABSOLUTE | SBSP_WRITENOHISTORY);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TabBand::OnNavigateForward failed to navigate (hr=0x%08X)", hr);
        m_tabs.NavigateBack(selected);
        m_internalNavigation = false;
    }
}

bool TabBand::OnShowHistoryMenu(const HistoryMenuRequest& request) {
    const TabLocation selected = m_tabs.SelectedLocation();
    if (!selected.IsValid()) {
        return false;
    }

    TabInfo* tab = m_tabs.Get(selected);
    if (!tab) {
        return false;
    }

    const NavigationHistory& history = tab->navigationHistory;
    const bool hasHistoryEntries = !history.entries.empty();
    const bool hasValidIndex = hasHistoryEntries && history.currentIndex >= 0 &&
                               history.currentIndex < static_cast<int>(history.entries.size());

    HMENU menu = CreatePopupMenu();
    if (!menu) {
        return false;
    }

    std::vector<std::pair<UINT, int>> commandToIndex;
    commandToIndex.reserve(history.entries.size());
    std::vector<std::wstring> labels;
    labels.reserve(history.entries.size());

    if (hasValidIndex) {
        const auto appendEntry = [&](int historyIndex) {
            if (historyIndex < 0 || historyIndex >= static_cast<int>(history.entries.size())) {
                return;
            }
            const auto& entry = history.entries[historyIndex];
            std::wstring label = entry.name.empty() ? entry.path : entry.name;
            if (label.empty()) {
                label = L"(Unknown)";
            }
            labels.push_back(label);

            MENUITEMINFOW item{};
            item.cbSize = sizeof(item);
            item.fMask = MIIM_ID | MIIM_STRING;
            item.wID = static_cast<UINT>(commandToIndex.size() + 1);
            item.dwTypeData = labels.back().data();
            item.cch = static_cast<UINT>(labels.back().size());
            InsertMenuItemW(menu, static_cast<UINT>(-1), TRUE, &item);
            commandToIndex.emplace_back(item.wID, historyIndex);
        };

        if (request.kind == HistoryMenuKind::kBack) {
            for (int index = history.currentIndex - 1; index >= 0; --index) {
                appendEntry(index);
            }
        } else {
            for (int index = history.currentIndex + 1; index < static_cast<int>(history.entries.size()); ++index) {
                appendEntry(index);
            }
        }
    }

    const bool menuHasHistory = !commandToIndex.empty();
    if (!menuHasHistory) {
        labels.emplace_back(L"(No history for this tab)");

        MENUITEMINFOW placeholder{};
        placeholder.cbSize = sizeof(placeholder);
        placeholder.fMask = MIIM_STRING | MIIM_STATE;
        placeholder.fState = MFS_DISABLED;
        placeholder.dwTypeData = labels.back().data();
        placeholder.cch = static_cast<UINT>(labels.back().size());
        InsertMenuItemW(menu, 0, TRUE, &placeholder);
    } else {
        SetMenuDefaultItem(menu, commandToIndex.front().first, FALSE);
    }

    POINT popupPoint{};
    popupPoint.y = request.buttonRect.bottom;

    UINT flags = TPM_TOPALIGN | TPM_LEFTBUTTON | TPM_RETURNCMD;
    if (request.kind == HistoryMenuKind::kBack) {
        popupPoint.x = request.buttonRect.left;
        flags |= TPM_LEFTALIGN;
    } else {
        popupPoint.x = request.buttonRect.right;
        flags |= TPM_RIGHTALIGN;
    }

    HWND ownerHwnd = m_window ? m_window->GetHwnd() : nullptr;
    if (!ownerHwnd) {
        DestroyMenu(menu);
        return false;
    }

    UINT selectedCommand = TrackPopupMenuEx(menu, flags, popupPoint.x, popupPoint.y, ownerHwnd, nullptr);
    DestroyMenu(menu);

    if (!menuHasHistory || selectedCommand == 0) {
        return true;  // Menu displayed but no selection.
    }

    const auto it = std::find_if(commandToIndex.begin(), commandToIndex.end(),
                                 [&](const auto& entry) { return entry.first == selectedCommand; });
    if (it == commandToIndex.end()) {
        return true;
    }

    auto entry = m_tabs.NavigateToHistory(selected, it->second);
    if (!entry || !entry->pidl || !m_shellBrowser) {
        return true;
    }

    m_internalNavigation = true;
    const HRESULT hr = m_shellBrowser->BrowseObject(entry->pidl.get(), SBSP_SAMEBROWSER | SBSP_ABSOLUTE | SBSP_WRITENOHISTORY);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TabBand::OnShowHistoryMenu failed to navigate (hr=0x%08X)", hr);
        m_internalNavigation = false;
    }

    return true;
}

void TabBand::OnToggleGroupCollapsed(int groupIndex) {
    m_tabs.ToggleGroupCollapsed(groupIndex);
    UpdateTabsUI();
    SaveSession();  // Immediately persist collapse state
}

void TabBand::OnUnhideAllInGroup(int groupIndex) {
    m_tabs.UnhideAllInGroup(groupIndex);
    UpdateTabsUI();
    SaveSession();  // Immediately persist unhide operation
}

void TabBand::OnCreateIslandAfter(int groupIndex) {
    m_tabs.CreateGroupAfter(groupIndex);
    UpdateTabsUI();
    SaveSession();  // Immediately persist new island
    SyncAllSavedGroups();
}

void TabBand::OnCloseIslandRequested(int groupIndex) {
    if (groupIndex < 0) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool removedSelectedGroup = (selected.groupIndex == groupIndex);

    auto removed = m_tabs.TakeGroup(groupIndex);
    if (!removed) {
        return;
    }
    CancelPendingPreviewForGroup(*removed);

    ClosedTabSet closedSet;
    closedSet.groupIndex = groupIndex;
    closedSet.groupRemoved = true;
    closedSet.groupInfo = CaptureGroupMetadata(*removed);
    for (size_t i = 0; i < removed->tabs.size(); ++i) {
        EnsureTabPath(removed->tabs[i]);
        closedSet.entries.push_back({static_cast<int>(i), std::move(removed->tabs[i])});
    }
    if (!closedSet.entries.empty()) {
        closedSet.selectionOriginalIndex = closedSet.entries.back().originalIndex;
        PushClosedSet(std::move(closedSet));
    }

    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (!removed->savedGroupId.empty()) {
        GroupStore::Instance().Remove(removed->savedGroupId);
    }

    if (removedSelectedGroup) {
        const TabLocation newSelection = m_tabs.SelectedLocation();
        if (newSelection.IsValid()) {
            NavigateToTab(newSelection);
        }
    }

    SaveSession();
}

void TabBand::OnEditGroupProperties(int groupIndex) {
    auto* group = m_tabs.GetGroup(groupIndex);
    if (!group) {
        return;
    }

    if (!group->savedGroupId.empty()) {
        OnShowOptionsDialog(OptionsDialogPage::kGroups, group->savedGroupId, true);
        return;
    }

    HWND hwnd = m_window ? m_window->GetHwnd() : nullptr;
    std::wstring name = group->name;
    COLORREF color = group->hasCustomOutline ? group->outlineColor : RGB(0, 120, 215);

    if (!PromptForTextInput(hwnd, L"Edit Island", L"Island name:", &name, &color)) {
        return;
    }

    std::wstring trimmed = TrimWhitespace(name);
    if (trimmed.empty()) {
        trimmed = group->name.empty() ? std::wstring(L"Island") : group->name;
    }

    const bool colorChanged = !group->hasCustomOutline || group->outlineColor != color;

    group->name = std::move(trimmed);
    group->hasCustomOutline = true;
    group->outlineColor = color;

    UpdateTabsUI();
    SyncSavedGroup(groupIndex);
    SaveSession();  // Immediately persist group name/color changes

    if (!group->savedGroupId.empty() && colorChanged) {
        auto& store = GroupStore::Instance();
        if (LoadGroupStoreForContext(L"TabBand::OnEditGroupProperties failed to load saved groups", store)) {
            store.UpdateColor(group->savedGroupId, color);
        }
    }
}

void TabBand::OnSetIslandLabel(int groupIndex, const std::wstring& label) {
    auto* group = m_tabs.GetGroup(groupIndex);
    if (!group) return;

    group->name = label;
    UpdateTabsUI();
    SyncSavedGroup(groupIndex);
    SaveSession();
}

void TabBand::OnDetachGroupRequested(int groupIndex) {
    const TabLocation selected = m_tabs.SelectedLocation();
    const bool wasSelectedInGroup = (selected.IsValid() && selected.groupIndex == groupIndex);

    auto removedGroup = m_tabs.TakeGroup(groupIndex);
    if (!removedGroup) {
        return;
    }
    CancelPendingPreviewForGroup(*removedGroup);

    std::shared_ptr<PendingWindowSeed> seed;
    if (!removedGroup->tabs.empty()) {
        seed = std::make_shared<PendingWindowSeed>();
        seed->type = WindowSeedType::Group;
        seed->group = std::move(*removedGroup);
        EnqueuePendingWindowSeed(seed);
        OpenTabInNewWindow(seed->group.tabs.front());
    }

    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }
    UpdateTabsUI();
    SyncAllSavedGroups();

    if (wasSelectedInGroup) {
        const TabLocation newSelection = m_tabs.SelectedLocation();
        if (newSelection.IsValid()) {
            NavigateToTab(newSelection);
        }
    }

    SaveSession();  // Immediately persist detached group
}

void TabBand::OnMoveTabRequested(TabLocation from, TabLocation to) {
    m_tabs.MoveTab(from, to);
    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist tab move
}

void TabBand::OnMoveGroupRequested(int fromGroup, int toGroup) {
    m_tabs.MoveGroup(fromGroup, toGroup);
    UpdateTabsUI();
    SaveSession();  // Immediately persist group move
}

void TabBand::OnMoveTabToNewGroup(TabLocation from, int insertIndex, bool headerVisible) {
    m_tabs.MoveTabToNewGroup(from, insertIndex, headerVisible);
    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist tab move to new group
}

std::optional<TabInfo> TabBand::DetachTabForTransfer(TabLocation location, bool* wasSelected,
                                                     bool ensurePlaceholderTab, bool* removedLastTab) {
    if (removedLastTab) {
        *removedLastTab = false;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const bool tabWasSelected = (selected.groupIndex == location.groupIndex && selected.tabIndex == location.tabIndex);
    if (wasSelected) {
        *wasSelected = tabWasSelected;
    }

    const bool wasLastTab = (m_tabs.TotalTabCount() == 1);
    auto removed = m_tabs.TakeTab(location);
    if (!removed) {
        if (wasSelected) {
            *wasSelected = false;
        }
        return std::nullopt;
    }

    CancelPendingPreviewForTab(*removed);

    if (removedLastTab) {
        *removedLastTab = wasLastTab;
    }

    if (m_tabs.TotalTabCount() == 0 && ensurePlaceholderTab) {
        EnsureTabForCurrentFolder();
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (tabWasSelected) {
        const TabLocation newSelection = m_tabs.SelectedLocation();
        if (newSelection.IsValid()) {
            NavigateToTab(newSelection);
        }
    }

    return removed;
}

TabLocation TabBand::InsertTransferredTab(TabInfo tab, int groupIndex, int tabIndex, bool createGroup,
                                          bool headerVisible, bool select) {
    if (createGroup) {
        groupIndex = m_tabs.CreateGroupAfter(groupIndex - 1, {}, headerVisible);
        tabIndex = 0;
    }

    if (groupIndex < 0) {
        groupIndex = 0;
    }
    const int groupCount = m_tabs.GroupCount();
    if (groupCount == 0) {
        m_tabs.CreateGroupAfter(-1, {}, headerVisible);
        groupIndex = 0;
        tabIndex = 0;
    } else if (groupIndex >= groupCount) {
        groupIndex = groupCount - 1;
    }

    auto* group = m_tabs.GetGroup(groupIndex);
    if (!group) {
        groupIndex = m_tabs.CreateGroupAfter(groupCount - 1, {}, headerVisible);
        group = m_tabs.GetGroup(groupIndex);
        tabIndex = 0;
    }
    if (group) {
        tabIndex = std::clamp(tabIndex, 0, static_cast<int>(group->tabs.size()));
    } else {
        tabIndex = 0;
    }

    TabLocation inserted = m_tabs.InsertTab(std::move(tab), groupIndex, tabIndex, select);
    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist inserted tab
    if (select && inserted.IsValid()) {
        NavigateToTab(inserted);
    }
    return inserted;
}

std::optional<TabGroup> TabBand::DetachGroupForTransfer(int groupIndex, bool* wasSelected) {
    const bool groupWasSelected = (m_tabs.SelectedLocation().groupIndex == groupIndex);
    if (wasSelected) {
        *wasSelected = groupWasSelected;
    }

    auto removed = m_tabs.TakeGroup(groupIndex);
    if (!removed) {
        if (wasSelected) {
            *wasSelected = false;
        }
        return std::nullopt;
    }

    CancelPendingPreviewForGroup(*removed);

    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (groupWasSelected) {
        const TabLocation newSelection = m_tabs.SelectedLocation();
        if (newSelection.IsValid()) {
            NavigateToTab(newSelection);
        }
    }

    SaveSession();  // Immediately persist group removal

    return removed;
}

int TabBand::InsertTransferredGroup(TabGroup group, int insertIndex, bool select) {
    if (insertIndex < 0) {
        insertIndex = 0;
    }
    insertIndex = m_tabs.InsertGroup(std::move(group), insertIndex);

    if (select) {
        if (auto* insertedGroup = m_tabs.GetGroup(insertIndex); insertedGroup) {
            if (!insertedGroup->tabs.empty()) {
                TabLocation location{insertIndex, 0};
                m_tabs.SetSelectedLocation(location);
                NavigateToTab(location);
            }
        }
    }

    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();  // Immediately persist transferred group

    return insertIndex;
}

void TabBand::OnSetGroupHeaderVisible(int groupIndex, bool visible) {
    m_tabs.SetGroupHeaderVisible(groupIndex, visible);
    UpdateTabsUI();
    SaveSession();  // Immediately persist group header visibility change
}

void TabBand::OnOpenTerminal(TabLocation location) {
    const std::wstring path = GetTabPath(location);
    if (path.empty() || !IsLikelyFileSystemPath(path)) {
        return;
    }
    std::wstring quoted = L"\"" + path + L"\"";
    if (LaunchShellExecute(L"wt.exe", L"-d " + quoted, path)) {
        return;
    }
    if (LaunchShellExecute(L"powershell.exe",
                           L"-NoExit -Command Set-Location -LiteralPath " + quoted, path)) {
        return;
    }
    LaunchShellExecute(L"cmd.exe", L"/K cd /d " + quoted, path);
}

void TabBand::OnOpenVSCode(TabLocation location) {
    const std::wstring path = GetTabPath(location);
    if (path.empty() || !IsLikelyFileSystemPath(path)) {
        return;
    }
    std::wstring quoted = L"\"" + path + L"\"";
    if (LaunchShellExecute(L"code.cmd", quoted, path)) {
        return;
    }
    LaunchShellExecute(L"code.exe", quoted, path);
}

void TabBand::OnCopyPath(TabLocation location) {
    const std::wstring path = GetTabPath(location);
    if (path.empty()) {
        return;
    }
    if (!OpenClipboard(m_window ? m_window->GetHwnd() : nullptr)) {
        return;
    }
    EmptyClipboard();
    const size_t bytes = (path.size() + 1) * sizeof(wchar_t);
    HGLOBAL global = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!global) {
        CloseClipboard();
        return;
    }
    wchar_t* buffer = static_cast<wchar_t*>(GlobalLock(global));
    if (!buffer) {
        GlobalFree(global);
        CloseClipboard();
        return;
    }
    memcpy(buffer, path.c_str(), bytes);
    GlobalUnlock(global);
    SetClipboardData(CF_UNICODETEXT, global);
    CloseClipboard();
}

void TabBand::OnFilesDropped(TabLocation location, const std::vector<std::wstring>& paths, bool move) {
    if (paths.empty()) {
        return;
    }
    PerformFileOperation(location, paths, move);
}

void TabBand::OnOpenFolderInNewTab(const std::wstring& path, bool select) {
    if (path.empty()) {
        return;
    }

    UniquePidl pidl = ParseDisplayName(path);
    if (!pidl) {
        pidl = ParseExplorerUrl(path);
    }
    if (!pidl && IsLikelyFileSystemPath(path)) {
        const DWORD attributes = GetFileAttributesW(path.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
            return;
        }
        pidl = ParseDisplayName(path);
    }
    if (!pidl) {
        return;
    }

    std::wstring name = GetDisplayName(pidl.get());
    if (name.empty()) {
        name = path;
    }
    std::wstring tooltip = GetCanonicalParsingName(pidl.get());
    if (tooltip.empty()) {
        tooltip = GetParsingName(pidl.get());
    }
    if (tooltip.empty()) {
        tooltip = name;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const int targetGroup = selected.groupIndex >= 0 ? selected.groupIndex : 0;
    TabLocation location = m_tabs.Add(std::move(pidl), name, tooltip, select, targetGroup);

    UpdateTabsUI();
    SyncAllSavedGroups();
    SaveSession();

    if (select && location.IsValid()) {
        NavigateToTab(location);
    }
}

void TabBand::CloseFrameWindowAsync() {
    HWND frame = GetFrameWindow();
    if (!frame) {
        return;
    }
    PostMessageW(frame, WM_CLOSE, 0, 0);
}

bool TabBand::TryRedirectToExistingWindow(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return false;
    }

    // Get the folder path from the PIDL
    std::wstring path = GetCanonicalParsingName(pidl);
    if (path.empty()) {
        path = GetParsingName(pidl);
    }
    if (path.empty()) {
        return false;
    }

    // Get our own top-level frame so we can exclude it from the search
    HWND ourFrame = GetFrameWindow();
    if (!ourFrame) {
        return false;
    }

    HWND candidate = FindShellTabsBandWindow(nullptr, ourFrame);

    if (!candidate) {
        return false;
    }

    // Send the folder path via WM_COPYDATA using the existing protocol
    COPYDATASTRUCT cds{};
    cds.dwData = SHELLTABS_COPYDATA_OPEN_FOLDER;
    cds.cbData = static_cast<DWORD>(path.size() * sizeof(wchar_t));
    cds.lpData = const_cast<wchar_t*>(path.c_str());

    const LRESULT result = SendMessageW(candidate, WM_COPYDATA,
        reinterpret_cast<WPARAM>(ourFrame), reinterpret_cast<LPARAM>(&cds));

    if (!result) {
        return false;
    }

    // Bring the target window to the foreground
    HWND targetFrame = GetAncestor(candidate, GA_ROOT);
    if (targetFrame) {
        if (IsIconic(targetFrame)) {
            ShowWindow(targetFrame, SW_RESTORE);
        }
        DWORD targetProcessId = 0;
        DWORD targetThread = GetWindowThreadProcessId(targetFrame, &targetProcessId);
        DWORD currentThread = GetCurrentThreadId();
        
        if (targetProcessId != 0) {
            AllowSetForegroundWindow(targetProcessId);
        }
        
        if (targetThread != currentThread) {
            AttachThreadInput(currentThread, targetThread, TRUE);
            SetForegroundWindow(targetFrame);
            BringWindowToTop(targetFrame);
            AttachThreadInput(currentThread, targetThread, FALSE);
        } else {
            SetForegroundWindow(targetFrame);
            BringWindowToTop(targetFrame);
        }
    }

    // Close this Explorer window
    PostMessageW(ourFrame, WM_CLOSE, 0, 0);
    return true;
}

void TabBand::EnsureTabPreview(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    if (selected.groupIndex != location.groupIndex || selected.tabIndex != location.tabIndex) {
        return;
    }

    CaptureActiveTabPreview();
}

HWND TabBand::GetFrameWindow() const {
    HWND candidate = nullptr;
    if (m_window) {
        HWND child = m_window->GetHwnd();
        if (child) {
            candidate = GetAncestor(child, GA_ROOT);
            if (!candidate) {
                candidate = child;
            }
        }
    }
    if (!candidate && m_site) {
        ComPtr<IOleWindow> oleWindow;
        if (SUCCEEDED(m_site.As(&oleWindow)) && oleWindow) {
            HWND siteWindow = nullptr;
            if (SUCCEEDED(oleWindow->GetWindow(&siteWindow)) && siteWindow) {
                candidate = GetAncestor(siteWindow, GA_ROOT);
                if (!candidate) {
                    candidate = siteWindow;
                }
            }
        }
    }
    return candidate;
}

TabManager::ExplorerWindowId TabBand::BuildWindowId() const {
    TabManager::ExplorerWindowId id;

    // Add validation for destroying state
    if (m_isDestroying.load()) {
        LogMessage(LogLevel::Warning, L"TabBand::BuildWindowId called during destruction");
        return id; // Return invalid ID if destroying
    }

    id.hwnd = GetFrameWindow();
    if (id.hwnd && !IsWindow(id.hwnd)) {
        LogMessage(LogLevel::Warning, L"TabBand::BuildWindowId invalid window handle");
        id.hwnd = nullptr; // Validate window handle
    }

    if (m_webBrowser) {
        // Add COM interface validation
        Microsoft::WRL::ComPtr<IUnknown> test;
        if (SUCCEEDED(m_webBrowser.As(&test)) && test) {
            id.frameCookie = reinterpret_cast<uintptr_t>(m_webBrowser.Get());
        } else {
            LogMessage(LogLevel::Warning, L"TabBand::BuildWindowId invalid webBrowser interface");
        }
    }
    return id;
}

void TabBand::CaptureActiveTabPreview() {
    if (!m_shellBrowser) {
        return;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const auto* tab = m_tabs.Get(selected);
    if (!tab || !tab->pidl) {
        return;
    }

    IShellView* rawView = nullptr;
    const HRESULT viewResult = m_shellBrowser->QueryActiveShellView(&rawView);
    if (FAILED(viewResult) || !rawView) {
        return;
    }

    ComPtr<IShellView> shellView;
    shellView.Attach(rawView);

    HWND viewWindow = nullptr;
    if (FAILED(shellView->GetWindow(&viewWindow)) || !viewWindow) {
        return;
    }

    // Use frame HWND as preview owner key.
    HWND frame = GetFrameWindow();
    std::wstring owner = frame ? std::to_wstring(reinterpret_cast<uintptr_t>(frame)) : std::wstring();
    PreviewCache::Instance().StorePreviewFromWindow(tab->pidl.get(), viewWindow, kPreviewImageSize,
                                                    owner);
}

namespace {

// Walk the view tree and find the scrollable list/grid window. Modern Explorer hosts
// either a SysListView32 or a UIItemsView (under DirectUIHWND). Fall back to the
// view window itself if neither is found.
HWND FindExplorerScrollable(HWND viewWindow) {
    if (!viewWindow) {
        return nullptr;
    }
    if (HWND lv = FindWindowExW(viewWindow, nullptr, WC_LISTVIEWW, nullptr)) {
        return lv;
    }
    if (HWND lvDescendant = FindDescendantByClassEnum(viewWindow, WC_LISTVIEWW)) {
        return lvDescendant;
    }
    return viewWindow;
}

}  // namespace

bool TabBand::GetCurrentScrollPosition(POINT& outPosition) {
    outPosition = {0, 0};

    if (!m_shellBrowser) {
        return false;
    }

    IShellView* rawView = nullptr;
    if (FAILED(m_shellBrowser->QueryActiveShellView(&rawView)) || !rawView) {
        return false;
    }

    ComPtr<IShellView> shellView;
    shellView.Attach(rawView);

    bool captured = false;

    // Strategy 1: Use IFolderView2::GetVisibleItem for a stable, view-agnostic
    // anchor. The encoded value is the topmost visible item index, stored in
    // outPosition.y so that the existing POINT plumbing carries it through.
    ComPtr<IFolderView2> folderView2;
    if (SUCCEEDED(shellView.As(&folderView2)) && folderView2) {
        int firstVisible = -1;
        HRESULT hr = folderView2->GetVisibleItem(-1, FALSE, &firstVisible);
        if (SUCCEEDED(hr) && firstVisible >= 0) {
            outPosition.x = -1;            // marker: y holds an item index
            outPosition.y = firstVisible;  // index of first visible item
            captured = true;
        }
    }

    // Strategy 2: Fallback to scroll info from the listview / view window.
    if (!captured) {
        HWND viewWindow = nullptr;
        if (FAILED(shellView->GetWindow(&viewWindow)) || !viewWindow) {
            return false;
        }
        HWND scrollWindow = FindExplorerScrollable(viewWindow);
        SCROLLINFO siVert = {sizeof(SCROLLINFO), SIF_POS, 0, 0, 0, 0, 0};
        SCROLLINFO siHorz = {sizeof(SCROLLINFO), SIF_POS, 0, 0, 0, 0, 0};
        const bool vert = GetScrollInfo(scrollWindow, SB_VERT, &siVert) != FALSE;
        const bool horz = GetScrollInfo(scrollWindow, SB_HORZ, &siHorz) != FALSE;
        if (vert || horz) {
            outPosition.x = horz ? siHorz.nPos : 0;
            outPosition.y = vert ? siVert.nPos : 0;
            captured = true;
        }
    }

    return captured;
}

bool TabBand::SetCurrentScrollPosition(const POINT& position) {
    if (!m_shellBrowser) {
        return false;
    }

    IShellView* rawView = nullptr;
    if (FAILED(m_shellBrowser->QueryActiveShellView(&rawView)) || !rawView) {
        return false;
    }

    ComPtr<IShellView> shellView;
    shellView.Attach(rawView);

    bool restored = false;

    // The x == -1 sentinel indicates that y carries an IFolderView item index.
    if (position.x == -1) {
        ComPtr<IFolderView> folderView;
        if (SUCCEEDED(shellView.As(&folderView)) && folderView) {
            const int index = position.y;
            int itemCount = 0;
            if (SUCCEEDED(folderView->ItemCount(SVGIO_ALLVIEW, &itemCount)) && index >= 0 &&
                index < itemCount) {
                // SVSI_POSITIONITEM scrolls the view so the item is visible.
                // SVSI_NOSTATECHANGE keeps focus / selection where they were.
                const HRESULT hr = folderView->SelectItem(
                    index, SVSI_POSITIONITEM | SVSI_NOSTATECHANGE);
                if (SUCCEEDED(hr)) {
                    restored = true;
                }
            }
        }
    }

    if (restored) {
        return true;
    }

    // Fallback: drive scroll bars directly.
    HWND viewWindow = nullptr;
    if (FAILED(shellView->GetWindow(&viewWindow)) || !viewWindow) {
        return false;
    }
    HWND scrollWindow = FindExplorerScrollable(viewWindow);

    if (position.y != 0) {
        SCROLLINFO siSet = {sizeof(SCROLLINFO), SIF_POS, 0, 0, 0, position.y, 0};
        SetScrollInfo(scrollWindow, SB_VERT, &siSet, TRUE);
        SendMessageW(scrollWindow, WM_VSCROLL, MAKEWPARAM(SB_THUMBPOSITION, position.y), 0);
        SendMessageW(scrollWindow, WM_VSCROLL, MAKEWPARAM(SB_ENDSCROLL, 0), 0);
        restored = true;
    }
    if (position.x > 0) {
        SCROLLINFO siSet = {sizeof(SCROLLINFO), SIF_POS, 0, 0, 0, position.x, 0};
        SetScrollInfo(scrollWindow, SB_HORZ, &siSet, TRUE);
        SendMessageW(scrollWindow, WM_HSCROLL, MAKEWPARAM(SB_THUMBPOSITION, position.x), 0);
        SendMessageW(scrollWindow, WM_HSCROLL, MAKEWPARAM(SB_ENDSCROLL, 0), 0);
        restored = true;
    }
    return restored;
}

void TabBand::SaveCurrentTabScrollPosition() {
    const TabLocation selected = m_tabs.SelectedLocation();
    auto* tab = m_tabs.Get(selected);
    if (!tab) {
        return;
    }

    POINT scrollPos{};
    if (GetCurrentScrollPosition(scrollPos)) {
        tab->scrollPosition = scrollPos;
        tab->hasScrollPosition = true;
        LogMessage(LogLevel::Info,
                   L"TabBand::SaveCurrentTabScrollPosition saved (%ld, %ld) for tab %ls",
                   scrollPos.x, scrollPos.y, tab->name.c_str());
    }

    // Capture Deep View State (View Mode, Icon Size, Selection)
    if (m_shellBrowser) {
        IShellView* rawView = nullptr;
        if (SUCCEEDED(m_shellBrowser->QueryActiveShellView(&rawView)) && rawView) {
            ComPtr<IShellView> shellView;
            shellView.Attach(rawView);
            ComPtr<IFolderView2> folderView2;
            if (SUCCEEDED(shellView.As(&folderView2)) && folderView2) {
                FOLDERVIEWMODE viewMode = FVM_AUTO;
                int iconSize = 0;
                if (SUCCEEDED(folderView2->GetViewModeAndIconSize(&viewMode, &iconSize))) {
                    tab->viewMode = static_cast<UINT>(viewMode);
                    tab->iconSize = iconSize;
                    tab->hasViewState = true;
                }
                
                tab->selectedItems.clear();
                ComPtr<IEnumIDList> enumIdList;
                if (SUCCEEDED(folderView2->Items(SVGIO_SELECTION, IID_PPV_ARGS(&enumIdList))) && enumIdList) {
                    PITEMID_CHILD childPidl = nullptr;
                    while (enumIdList->Next(1, &childPidl, nullptr) == S_OK) {
                        tab->selectedItems.emplace_back(childPidl);
                    }
                }
            }
        }
    }
}

bool TabBand::RestoreCurrentTabScrollPosition() {
    const TabLocation selected = m_tabs.SelectedLocation();
    const auto* tab = m_tabs.Get(selected);
    // Restore Deep View State first so scroll restoration applies to the correct layout
    if (tab->hasViewState && m_shellBrowser) {
        IShellView* rawView = nullptr;
        if (SUCCEEDED(m_shellBrowser->QueryActiveShellView(&rawView)) && rawView) {
            ComPtr<IShellView> shellView;
            shellView.Attach(rawView);
            ComPtr<IFolderView2> folderView2;
            if (SUCCEEDED(shellView.As(&folderView2)) && folderView2) {
                folderView2->SetViewModeAndIconSize(static_cast<FOLDERVIEWMODE>(tab->viewMode), tab->iconSize);
                
                // Clear existing selection
                folderView2->SelectItem(-1, SVSI_DESELECTOTHERS);
                
                // Restore selection
                for (const auto& selPidl : tab->selectedItems) {
                    // Try to find the item index
                    // int itemIndex = -1;
                    // if (SUCCEEDED(folderView2->IndexOf(selPidl.get(), &itemIndex)) && itemIndex >= 0) {
                    //     folderView2->SelectItem(itemIndex, SVSI_SELECT | SVSI_NOSTATECHANGE);
                    // }
                }
            }
        }
    }

    if (!tab->hasScrollPosition) {
        return true;  // nothing to restore — treat as success
    }

    // Skip if we have nothing meaningful to restore (default zeros).
    if (tab->scrollPosition.x == 0 && tab->scrollPosition.y == 0) {
        return true;
    }

    if (SetCurrentScrollPosition(tab->scrollPosition)) {
        LogMessage(LogLevel::Info,
                   L"TabBand::RestoreCurrentTabScrollPosition restored (%ld, %ld) for tab %ls",
                   tab->scrollPosition.x, tab->scrollPosition.y, tab->name.c_str());
        return true;
    }
    return false;
}

std::mutex TabBand::s_scrollRestoreTimerLock;
std::unordered_map<UINT_PTR, TabBand*> TabBand::s_scrollRestoreTimers;

void TabBand::ScheduleScrollRestoreRetries(TabLocation location) {
    // Cancel any in-flight retry for a previous tab switch.
    if (m_scrollRestoreTimerId) {
        KillTimer(nullptr, m_scrollRestoreTimerId);
        std::lock_guard<std::mutex> lock(s_scrollRestoreTimerLock);
        s_scrollRestoreTimers.erase(m_scrollRestoreTimerId);
        m_scrollRestoreTimerId = 0;
    }
    m_scrollRestore.location = location;
    m_scrollRestore.attemptsRemaining = 6;  // ~6 retries up to ~750ms total
    UINT_PTR id = SetTimer(nullptr, 0, 60, &TabBand::ScrollRestoreTimerProc);
    if (!id) {
        return;
    }
    m_scrollRestoreTimerId = id;
    std::lock_guard<std::mutex> lock(s_scrollRestoreTimerLock);
    s_scrollRestoreTimers[id] = this;
}

void CALLBACK TabBand::ScrollRestoreTimerProc(HWND, UINT, UINT_PTR timerId, DWORD) {
    TabBand* self = nullptr;
    {
        std::lock_guard<std::mutex> lock(s_scrollRestoreTimerLock);
        auto it = s_scrollRestoreTimers.find(timerId);
        if (it == s_scrollRestoreTimers.end()) {
            KillTimer(nullptr, timerId);
            return;
        }
        self = it->second;
    }
    if (self) {
        self->HandleScrollRestoreTimer(timerId);
    }
}

void TabBand::HandleScrollRestoreTimer(UINT_PTR timerId) {
    if (timerId != m_scrollRestoreTimerId) {
        KillTimer(nullptr, timerId);
        std::lock_guard<std::mutex> lock(s_scrollRestoreTimerLock);
        s_scrollRestoreTimers.erase(timerId);
        return;
    }

    // If selection changed since we scheduled, abort.
    const TabLocation selected = m_tabs.SelectedLocation();
    bool sameLocation = selected.groupIndex == m_scrollRestore.location.groupIndex &&
                        selected.tabIndex == m_scrollRestore.location.tabIndex;

    bool stop = !sameLocation;
    if (sameLocation) {
        if (RestoreCurrentTabScrollPosition()) {
            stop = true;
        } else if (--m_scrollRestore.attemptsRemaining <= 0) {
            stop = true;
        }
    }

    if (stop) {
        KillTimer(nullptr, timerId);
        std::lock_guard<std::mutex> lock(s_scrollRestoreTimerLock);
        s_scrollRestoreTimers.erase(timerId);
        m_scrollRestoreTimerId = 0;
    }
}

void TabBand::EnsureWindow() {
    if (m_window) {
        return;
    }

    LogMessage(LogLevel::Info, L"TabBand::EnsureWindow creating band window (this=%p)", this);
    HWND parent = nullptr;
    if (m_site) {
        ComPtr<IOleWindow> oleWindow;
        if (SUCCEEDED(m_site.As(&oleWindow)) && oleWindow) {
            oleWindow->GetWindow(&parent);
        }
    }

    auto window = std::make_unique<TabBandWindow>(this);
    if (window->Create(parent)) {
        m_window = std::move(window);
        TabBandDockMode preferred = m_requestedDockMode;
        if (preferred == TabBandDockMode::kAutomatic) {
            TabBandDockMode optionDock = m_optionsLoaded ? m_options.tabDockMode : TabBandDockMode::kAutomatic;
            if (optionDock != TabBandDockMode::kAutomatic) {
                preferred = optionDock;
            }
        }
        if (preferred == TabBandDockMode::kAutomatic) {
            preferred = TabBandDockMode::kTop;
        }
        if (m_requestedDockMode == TabBandDockMode::kAutomatic) {
            m_requestedDockMode = preferred;
        }
        if (m_window) {
            m_window->SetPreferredDockMode(preferred);
        }
        LogMessage(LogLevel::Info, L"TabBand::EnsureWindow created window hwnd=%p",
                   m_window ? m_window->GetHwnd() : nullptr);
        if (m_sessionFlushTimerPending) {
            StartSessionFlushTimer();
        }
    } else {
        LogMessage(LogLevel::Error, L"TabBand::EnsureWindow failed to create window");
    }
}

void TabBand::EnsureOptionsLoaded() const {
    if (m_optionsLoaded) {
        return;
    }
    auto& store = OptionsStore::Instance();
    LoadOptionsStoreForContext(L"TabBand::EnsureOptionsLoaded failed to load options", store);
    m_options = store.Get();
    m_optionsLoaded = true;
}

void TabBand::DisconnectSite() {
    // Thread-safe cleanup with mutex protection
    std::lock_guard<std::mutex> lock(m_cleanupMutex);

    // Check if already destroyed to prevent double cleanup
    if (m_isDestroying.load()) {
        LogMessage(LogLevel::Info, L"TabBand::DisconnectSite already in progress (this=%p)", this);
        return;
    }

    // Mark as destroying to prevent concurrent access
    m_isDestroying = true;

    LogMessage(LogLevel::Info, L"TabBand::DisconnectSite (this=%p)", this);

    // Cancel any in-flight scroll-restore retry timer first.
    if (m_scrollRestoreTimerId) {
        KillTimer(nullptr, m_scrollRestoreTimerId);
        std::lock_guard<std::mutex> timerLock(s_scrollRestoreTimerLock);
        s_scrollRestoreTimers.erase(m_scrollRestoreTimerId);
        m_scrollRestoreTimerId = 0;
    }

    // Step 1: Cancel background operations safely
    try {
        CancelInitializationWorker();
        m_backgroundInitializationActive = false;
        m_sessionPersistenceReady = false;
        m_pendingGroupSeed.reset();
        m_pendingStandaloneSeed = false;
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during worker cancellation");
    }

    // Step 2: Save session before cleanup (with error handling)
    try {
        SaveSession();
        StopSessionFlushTimer();
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during session cleanup");
    }

    // Step 3: Cancel pending operations safely
    try {
        m_tabs.ClearWindowId();

        // Cancel previews with bounds checking
        for (int groupIndex = 0; groupIndex < m_tabs.GroupCount(); ++groupIndex) {
            if (const TabGroup* group = m_tabs.GetGroup(groupIndex)) {
                CancelPendingPreviewForGroup(*group);
            }
        }

        HWND frame = GetFrameWindow();
        if (frame) {
            std::wstring owner = std::to_wstring(reinterpret_cast<uintptr_t>(frame));
            PreviewCache::Instance().CancelPendingCapturesForOwner(owner);
        }
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during preview cleanup");
    }

    // Step 4: Disconnect browser events safely
    try {
        if (m_browserEvents) {
            auto browserEvents = std::move(m_browserEvents);
            browserEvents->DetachOwner();

            const bool dispatching = browserEvents->IsDispatching();
            const HRESULT hr = browserEvents->Disconnect();
            if (dispatching || FAILED(hr)) {
                LogMessage(LogLevel::Warning,
                           L"TabBand::DisconnectSite retained detached BrowserEvents sink (dispatching=%d, hr=0x%08X)",
                           dispatching ? 1 : 0,
                           static_cast<unsigned int>(hr));
                browserEvents.release();
            }
        }
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during browser events disconnect");
    }

    // Step 5: Release COM interfaces safely (order is important)
    try {
        m_webBrowser.Reset();
        m_shellBrowser.Reset();
        m_site.Reset();
        m_siteOleWindow.Reset();
        m_dockingSite.Reset();
        m_dockingFrame.Reset();
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during COM interface cleanup");
    }

    // Step 6: Cleanup window safely
    try {
        if (m_window) {
            m_window->SetSite(nullptr);
            m_window->Destroy();
            m_window.reset();
        }
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during window cleanup");
    }

    // Step 7: Final cleanup
    try {
        m_tabs.Clear();
        m_internalNavigation = false;
        m_allowExternalNewWindows = 0;
        if (m_sessionStore) {
            CrashRecoveryLogger::Instance().ClearWindowState(m_sessionStore->Slot());
            m_sessionStore.reset();
        }
    }
    catch (...) {
        LogMessage(LogLevel::Error, L"TabBand::DisconnectSite exception during final cleanup");
    }

    LogMessage(LogLevel::Info, L"TabBand::DisconnectSite completed (this=%p)", this);
}

void TabBand::InitializeTabs() {
    LogScope scope(L"TabBand::InitializeTabs");
    CancelInitializationWorker();
    StopSessionFlushTimer();
    m_backgroundInitializationActive = false;
    m_sessionPersistenceReady = false;
    m_lastSessionUnclean = false;

    for (int groupIndex = 0; groupIndex < m_tabs.GroupCount(); ++groupIndex) {
        if (const TabGroup* group = m_tabs.GetGroup(groupIndex)) {
            CancelPendingPreviewForGroup(*group);
        }
    }
    HWND frame = GetFrameWindow();
    if (frame) {
        std::wstring owner = std::to_wstring(reinterpret_cast<uintptr_t>(frame));
        PreviewCache::Instance().CancelPendingCapturesForOwner(owner);
    }
    m_tabs.Clear();

    m_pendingGroupSeed.reset();
    m_pendingStandaloneSeed = false;
    m_hadPendingSeed = false;

    std::shared_ptr<PendingWindowSeed> pendingSeed = DequeuePendingWindowSeed();
    if (pendingSeed) {
        if (pendingSeed->type == WindowSeedType::Group) {
            m_pendingGroupSeed = std::move(pendingSeed->group);
        } else {
            m_pendingStandaloneSeed = true;
        }
    }

    EnsureSessionStore();

    UniquePidl placeholder = QueryCurrentFolder();
    if (placeholder) {
        std::wstring name = GetDisplayName(placeholder.get());
        if (name.empty()) {
            name = L"Tab";
        }
        const bool hideHeader = !pendingSeed || pendingSeed->type == WindowSeedType::StandaloneTab;
        TabLocation location = m_tabs.Add(std::move(placeholder), name, name, true);
        if (location.IsValid() && hideHeader) {
            m_tabs.SetGroupHeaderVisible(location.groupIndex, false);
        }
        LogMessage(LogLevel::Info, L"TabBand::InitializeTabs seeded placeholder tab %ls", name.c_str());
    }

    const bool hasPendingSeed = static_cast<bool>(pendingSeed);
    m_hadPendingSeed = hasPendingSeed;
    const uint64_t sequence = ++m_initializationSequence;
    m_backgroundInitializationActive = true;
    m_initializationThread = std::jthread([this, sequence, hasPendingSeed](std::stop_token stopToken) {
        RunBackgroundInitialization(std::move(stopToken), sequence, hasPendingSeed);
    });
}

void TabBand::UpdateTabsUI() {
    const auto items = m_tabs.BuildView();
    LogMessage(LogLevel::Info, L"TabBand::UpdateTabsUI applying %llu items",
               static_cast<unsigned long long>(items.size()));
    if (m_window) {
        m_window->SetTabs(items);
    }
    SaveSession();
}

void TabBand::EnsureSessionStore() {
    if (m_sessionStore) {
        return;
    }
    LogMessage(LogLevel::Info, L"TabBand::EnsureSessionStore creating store (this=%p)", this);
    m_sessionStore = std::make_unique<SessionStore>();
}

bool TabBand::RestoreSession() {
    if (!m_sessionStore) {
        return false;
    }

    SessionData data;
    if (!m_sessionStore->Load(data)) {
        LogMessage(LogLevel::Warning, L"TabBand::RestoreSession load failed");
        return false;
    }

    return RestoreSessionFromData(data);
}

bool TabBand::RestoreSessionFromData(const SessionData& data) {
    LogMessage(LogLevel::Info, L"TabBand::RestoreSession loaded %llu groups",
               static_cast<unsigned long long>(data.groups.size()));

    m_dockMode = data.dockMode;
    if (m_dockMode == TabBandDockMode::kAutomatic) {
        if (!m_optionsLoaded) {
            EnsureOptionsLoaded();
        }
        m_dockMode = m_options.tabDockMode;
    }
    m_requestedDockMode = m_dockMode;
    if (m_window && m_requestedDockMode != TabBandDockMode::kAutomatic) {
        m_window->SetPreferredDockMode(m_requestedDockMode);
    }

    std::vector<TabGroup> groups;
    groups.reserve(data.groups.size());
    for (const auto& groupData : data.groups) {
        TabGroup group;
        group.name = groupData.name.empty() ? L"Island" : groupData.name;
        group.collapsed = groupData.collapsed;
        group.headerVisible = groupData.headerVisible;
        group.hasCustomOutline = groupData.hasOutline;
        group.outlineColor = groupData.outlineColor;
        group.outlineStyle = groupData.outlineStyle;
        group.savedGroupId = groupData.savedGroupId;
        std::vector<const SessionTab*> pinnedTabs;
        std::vector<const SessionTab*> unpinnedTabs;
        pinnedTabs.reserve(groupData.tabs.size());
        unpinnedTabs.reserve(groupData.tabs.size());
        for (const auto& tabData : groupData.tabs) {
            if (tabData.pinned) {
                pinnedTabs.push_back(&tabData);
            } else {
                unpinnedTabs.push_back(&tabData);
            }
        }

        auto appendTab = [&](const SessionTab& tabData) {
            UniquePidl pidl = ParseDisplayName(tabData.path);
            if (!pidl) {
                return;
            }
            TabInfo tab;
            tab.pidl = std::move(pidl);
            tab.name = tabData.name;
            if (tab.name.empty()) {
                tab.name = GetDisplayName(tab.pidl.get());
            }
            if (tab.name.empty()) {
                tab.name = L"Tab";
            }
            tab.tooltip = tabData.tooltip.empty() ? tab.name : tabData.tooltip;
            tab.hidden = tabData.hidden;
            tab.pinned = tabData.pinned;
            tab.path = tabData.path;
            tab.lastActivatedTick = tabData.lastActivatedTick;
            tab.activationOrdinal = tabData.activationOrdinal;
            tab.hasScrollPosition = tabData.hasScrollPosition;
            tab.scrollPosition.x = tabData.scrollX;
            tab.scrollPosition.y = tabData.scrollY;
            tab.navigationHistory.currentIndex = tabData.historyIndex;
            for (const auto& he : tabData.history) {
                NavigationHistoryEntry entry;
                entry.path = he.path;
                entry.name = he.name;
                entry.timestamp = he.timestamp;
                entry.pidl = ParseDisplayName(he.path);
                tab.navigationHistory.entries.push_back(std::move(entry));
            }
            tab.RefreshNormalizedLookupKey();
            group.tabs.emplace_back(std::move(tab));
        };

        for (const SessionTab* entry : pinnedTabs) {
            appendTab(*entry);
        }
        for (const SessionTab* entry : unpinnedTabs) {
            appendTab(*entry);
        }
        if (!group.tabs.empty() || !group.savedGroupId.empty()) {
            groups.emplace_back(std::move(group));
        }
    }

    if (groups.empty()) {
        return false;
    }

    m_restoringSession = true;
    m_tabs.Restore(std::move(groups), data.selectedGroup, data.selectedTab, data.groupSequence);
    m_restoringSession = false;

    m_closedTabHistory.clear();
    if (data.lastClosed) {
        if (auto restored = BuildClosedSetFromSession(*data.lastClosed)) {
            m_closedTabHistory.push_back(std::move(*restored));
        }
    }
    return true;
}

void TabBand::SaveSession() {
    if (m_restoringSession) {
        LogMessage(LogLevel::Info, L"TabBand::SaveSession skipped - restoring session");
        return;
    }

    if (!m_sessionPersistenceReady) {
        LogMessage(LogLevel::Warning, L"TabBand::SaveSession skipped - persistence not ready");
        return;
    }

    EnsureSessionStore();
    if (!m_sessionStore) {
        LogMessage(LogLevel::Warning, L"TabBand::SaveSession skipped - no session store");
        return;
    }

    EnsureOptionsLoaded();

    SessionData data;
    const TabLocation selected = m_tabs.SelectedLocation();
    data.selectedGroup = selected.groupIndex;
    data.selectedTab = selected.tabIndex;
    data.groupSequence = m_tabs.NextGroupSequence();
    data.dockMode = m_dockMode != TabBandDockMode::kAutomatic ? m_dockMode : m_requestedDockMode;
    if (data.dockMode == TabBandDockMode::kAutomatic) {
        data.dockMode = m_options.tabDockMode;
    }

    const int groupCount = m_tabs.GroupCount();
    for (int i = 0; i < groupCount; ++i) {
        const TabGroup* group = m_tabs.GetGroup(i);
        if (!group) {
            continue;
        }

        SessionGroup storedGroup;
        storedGroup.name = group->name;
        storedGroup.collapsed = group->collapsed;
        storedGroup.headerVisible = group->headerVisible;
        storedGroup.hasOutline = group->hasCustomOutline;
        storedGroup.outlineColor = group->outlineColor;
        storedGroup.outlineStyle = group->outlineStyle;
        storedGroup.savedGroupId = group->savedGroupId;

        auto appendTab = [&](const TabInfo& tab) {
            SessionTab storedTab;
            storedTab.path = tab.path;
            if (storedTab.path.empty()) {
                storedTab.path = GetParsingName(tab.pidl.get());
            }
            if (storedTab.path.empty()) {
                return false;
            }
            storedTab.name = tab.name;
            storedTab.tooltip = tab.tooltip;
            storedTab.hidden = tab.hidden;
            storedTab.pinned = tab.pinned;
            storedTab.lastActivatedTick = tab.lastActivatedTick;
            storedTab.activationOrdinal = tab.activationOrdinal;
            storedTab.hasScrollPosition = tab.hasScrollPosition;
            storedTab.scrollX = tab.scrollPosition.x;
            storedTab.scrollY = tab.scrollPosition.y;
            storedTab.historyIndex = tab.navigationHistory.currentIndex;
            for (const auto& entry : tab.navigationHistory.entries) {
                SessionHistoryEntry he;
                he.path = entry.path;
                if (he.path.empty() && entry.pidl) {
                    he.path = GetParsingName(entry.pidl.get());
                }
                he.name = entry.name;
                he.timestamp = entry.timestamp;
                storedTab.history.push_back(std::move(he));
            }
            storedGroup.tabs.emplace_back(std::move(storedTab));
            return true;
        };

        for (const auto& tab : group->tabs) {
            if (tab.pinned) {
                appendTab(tab);
            }
        }
        for (const auto& tab : group->tabs) {
            if (!tab.pinned) {
                appendTab(tab);
            }
        }

        if (!storedGroup.tabs.empty() || !storedGroup.savedGroupId.empty()) {
            data.groups.emplace_back(std::move(storedGroup));
        }
    }

    if (!m_closedTabHistory.empty()) {
        if (auto stored = BuildSessionClosedSet(m_closedTabHistory.back())) {
            data.lastClosed = std::move(stored);
        }
    }

    if (data.groups.empty()) {
        LogMessage(LogLevel::Warning, L"TabBand::SaveSession skipped - no groups to save");
        return;
    }

    int totalTabs = 0;
    std::vector<std::wstring> allPaths;
    int globalSelectedIndex = -1;
    for (int g = 0; g < data.groups.size(); ++g) {
        const auto& group = data.groups[g];
        for (int t = 0; t < group.tabs.size(); ++t) {
            allPaths.push_back(group.tabs[t].path);
            if (g == data.selectedGroup && t == data.selectedTab) {
                globalSelectedIndex = static_cast<int>(allPaths.size() - 1);
            }
        }
        totalTabs += static_cast<int>(group.tabs.size());
    }

    if (m_sessionStore) {
        CrashRecoveryLogger::Instance().LogWindowState(m_sessionStore->Slot(), allPaths, globalSelectedIndex);
    }

    if (m_sessionStore->Save(data)) {
        LogMessage(LogLevel::Info, L"TabBand::SaveSession saved %d groups, %d tabs",
                   static_cast<int>(data.groups.size()), totalTabs);
    } else {
        LogMessage(LogLevel::Error, L"TabBand::SaveSession FAILED to save %d groups, %d tabs",
                   static_cast<int>(data.groups.size()), totalTabs);
    }
}

void TabBand::StartSessionFlushTimer() {
    if (!m_sessionStore) {
        m_sessionFlushTimerActive = false;
        m_sessionFlushTimerPending = false;
        return;
    }

    HWND hwnd = m_window ? m_window->GetHwnd() : nullptr;
    if (!hwnd) {
        m_sessionFlushTimerPending = true;
        return;
    }

    if (m_sessionFlushTimerActive) {
        m_sessionFlushTimerPending = false;
        return;
    }

    // Reduced from 15s to 3s for more frequent crash recovery updates.
    // FILE_FLAG_WRITE_THROUGH ensures data hits disk immediately on each save.
    constexpr UINT kSessionFlushIntervalMs = 3000;
    if (SetTimer(hwnd, TabBandWindow::SessionFlushTimerId(), kSessionFlushIntervalMs, nullptr)) {
        m_sessionFlushTimerActive = true;
        m_sessionFlushTimerPending = false;
    } else {
        LogMessage(LogLevel::Warning, L"TabBand::StartSessionFlushTimer failed (hwnd=%p, error=%lu)", hwnd,
                   GetLastError());
        m_sessionFlushTimerPending = true;
    }
}

void TabBand::StopSessionFlushTimer() {
    m_sessionFlushTimerPending = false;
    if (!m_sessionFlushTimerActive) {
        return;
    }

    HWND hwnd = m_window ? m_window->GetHwnd() : nullptr;
    if (hwnd) {
        KillTimer(hwnd, TabBandWindow::SessionFlushTimerId());
    }
    m_sessionFlushTimerActive = false;
}

void TabBand::OnPeriodicSessionFlush() {
    SaveSession();
}

void TabBand::RetrySessionClaim() {
    if (!m_sessionStore) {
        return;
    }

    auto& coordinator = SessionCoordinator::Instance();
    if (!coordinator.HasPendingData()) {
        LogMessage(LogLevel::Info, L"TabBand::RetrySessionClaim no pending data");
        return;
    }

    SessionData data;
    if (!m_sessionStore->Load(data)) {
        LogMessage(LogLevel::Warning, L"TabBand::RetrySessionClaim claim still suppressed");
        return;
    }

    LogMessage(LogLevel::Info, L"TabBand::RetrySessionClaim restoring %d groups",
               static_cast<int>(data.groups.size()));

    m_tabs.Clear();
    RestoreSessionFromData(data);
    if (m_window && m_window->GetHwnd()) {
        InvalidateRect(m_window->GetHwnd(), nullptr, TRUE);
    }
    SaveSession();
}

TabBand::ClosedGroupMetadata TabBand::CaptureGroupMetadata(const TabGroup& group) const {
    ClosedGroupMetadata metadata;
    metadata.name = group.name;
    metadata.collapsed = group.collapsed;
    metadata.headerVisible = group.headerVisible;
    metadata.hasOutline = group.hasCustomOutline;
    metadata.outlineColor = group.outlineColor;
    metadata.outlineStyle = group.outlineStyle;
    metadata.savedGroupId = group.savedGroupId;
    return metadata;
}

void TabBand::EnsureTabPath(TabInfo& tab) const {
    if (tab.path.empty()) {
        tab.path = GetParsingName(tab.pidl.get());
    }
    if (tab.normalizedLookupKey.empty()) {
        tab.RefreshNormalizedLookupKey();
    }
}

void TabBand::PushClosedSet(ClosedTabSet set) {
    if (set.entries.empty()) {
        return;
    }
    constexpr size_t kMaxHistory = 16;
    m_closedTabHistory.push_back(std::move(set));
    if (m_closedTabHistory.size() > kMaxHistory) {
        m_closedTabHistory.erase(m_closedTabHistory.begin());
    }
}

std::optional<SessionClosedSet> TabBand::BuildSessionClosedSet(const ClosedTabSet& set) const {
    if (set.entries.empty()) {
        return std::nullopt;
    }

    SessionClosedSet stored;
    stored.groupIndex = set.groupIndex;
    stored.groupRemoved = set.groupRemoved;
    stored.selectionIndex = set.selectionOriginalIndex;

    if (set.groupInfo) {
        stored.hasGroupInfo = true;
        stored.groupInfo.name = set.groupInfo->name;
        stored.groupInfo.collapsed = set.groupInfo->collapsed;
        stored.groupInfo.headerVisible = set.groupInfo->headerVisible;
        stored.groupInfo.hasOutline = set.groupInfo->hasOutline;
        stored.groupInfo.outlineColor = set.groupInfo->outlineColor;
        stored.groupInfo.outlineStyle = set.groupInfo->outlineStyle;
        stored.groupInfo.savedGroupId = set.groupInfo->savedGroupId;
    }

    for (const auto& entry : set.entries) {
        SessionClosedTab storedTab;
        storedTab.index = entry.originalIndex;
        storedTab.tab.name = entry.tab.name;
        storedTab.tab.tooltip = entry.tab.tooltip;
        storedTab.tab.hidden = entry.tab.hidden;
        storedTab.tab.pinned = entry.tab.pinned;
        storedTab.tab.path = entry.tab.path;
        if (storedTab.tab.path.empty()) {
            storedTab.tab.path = GetParsingName(entry.tab.pidl.get());
        }
        if (storedTab.tab.path.empty()) {
            return std::nullopt;
        }
        stored.tabs.emplace_back(std::move(storedTab));
    }

    if (stored.tabs.empty()) {
        return std::nullopt;
    }

    return stored;
}

std::optional<TabBand::ClosedTabSet> TabBand::BuildClosedSetFromSession(const SessionClosedSet& stored) const {
    if (stored.tabs.empty()) {
        return std::nullopt;
    }

    ClosedTabSet set;
    set.groupIndex = stored.groupIndex;
    set.groupRemoved = stored.groupRemoved;
    set.selectionOriginalIndex = stored.selectionIndex;

    if (stored.hasGroupInfo) {
        ClosedGroupMetadata metadata;
        metadata.name = stored.groupInfo.name;
        metadata.collapsed = stored.groupInfo.collapsed;
        metadata.headerVisible = stored.groupInfo.headerVisible;
        metadata.hasOutline = stored.groupInfo.hasOutline;
        metadata.outlineColor = stored.groupInfo.outlineColor;
        metadata.outlineStyle = stored.groupInfo.outlineStyle;
        metadata.savedGroupId = stored.groupInfo.savedGroupId;
        set.groupInfo = std::move(metadata);
    }

    for (const auto& storedTab : stored.tabs) {
        UniquePidl pidl = ParseDisplayName(storedTab.tab.path);
        if (!pidl) {
            continue;
        }

        TabInfo tab;
        tab.pidl = std::move(pidl);
        tab.name = storedTab.tab.name;
        if (tab.name.empty()) {
            tab.name = GetDisplayName(tab.pidl.get());
        }
        if (tab.name.empty()) {
            tab.name = L"Tab";
        }
        tab.tooltip = storedTab.tab.tooltip.empty() ? tab.name : storedTab.tab.tooltip;
        tab.hidden = storedTab.tab.hidden;
        tab.pinned = storedTab.tab.pinned;
        tab.path = storedTab.tab.path;
        EnsureTabPath(tab);
        set.entries.push_back({storedTab.index, std::move(tab)});
    }

    if (set.entries.empty()) {
        return std::nullopt;
    }

    return set;
}

void TabBand::ApplyOptionsChanges(const ShellTabsOptions& previousOptions) {
    if (previousOptions.tabDockMode != m_options.tabDockMode) {
        if (m_requestedDockMode == previousOptions.tabDockMode ||
            m_requestedDockMode == TabBandDockMode::kAutomatic) {
            m_requestedDockMode = m_options.tabDockMode;
            TabBandDockMode preferred = m_requestedDockMode;
            if (preferred == TabBandDockMode::kAutomatic) {
                preferred = TabBandDockMode::kTop;
            }
            if (m_window) {
                m_window->SetPreferredDockMode(preferred);
            }
        }
    }

    if (!previousOptions.persistGroupPaths && m_options.persistGroupPaths) {
        SyncAllSavedGroups();
    }

    const bool backgroundChanged =
        previousOptions.enableBreadcrumbGradient != m_options.enableBreadcrumbGradient;
    const bool fontChanged =
        previousOptions.enableBreadcrumbFontGradient != m_options.enableBreadcrumbFontGradient;
    const bool backgroundTransparencyChanged =
        previousOptions.breadcrumbGradientTransparency != m_options.breadcrumbGradientTransparency;
    const bool fontBrightnessChanged =
        previousOptions.breadcrumbFontBrightness != m_options.breadcrumbFontBrightness;
    const bool backgroundColorsChanged =
        previousOptions.useCustomBreadcrumbGradientColors != m_options.useCustomBreadcrumbGradientColors ||
        previousOptions.breadcrumbGradientStartColor != m_options.breadcrumbGradientStartColor ||
        previousOptions.breadcrumbGradientEndColor != m_options.breadcrumbGradientEndColor;
    const bool fontColorsChanged =
        previousOptions.useCustomBreadcrumbFontColors != m_options.useCustomBreadcrumbFontColors ||
        previousOptions.breadcrumbFontGradientStartColor != m_options.breadcrumbFontGradientStartColor ||
        previousOptions.breadcrumbFontGradientEndColor != m_options.breadcrumbFontGradientEndColor;
    const bool tabColorsChanged =
        previousOptions.useCustomTabSelectedColor != m_options.useCustomTabSelectedColor ||
        previousOptions.customTabSelectedColor != m_options.customTabSelectedColor ||
        previousOptions.useCustomTabUnselectedColor != m_options.useCustomTabUnselectedColor ||
        previousOptions.customTabUnselectedColor != m_options.customTabUnselectedColor;
    const bool glowEnabledChanged = previousOptions.enableNeonGlow != m_options.enableNeonGlow;
    const bool glowCustomChanged =
        previousOptions.useCustomNeonGlowColors != m_options.useCustomNeonGlowColors;
    const bool glowGradientChanged =
        previousOptions.useNeonGlowGradient != m_options.useNeonGlowGradient;
    const bool glowColorChanged = previousOptions.neonGlowPrimaryColor != m_options.neonGlowPrimaryColor ||
                                  previousOptions.neonGlowSecondaryColor != m_options.neonGlowSecondaryColor;
    const bool glowPaletteChanged = previousOptions.glowPalette != m_options.glowPalette;
    const bool accentColorsChanged =
        previousOptions.useExplorerAccentColors != m_options.useExplorerAccentColors;

    // Check for folder background changes
    const bool folderBackgroundEnabledChanged =
        previousOptions.enableFolderBackgrounds != m_options.enableFolderBackgrounds;
    const bool universalBackgroundChanged =
        previousOptions.universalFolderBackgroundImage != m_options.universalFolderBackgroundImage;
    const bool folderBackgroundsChanged =
        previousOptions.folderBackgroundEntries != m_options.folderBackgroundEntries;
    if (backgroundChanged || fontChanged || backgroundTransparencyChanged || fontBrightnessChanged ||
        backgroundColorsChanged || fontColorsChanged || tabColorsChanged || glowEnabledChanged ||
        glowCustomChanged || glowGradientChanged || glowColorChanged || glowPaletteChanged ||
        accentColorsChanged || folderBackgroundEnabledChanged || universalBackgroundChanged ||
        folderBackgroundsChanged) {
        const UINT message = GetOptionsChangedMessage();
        if (message != 0) {
            SendNotifyMessageW(HWND_BROADCAST, message, 0, 0);
        }

    }
}


UniquePidl TabBand::QueryCurrentFolder() const {
    return GetCurrentFolderPidL(m_shellBrowser, m_webBrowser);
}

void TabBand::CancelPendingPreviewForTab(const TabInfo& tab) const {
    if (tab.pidl) {
        PreviewCache::Instance().CancelPendingCapturesForKey(tab.pidl.get());
    }
}

void TabBand::CancelPendingPreviewForGroup(const TabGroup& group) const {
    for (const auto& tab : group.tabs) {
        CancelPendingPreviewForTab(tab);
    }
}

namespace {
    HRESULT SafeBrowseObject(IShellBrowser* browser, PCIDLIST_ABSOLUTE pidl, UINT flags) {
        __try {
            return browser->BrowseObject(pidl, flags);
        } __except(EXCEPTION_EXECUTE_HANDLER) {
            return E_FAIL;
        }
    }
}

void TabBand::NavigateToTab(TabLocation location) {
    LogMessage(LogLevel::Info, L"NavigateToTab location=(group=%d,tab=%d) shellBrowser=%p",
               location.groupIndex, location.tabIndex, m_shellBrowser.Get());
    if (!m_shellBrowser) {
        LogMessage(LogLevel::Warning, L"NavigateToTab: no shell browser, aborting");
        return;
    }
    auto* tab = m_tabs.Get(location);
    if (!tab) {
        LogMessage(LogLevel::Warning, L"NavigateToTab: tab is null, aborting");
        return;
    }

    if (tab->hibernated) {
        m_tabs.WakeHibernatedTab(location);
    }

    if (!tab->pidl) {
        LogMessage(LogLevel::Warning, L"NavigateToTab: pidl is null, aborting");
        return;
    }

    const auto current = m_tabs.SelectedLocation();
    if (current.groupIndex != location.groupIndex || current.tabIndex != location.tabIndex) {
        // Save scroll position and preview before navigating away
        SaveCurrentTabScrollPosition();
        CaptureActiveTabPreview();
    }

    m_tabs.SetGroupCollapsed(location.groupIndex, false);
    m_tabs.SetSelectedLocation(location);
    // Don't SaveSession() here — it does synchronous disk I/O with
    // FILE_FLAG_WRITE_THROUGH which is the dominant pause when switching tabs.
    // The 3-second periodic flush timer (StartSessionFlushTimer) and the
    // explicit save in DisconnectSite are sufficient to persist selection.
    m_internalNavigation = true;
    EnsureFtpNamespaceBinding(tab->pidl.get());
    LogMessage(LogLevel::Info, L"NavigateToTab: calling BrowseObject");
    HRESULT hr = E_FAIL;
    // Removable media (USB sticks, network drives) can be yanked between tab
    // creation and tab activation. Navigating to a stale PIDL into the shell
    // can produce structured exceptions deep inside Explorer/COM — guard the
    // call so the deskband does not bring down the host.
    hr = SafeBrowseObject(m_shellBrowser.Get(), tab->pidl.get(), SBSP_SAMEBROWSER | SBSP_WRITENOHISTORY);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning,
                   L"NavigateToTab: BrowseObject SEH exception (likely stale PIDL after device removal)");
    }
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"NavigateToTab: BrowseObject failed (hr=0x%08X)", hr);
        m_internalNavigation = false;
    } else {
        LogMessage(LogLevel::Info, L"NavigateToTab: BrowseObject succeeded");
        
        // Feature: Scriptable Tab Hooks (Automation)
        std::wstring hookScript = L"C:\\ShellTabsHooks\\" + tab->name + L".ps1";
        if (GetFileAttributesW(hookScript.c_str()) != INVALID_FILE_ATTRIBUTES) {
            ShellExecuteW(nullptr, L"open", L"powershell.exe", 
                (L"-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + hookScript + L"\"").c_str(), 
                nullptr, SW_HIDE);
        }
    }
}

void TabBand::EnsureTabForCurrentFolder() {
    UniquePidl current = QueryCurrentFolder();
    if (!current) {
        // Fallback for virtual namespaces (Control Panel, Network, Recycle Bin)
        // where IFolderView::GetFolder may not produce a usable PIDL. Re-parse the
        // browser's location URL — this almost always succeeds for shell:: paths.
        if (m_webBrowser) {
            BSTR location = nullptr;
            if (SUCCEEDED(m_webBrowser->get_LocationURL(&location)) && location) {
                std::wstring url(location, SysStringLen(location));
                SysFreeString(location);
                if (!url.empty()) {
                    current = ParseExplorerUrl(url);
                }
            }
            if (!current) {
                BSTR name = nullptr;
                if (SUCCEEDED(m_webBrowser->get_LocationName(&name)) && name) {
                    std::wstring locationName(name, SysStringLen(name));
                    SysFreeString(name);
                    if (!locationName.empty()) {
                        current = ParseDisplayName(locationName);
                    }
                }
            }
        }
        if (!current) {
            return;
        }
    }

    if (m_pendingWindowRedirect) {
        m_pendingWindowRedirect = false;
        if (TryRedirectToExistingWindow(current.get())) {
            return;
        }
    }

    std::wstring name = GetDisplayName(current.get());
    if (name.empty()) {
        name = L"Tab";
    }
    const auto computeNormalizedPath = [](PCIDLIST_ABSOLUTE value) {
        std::wstring canonical = GetCanonicalParsingName(value);
        if (!canonical.empty()) {
            return canonical;
        }
        return GetParsingName(value);
    };
    const std::wstring parsingName = computeNormalizedPath(current.get());

    bool isInitialNav = m_isInitialNavigation;
    m_isInitialNavigation = false;

    if (isInitialNav && current) {
        m_initialNavigationPidl = ClonePidl(current.get());
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    if (selected.IsValid()) {
        if (isInitialNav) {
            TabLocation existing = m_tabs.Find(current.get());
            if (existing.IsValid()) {
                if (existing != selected) {
                    m_tabs.SetSelectedLocation(existing);
                    SyncSavedGroup(existing.groupIndex);
                }
                return;
            } else {
                TabLocation location = m_tabs.Add(ClonePidl(current.get()), name, name, true, selected.groupIndex);
                if (auto* tab = m_tabs.Get(location)) {
                    tab->path = !parsingName.empty() ? parsingName : computeNormalizedPath(tab->pidl.get());
                    tab->RefreshNormalizedLookupKey();
                    tab->tooltip = tab->path;
                    NavigationHistoryEntry entry;
                    entry.pidl = ClonePidl(current.get());
                    entry.path = tab->path;
                    entry.name = name;
                    entry.timestamp = GetTickCount64();
                    tab->navigationHistory.entries.push_back(std::move(entry));
                    tab->navigationHistory.currentIndex = 0;
                }
                UpdateTabsUI();
                SyncAllSavedGroups();
                return;
            }
        }

        if (auto* tab = m_tabs.Get(selected)) {
            const std::wstring oldKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            const std::wstring oldPath = tab->path;

            tab->pidl = ClonePidl(current.get());
            tab->name = name;
            tab->tooltip = name;
            tab->hidden = false;
            tab->path = !parsingName.empty() ? parsingName : computeNormalizedPath(tab->pidl.get());
            tab->RefreshNormalizedLookupKey();
            const std::wstring newKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            if (!oldKey.empty() && oldKey != newKey) {
                IconCache::Instance().InvalidateFamily(oldKey);
            }

            if (tab->navigationHistory.entries.empty()) {
                // History was not initialized (shouldn't happen now that we initialize on tab creation)
                // Record current location as a safety fallback
                m_tabs.RecordNavigation(selected, ClonePidl(current.get()), tab->path, tab->name);
            } else if (!m_internalNavigation && oldPath != tab->path) {
                // Before recording as new navigation, check if this matches an adjacent
                // history entry (indicating an implicit back/forward from Explorer native navigation)
                if (!m_tabs.TryMatchImplicitNavigation(selected, tab->path)) {
                    m_tabs.RecordNavigation(selected, ClonePidl(current.get()), tab->path, tab->name);
                }
            }

            m_tabs.SetGroupCollapsed(selected.groupIndex, false);
            SyncSavedGroup(selected.groupIndex);
            return;
        }
    }

    TabLocation existing = m_tabs.Find(current.get());
    if (existing.IsValid()) {
        if (auto* tab = m_tabs.Get(existing)) {
            const std::wstring oldKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            tab->hidden = false;
            tab->name = name;
            tab->tooltip = name;
            tab->path = !parsingName.empty() ? parsingName : computeNormalizedPath(tab->pidl.get());
            tab->RefreshNormalizedLookupKey();
            const std::wstring newKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            if (!oldKey.empty() && oldKey != newKey) {
                IconCache::Instance().InvalidateFamily(oldKey);
            }
            if (tab->navigationHistory.entries.empty()) {
                m_tabs.RecordNavigation(existing, ClonePidl(current.get()), tab->path, tab->name);
            }
        }
        m_tabs.SetGroupCollapsed(existing.groupIndex, false);
        m_tabs.SetSelectedLocation(existing);
        SyncSavedGroup(existing.groupIndex);
        return;
    }

    UniquePidl clone = ClonePidl(current.get());
    if (!clone) {
        return;
    }

    const bool shouldHideIndicator = (m_tabs.TotalTabCount() == 0);
    TabLocation location = m_tabs.Add(std::move(clone), name, name, true);
    if (!location.IsValid()) {
        return;
    }

    // Initialize navigation history with the starting location
    // This ensures back/forward navigation works correctly from the first navigation
    if (auto* newTab = m_tabs.Get(location)) {
        m_tabs.RecordNavigation(location, ClonePidl(newTab->pidl.get()), newTab->path, newTab->name);
    }

    if (shouldHideIndicator) {
        m_tabs.SetGroupHeaderVisible(location.groupIndex, false);
    }

    SyncSavedGroup(location.groupIndex);
}

void TabBand::OpenTabInNewWindow(const TabInfo& tab) {
    if (!m_shellBrowser || !tab.pidl) {
        return;
    }
    ++m_allowExternalNewWindows;
    const HRESULT hr = m_shellBrowser->BrowseObject(tab.pidl.get(), SBSP_NEWBROWSER | SBSP_EXPLORE | SBSP_ABSOLUTE);
    if (FAILED(hr) && m_allowExternalNewWindows > 0) {
        --m_allowExternalNewWindows;
    }
}

std::vector<std::pair<TabLocation, std::wstring>> TabBand::GetHiddenTabs(int groupIndex) const {
    return m_tabs.GetHiddenTabs(groupIndex);
}

int TabBand::GetGroupCount() const noexcept {
    return m_tabs.GroupCount();
}

bool TabBand::IsGroupHeaderVisible(int groupIndex) const {
    return m_tabs.IsGroupHeaderVisible(groupIndex);
}

bool TabBand::BuildExplorerContextMenu(TabLocation location, HMENU menu, UINT idFirst, UINT idLast,
                                       Microsoft::WRL::ComPtr<IContextMenu>* menuOut,
                                       Microsoft::WRL::ComPtr<IContextMenu2>* menu2Out,
                                       Microsoft::WRL::ComPtr<IContextMenu3>* menu3Out,
                                       UINT* usedLast) const {
    if (!menu || !menuOut || !location.IsValid()) {
        return false;
    }

    const auto* tab = m_tabs.Get(location);
    if (!tab || !tab->pidl) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellFolder> parentFolder;
    PCUITEMID_CHILD child = nullptr;
    if (FAILED(SHBindToParent(tab->pidl.get(), IID_PPV_ARGS(&parentFolder), &child))) {
        return false;
    }

    Microsoft::WRL::ComPtr<IContextMenu> contextMenu;
    HWND hwnd = m_window ? m_window->GetHwnd() : nullptr;
    if (FAILED(parentFolder->GetUIObjectOf(
            hwnd, 1, &child, IID_IContextMenu, nullptr,
            reinterpret_cast<void**>(contextMenu.ReleaseAndGetAddressOf())))) {
        return false;
    }

    HRESULT hr = contextMenu->QueryContextMenu(menu, 0, idFirst, idLast, CMF_EXPLORE | CMF_NORMAL);
    if (FAILED(hr)) {
        return false;
    }

    const UINT inserted = static_cast<UINT>(HRESULT_CODE(hr));
    if (inserted == 0) {
        return false;
    }

    if (menu3Out) {
        contextMenu.As(menu3Out);
    }
    if (menu2Out) {
        if (menu3Out && menu3Out->Get()) {
            menu3Out->As(menu2Out);
        } else {
            contextMenu.As(menu2Out);
        }
    }

    *menuOut = std::move(contextMenu);
    if (usedLast) {
        *usedLast = idFirst + inserted - 1;
    }
    return true;
}

bool TabBand::InvokeExplorerContextCommand(TabLocation location, IContextMenu* menu, UINT commandId,
                                           UINT idFirst, const POINT& ptInvoke) const {
    if (!menu || !location.IsValid() || commandId < idFirst) {
        return false;
    }

    const auto* tab = m_tabs.Get(location);
    if (!tab) {
        return false;
    }

    const UINT verb = commandId - idFirst;
    std::wstring directory = GetTabPath(location);

    CMINVOKECOMMANDINFOEX info{};
    info.cbSize = sizeof(info);
    info.fMask = CMIC_MASK_UNICODE | CMIC_MASK_PTINVOKE;
    if (!directory.empty() && IsLikelyFileSystemPath(directory)) {
        info.lpDirectoryW = directory.c_str();
    }
    info.hwnd = m_window ? m_window->GetHwnd() : nullptr;
    info.lpVerb = MAKEINTRESOURCEA(verb);
    info.lpVerbW = MAKEINTRESOURCEW(verb);
    info.nShow = SW_SHOWNORMAL;
    info.ptInvoke = ptInvoke;

    return SUCCEEDED(menu->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&info)));
}

bool TabBand::HandleNewWindowRequest(const std::wstring& targetUrl) {
    if (m_allowExternalNewWindows > 0) {
        --m_allowExternalNewWindows;
        return false;
    }

    std::vector<UniquePidl> targets;
    if (!targetUrl.empty()) {
        if (auto pidl = ParseExplorerUrl(targetUrl)) {
            targets.emplace_back(std::move(pidl));
        }
    }

    if (targets.empty()) {
        auto selection = GetSelectedItemsPidL(m_shellBrowser);
        if (!selection.empty()) {
            targets = std::move(selection);
        }
    }

    if (targets.empty()) {
        if (auto pidl = QueryCurrentFolder()) {
            targets.emplace_back(std::move(pidl));
        }
    }

    if (targets.empty()) {
        return false;
    }

    bool opened = false;
    TabLocation navigateTo;
    bool haveNavigateTarget = false;

    for (auto& pidl : targets) {
        if (!pidl) {
            continue;
        }

        std::wstring name = GetDisplayName(pidl.get());
        if (name.empty()) {
            name = L"Tab";
        }
        std::wstring tooltip = GetParsingName(pidl.get());
        if (tooltip.empty()) {
            tooltip = name;
        }

        const bool selectCurrent = !haveNavigateTarget;
        TabLocation location = m_tabs.Add(std::move(pidl), name, tooltip, selectCurrent, -1);
        if (location.IsValid()) {
            opened = true;
            if (selectCurrent && !haveNavigateTarget) {
                navigateTo = location;
                haveNavigateTarget = true;
            }
        }
    }

    if (!opened) {
        return false;
    }

    UpdateTabsUI();
    SyncAllSavedGroups();
    if (haveNavigateTarget) {
        QueueNavigateTo(navigateTo);
    }
    return true;
}

bool TabBand::LaunchShellExecute(const std::wstring& application, const std::wstring& parameters,
                                 const std::wstring& workingDirectory) const {
    const wchar_t* params = parameters.empty() ? nullptr : parameters.c_str();
    const wchar_t* directory = workingDirectory.empty() ? nullptr : workingDirectory.c_str();
    HINSTANCE result = ShellExecuteW(nullptr, L"open", application.c_str(), params, directory, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(result) > 32;
}

std::wstring TabBand::GetTabPath(TabLocation location) const {
    const auto* tab = m_tabs.Get(location);
    if (!tab) {
        return {};
    }
    if (!tab->path.empty()) {
        return tab->path;
    }
    return GetParsingName(tab->pidl.get());
}

void TabBand::PerformFileOperation(TabLocation location, const std::vector<std::wstring>& paths, bool move) {
    if (!m_shellBrowser) {
        return;
    }
    const std::wstring destinationPath = GetTabPath(location);
    if (destinationPath.empty() || !IsLikelyFileSystemPath(destinationPath)) {
        return;
    }

    Microsoft::WRL::ComPtr<IFileOperation> operation;
    if (FAILED(CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&operation)))) {
        return;
    }

    DWORD flags = FOF_NOERRORUI | FOFX_SHOWELEVATIONPROMPT | FOF_NOCONFIRMATION;
    operation->SetOperationFlags(flags);

    Microsoft::WRL::ComPtr<IShellItem> destinationItem;
    if (FAILED(SHCreateItemFromParsingName(destinationPath.c_str(), nullptr, IID_PPV_ARGS(&destinationItem)))) {
        return;
    }

    for (const auto& sourcePath : paths) {
        Microsoft::WRL::ComPtr<IShellItem> sourceItem;
        if (FAILED(SHCreateItemFromParsingName(sourcePath.c_str(), nullptr, IID_PPV_ARGS(&sourceItem)))) {
            continue;
        }
        if (move) {
            operation->MoveItem(sourceItem.Get(), destinationItem.Get(), nullptr, nullptr);
        } else {
            operation->CopyItem(sourceItem.Get(), destinationItem.Get(), nullptr, nullptr);
        }
    }

    operation->PerformOperations();
}

std::optional<std::vector<std::wstring>> TabBand::GetSavedGroupNames() const {
    auto& store = GroupStore::Instance();
    if (!LoadGroupStoreForContext(L"TabBand::GetSavedGroupNames failed to load saved groups", store)) {
        return std::nullopt;
    }
    return store.GroupNames();
}

std::wstring TabBand::GetSavedGroupId(int groupIndex) const {
    const TabGroup* group = m_tabs.GetGroup(groupIndex);
    if (!group) {
        return {};
    }
    return group->savedGroupId;
}

void TabBand::OnCreateSavedGroup(int afterGroup) {
    HWND hwnd = m_window ? m_window->GetHwnd() : nullptr;
    std::wstring name = L"New Group";
    COLORREF color = RGB(0, 120, 215);
    if (!PromptForTextInput(hwnd, L"Create Tab Group", L"Group name:", &name, &color)) {
        return;
    }
    if (name.empty()) {
        return;
    }

    auto& store = GroupStore::Instance();
    if (!LoadGroupStoreForContext(L"TabBand::OnCreateSavedGroup failed to load saved groups", store)) {
        return;
    }
    if (store.Find(name)) {
        MessageBoxW(hwnd, L"A saved group with that name already exists.", L"ShellTabs", MB_OK | MB_ICONWARNING);
        return;
    }

    SavedGroup saved;
    saved.name = name;
    saved.color = color;
    saved.outlineStyle = TabGroupOutlineStyle::kSolid;
    store.Upsert(saved);

    const int groupIndex = m_tabs.CreateGroupAfter(afterGroup, name, true);
    if (auto* group = m_tabs.GetGroup(groupIndex)) {
        group->savedGroupId = name;
        group->hasCustomOutline = true;
        group->outlineColor = color;
        group->outlineStyle = TabGroupOutlineStyle::kSolid;
        group->headerVisible = true;
        group->collapsed = false;
    }
    UpdateTabsUI();
    SyncSavedGroup(groupIndex);
}

void TabBand::OnLoadSavedGroup(const std::wstring& name, int afterGroup) {
    auto& store = GroupStore::Instance();
    if (!LoadGroupStoreForContext(L"TabBand::OnLoadSavedGroup failed to load saved groups", store)) {
        return;
    }
    const SavedGroup* saved = store.Find(name);
    if (!saved) {
        return;
    }

    const int groupIndex = m_tabs.CreateGroupAfter(afterGroup, saved->name, true);
    auto* group = m_tabs.GetGroup(groupIndex);
    if (!group) {
        return;
    }
    group->savedGroupId = saved->name;
    group->hasCustomOutline = true;
    group->outlineColor = saved->color;
    group->outlineStyle = saved->outlineStyle;
    group->headerVisible = true;
    group->collapsed = false;

    bool selectFirst = true;
    bool addedAny = false;
    for (const auto& path : saved->tabPaths) {
        UniquePidl pidl = ParseDisplayName(path);
        if (!pidl) {
            continue;
        }
        std::wstring tabName = GetDisplayName(pidl.get());
        if (tabName.empty()) {
            tabName = path;
        }
        TabLocation location = m_tabs.Add(std::move(pidl), tabName, tabName, selectFirst, groupIndex);
        if (selectFirst) {
            selectFirst = false;
        }
        if (location.IsValid()) {
            addedAny = true;
        }
    }

    UpdateTabsUI();
    SyncSavedGroup(groupIndex);

    if (addedAny) {
        TabLocation selection = m_tabs.SelectedLocation();
        if (!selection.IsValid() || selection.groupIndex != groupIndex) {
            selection = {groupIndex, 0};
        }
        if (selection.IsValid()) {
            NavigateToTab(selection);
        }
    }
}

void TabBand::OnShowOptionsDialog(OptionsDialogPage initialPage, const std::wstring& focusGroupId,
                                  bool editFocusedGroup) {
    EnsureOptionsLoaded();
    ShellTabsOptions previousOptions = m_options;

    HWND owner = nullptr;
    if (m_window) {
        owner = m_window->GetHwnd();
        if (owner) {
            owner = GetAncestor(owner, GA_ROOT);
        }
    }

    OptionsDialogResult dialog =
        ShowOptionsDialog(owner, initialPage, focusGroupId.empty() ? nullptr : focusGroupId.c_str(),
                          editFocusedGroup);
    if (!dialog.saved) {
        return;
    }

    m_optionsLoaded = false;
    EnsureOptionsLoaded();
    if (dialog.optionsChanged) {
        ApplyOptionsChanges(previousOptions);
    } else {
        const UINT message = GetOptionsChangedMessage();
        if (message != 0) {
            SendNotifyMessageW(HWND_BROADCAST, message, 0, 0);
        }
    }
    if (m_window) {
        m_window->RefreshTheme();
        if (HWND hwnd = m_window->GetHwnd()) {
            InvalidateRect(hwnd, nullptr, TRUE);
        }
    }
    if (dialog.groupsChanged) {
        auto& store = GroupStore::Instance();
        LoadGroupStoreForContext(L"TabBand::OnShowOptionsDialog failed to load saved groups", store);

        for (const auto& removedId : dialog.removedGroupIds) {
            store.Remove(removedId);
        }

        for (const auto& rename : dialog.renamedGroups) {
            if (_wcsicmp(rename.first.c_str(), rename.second.c_str()) != 0) {
                store.Remove(rename.first);
            }
        }

        for (const auto& savedGroup : dialog.savedGroups) {
            store.Upsert(savedGroup);
        }

        store.RecordChanges(dialog.renamedGroups, dialog.removedGroupIds);
        const std::vector<SavedGroup> updatedGroups = store.Groups();
        const bool metadataUpdated =
            ApplySavedGroupMetadata(updatedGroups, dialog.renamedGroups, dialog.removedGroupIds);

        m_skipSavedGroupSync = true;
        SyncAllSavedGroups();
        if (metadataUpdated) {
            UpdateTabsUI();
            SaveSession();
        }
        m_processedGroupStoreGeneration = store.ChangeGeneration();
    }
}

void TabBand::OnDeferredNavigate() {
    m_deferredNavigationPosted = false;
    TabLocation target = m_pendingNavigation;
    m_pendingNavigation = {};
    if (target.IsValid()) {
        NavigateToTab(target);
    }
}

void TabBand::OnDockingModeChanged(TabBandDockMode mode) {
    if (mode == TabBandDockMode::kAutomatic) {
        return;
    }
    if (mode == m_dockMode) {
        return;
    }
    m_dockMode = mode;
    m_requestedDockMode = mode;
    SaveSession();
}

void TabBand::QueueNavigateTo(TabLocation location) {
    if (!location.IsValid() || !m_window) {
        return;
    }
    m_pendingNavigation = location;
    if (m_deferredNavigationPosted) {
        return;
    }
    HWND hwnd = m_window->GetHwnd();
    if (!hwnd) {
        return;
    }
    if (PostMessageW(hwnd, WM_SHELLTABS_DEFER_NAVIGATE, 0, 0)) {
        m_deferredNavigationPosted = true;
    }
}

bool TabBand::ApplySavedGroupMetadata(const std::vector<SavedGroup>& savedGroups,
                                      const std::vector<std::pair<std::wstring, std::wstring>>& renamedGroups,
                                      const std::vector<std::wstring>& removedGroupIds) {
    const auto caseEquals = [](const std::wstring& left, const std::wstring& right) {
        return _wcsicmp(left.c_str(), right.c_str()) == 0;
    };

    bool changed = false;
    const int groupCount = m_tabs.GroupCount();
    for (int i = 0; i < groupCount; ++i) {
        TabGroup* group = m_tabs.GetGroup(i);
        if (!group) {
            continue;
        }

        bool removed = false;
        for (const auto& removedId : removedGroupIds) {
            if (caseEquals(group->savedGroupId, removedId)) {
                if (!group->savedGroupId.empty()) {
                    group->savedGroupId.clear();
                    changed = true;
                }
                removed = true;
                break;
            }
        }
        if (removed) {
            continue;
        }

        for (const auto& rename : renamedGroups) {
            if (caseEquals(group->savedGroupId, rename.first)) {
                if (!caseEquals(group->savedGroupId, rename.second)) {
                    group->savedGroupId = rename.second;
                    changed = true;
                }
                if (group->name.empty() || caseEquals(group->name, rename.first)) {
                    if (!caseEquals(group->name, rename.second)) {
                        group->name = rename.second;
                        changed = true;
                    }
                }
                break;
            }
        }

        if (group->savedGroupId.empty()) {
            continue;
        }

        const SavedGroup* savedMatch = nullptr;
        for (const auto& saved : savedGroups) {
            if (caseEquals(saved.name, group->savedGroupId)) {
                savedMatch = &saved;
                break;
            }
        }
        if (!savedMatch) {
            continue;
        }

        if (!group->hasCustomOutline || group->outlineColor != savedMatch->color) {
            group->hasCustomOutline = true;
            group->outlineColor = savedMatch->color;
            changed = true;
        }
        if (group->outlineStyle != savedMatch->outlineStyle) {
            group->outlineStyle = savedMatch->outlineStyle;
            changed = true;
        }
        if (group->name.empty() || caseEquals(group->name, savedMatch->name)) {
            if (!caseEquals(group->name, savedMatch->name)) {
                group->name = savedMatch->name;
                changed = true;
            }
        }
    }

    return changed;
}

void TabBand::OnSavedGroupsChanged() {
    auto& store = GroupStore::Instance();
    if (!LoadGroupStoreForContext(L"TabBand::OnSavedGroupsChanged failed to load saved groups", store)) {
        return;
    }

    const uint64_t generation = store.ChangeGeneration();
    if (generation != 0 && generation == m_processedGroupStoreGeneration) {
        return;
    }

    const auto& savedGroups = store.Groups();
    const auto& renamedGroups = store.LastRenamedGroups();
    const auto& removedGroupIds = store.LastRemovedGroups();
    const bool updated = ApplySavedGroupMetadata(savedGroups, renamedGroups, removedGroupIds);
    if (updated) {
        UpdateTabsUI();
        SaveSession();
    }
    m_processedGroupStoreGeneration = generation;
}

void TabBand::SyncSavedGroup(int groupIndex) const {
    // Saved groups (those with a non-empty savedGroupId) are intentionally promoted
    // to durable storage by the user; their current tab set should always be
    // mirrored back to groups.db so that re-loading the saved group reproduces
    // the state the user last left it in. Unsaved islands are skipped entirely
    // because they have no savedGroupId.
    const auto* group = m_tabs.GetGroup(groupIndex);
    if (!group || group->savedGroupId.empty()) {
        return;
    }
    std::vector<std::wstring> paths;
    paths.reserve(group->tabs.size());
    for (const auto& tab : group->tabs) {
        if (!tab.path.empty()) {
            paths.push_back(tab.path);
        }
    }
    if (!GroupStore::Instance().UpdateTabs(group->savedGroupId, paths)) {
        SavedGroup saved;
        saved.name = group->savedGroupId;
        saved.color = group->hasCustomOutline ? group->outlineColor : RGB(0, 120, 215);
        saved.tabPaths = std::move(paths);
        saved.outlineStyle = group->outlineStyle;
        GroupStore::Instance().Upsert(std::move(saved));
    }
}

void TabBand::SyncAllSavedGroups() const {
    if (m_skipSavedGroupSync) {
        m_skipSavedGroupSync = false;
        return;
    }
    const int groupCount = m_tabs.GroupCount();
    for (int i = 0; i < groupCount; ++i) {
        SyncSavedGroup(i);
    }
}

void TabBand::CancelInitializationWorker() {
    if (m_initializationThread.joinable()) {
        m_initializationThread.request_stop();
        m_initializationThread.join();
        m_initializationThread = std::jthread();
    }
    m_backgroundInitializationActive = false;
}

void TabBand::RunBackgroundInitialization(std::stop_token stopToken, uint64_t sequence, bool hasPendingSeed) {
    LogMessage(LogLevel::Info, L"TabBand::RunBackgroundInitialization begin (this=%p, seq=%llu)", this,
               static_cast<unsigned long long>(sequence));

    auto result = std::make_unique<InitializationResult>();
    result->sequence = sequence;

    if (stopToken.stop_requested()) {
        return;
    }

    auto& groupStore = GroupStore::Instance();
    result->groupStoreLoaded =
        LoadGroupStoreForContext(L"TabBand::InitializeTabs async failed to load saved groups", groupStore);

    if (stopToken.stop_requested()) {
        return;
    }

    ShellTabsOptions optionsSnapshot = m_options;
    auto& optionsStore = OptionsStore::Instance();
    if (LoadOptionsStoreForContext(L"TabBand::InitializeTabs async failed to load options", optionsStore)) {
        optionsSnapshot = optionsStore.Get();
        result->optionsLoaded = true;
    }
    result->options = optionsSnapshot;

    if (stopToken.stop_requested()) {
        return;
    }

    // Session restore: coordinator handles crash markers and file loading.
    if (m_sessionStore && result->groupStoreLoaded) {
        bool wasCrash = SessionCoordinator::Instance().WasCrash();
        bool reopenOnCrash = result->optionsLoaded ? optionsSnapshot.reopenOnCrash : m_options.reopenOnCrash;
        bool shouldRestore = !hasPendingSeed;
        if (wasCrash && !reopenOnCrash) {
            shouldRestore = false;
        }
        result->shouldRestoreSession = shouldRestore;
        if (shouldRestore && !stopToken.stop_requested()) {
            if (wasCrash) {
                // After a crash, defer the claim to a timer so that transient
                // "probe" windows (which Explorer creates and closes within
                // ~300ms) never touch the pending pool.  The timer in
                // HandleInitializationResult will claim after the window has
                // proven it is long-lived.
                LogMessage(LogLevel::Info,
                           L"TabBand::RunBackgroundInitialization deferring crash session claim to timer");
            } else {
                SessionData data;
                std::vector<std::wstring> crashPaths;
                int crashSelectedIndex = -1;
                bool fastRestored = false;

                if (m_sessionStore && CrashRecoveryLogger::Instance().ReadCrashedState(m_sessionStore->Slot(), crashPaths, crashSelectedIndex)) {
                    SessionGroup group;
                    group.name = L"Recovered";
                    for (const auto& path : crashPaths) {
                        SessionTab tab;
                        tab.path = path;
                        group.tabs.push_back(tab);
                    }
                    data.groups.push_back(group);
                    data.selectedGroup = 0;
                    data.selectedTab = crashSelectedIndex >= 0 ? crashSelectedIndex : 0;
                    
                    result->sessionData = std::move(data);
                    result->hasSessionData = true;
                    fastRestored = true;
                    LogMessage(LogLevel::Info, L"TabBand::RunBackgroundInitialization fast crash recovery successful");
                }

                if (!fastRestored) {
                    if (m_sessionStore->Load(data)) {
                        result->sessionData = std::move(data);
                        result->hasSessionData = true;
                    } else {
                        LogMessage(LogLevel::Warning,
                                   L"TabBand::RunBackgroundInitialization no session data to claim");
                    }
                }
            }
        }
    }

    if (stopToken.stop_requested()) {
        return;
    }

    PostInitializationResult(std::move(result));
}

void TabBand::PostInitializationResult(std::unique_ptr<InitializationResult> result) {
    if (!result) {
        return;
    }
    HWND hwnd = m_window ? m_window->GetHwnd() : nullptr;
    if (!hwnd) {
        return;
    }
    auto* payload = result.release();
    if (!PostMessageW(hwnd, WM_SHELLTABS_INITIALIZATION_COMPLETE, 0,
                      reinterpret_cast<LPARAM>(payload))) {
        delete payload;
    }
}

void TabBand::HandleInitializationResult(std::unique_ptr<InitializationResult> result) {
    if (!result) {
        return;
    }
    if (result->sequence != m_initializationSequence) {
        return;
    }

    LogMessage(LogLevel::Info, L"TabBand::HandleInitializationResult applying sequence %llu",
               static_cast<unsigned long long>(result->sequence));

    m_backgroundInitializationActive = false;

    const bool hadOptions = m_optionsLoaded;
    ShellTabsOptions previousOptions = m_options;
    if (result->optionsLoaded) {
        m_options = result->options;
        m_optionsLoaded = true;
        if (hadOptions && previousOptions != m_options) {
            ApplyOptionsChanges(previousOptions);
        }
    }

    if (m_optionsLoaded && m_requestedDockMode == TabBandDockMode::kAutomatic) {
        m_requestedDockMode = m_options.tabDockMode;
    }
    if (m_window) {
        TabBandDockMode preferred = m_requestedDockMode;
        if (preferred == TabBandDockMode::kAutomatic && m_optionsLoaded) {
            preferred = m_options.tabDockMode;
        }
        if (preferred == TabBandDockMode::kAutomatic) {
            preferred = TabBandDockMode::kTop;
        }
        m_window->SetPreferredDockMode(preferred);
    }

    m_lastSessionUnclean = SessionCoordinator::Instance().WasCrash();

    m_tabs.Clear();

    bool restored = false;
    if (result->shouldRestoreSession && result->hasSessionData) {
        restored = RestoreSessionFromData(result->sessionData);
    }

    // Crash state has been handled (restored or not). Clear the flag so
    // subsequent windows aren't treated as post-crash and so the crash
    // marker can eventually be cleaned when the last window unregisters.
    SessionCoordinator::Instance().ClearCrashState();

    // If we expected to restore but couldn't claim data (cooldown suppressed),
    // and pending data still exists, schedule a retry after the cooldown expires.
    if (!restored && result->shouldRestoreSession && !result->hasSessionData &&
        SessionCoordinator::Instance().HasPendingData() && m_window) {
        HWND hwnd = m_window->GetHwnd();
        if (hwnd) {
            LogMessage(LogLevel::Info,
                       L"TabBand::HandleInitializationResult scheduling session retry in 4s");
            SetTimer(hwnd, TabBandWindow::SessionRetryTimerId(), 4000, nullptr);
        }
    }

    bool handledPendingSeed = false;
    if (!restored && m_pendingGroupSeed && !m_pendingGroupSeed->tabs.empty()) {
        std::vector<TabGroup> groups;
        groups.emplace_back(std::move(*m_pendingGroupSeed));
        m_tabs.Restore(std::move(groups), 0, 0, m_tabs.NextGroupSequence());
        handledPendingSeed = true;
    }
    m_pendingGroupSeed.reset();

    bool pendingStandalone = m_pendingStandaloneSeed;
    m_pendingStandaloneSeed = false;

    if (!handledPendingSeed) {
        if (!restored || m_tabs.TotalTabCount() == 0) {
            UniquePidl pidl = QueryCurrentFolder();
            if (pidl) {
                std::wstring name = GetDisplayName(pidl.get());
                if (name.empty()) {
                    name = L"Tab";
                }
                const bool hideHeader = !m_hadPendingSeed || pendingStandalone;
                TabLocation location = m_tabs.Add(std::move(pidl), name, name, true);
                if (location.IsValid() && hideHeader) {
                    m_tabs.SetGroupHeaderVisible(location.groupIndex, false);
                }
                LogMessage(LogLevel::Info, L"TabBand::InitializeTabs hydrated tab %ls", name.c_str());
            }
        } else {
            if (m_initialNavigationPidl) {
                TabLocation existing = m_tabs.Find(m_initialNavigationPidl.get());
                if (existing.IsValid()) {
                    m_tabs.SetSelectedLocation(existing);
                    SyncSavedGroup(existing.groupIndex);
                } else {
                    std::wstring name = GetDisplayName(m_initialNavigationPidl.get());
                    if (name.empty()) {
                        name = L"Tab";
                    }
                    const TabLocation selection = m_tabs.SelectedLocation();
                    int targetGroup = selection.IsValid() ? selection.groupIndex : 0;
                    TabLocation location = m_tabs.Add(std::move(m_initialNavigationPidl), name, name, true, targetGroup);
                    if (location.IsValid()) {
                        auto* tab = m_tabs.Get(location);
                        if (tab) {
                            std::wstring canonical = GetCanonicalParsingName(tab->pidl.get());
                            if (canonical.empty()) {
                                canonical = GetParsingName(tab->pidl.get());
                            }
                            tab->path = canonical;
                            tab->RefreshNormalizedLookupKey();
                            tab->tooltip = tab->path;
                            NavigationHistoryEntry entry;
                            entry.pidl = ClonePidl(tab->pidl.get());
                            entry.path = tab->path;
                            entry.name = name;
                            entry.timestamp = GetTickCount64();
                            tab->navigationHistory.entries.push_back(std::move(entry));
                            tab->navigationHistory.currentIndex = 0;
                        }
                        m_tabs.SetSelectedLocation(location);
                        SyncSavedGroup(location.groupIndex);
                    }
                }
                m_initialNavigationPidl.reset();
            } else {
                const TabLocation selection = m_tabs.SelectedLocation();
                if (selection.IsValid()) {
                    const bool previousRestoring = m_restoringSession;
                    m_restoringSession = true;
                    NavigateToTab(selection);
                    m_restoringSession = previousRestoring;
                }
            }
        }
    }

    if (result->groupStoreLoaded) {
        SyncAllSavedGroups();
    }

    // Session persistence is ready as soon as we have a session store.
    if (m_sessionStore) {
        m_sessionPersistenceReady = true;
        StartSessionFlushTimer();
    } else {
        m_sessionPersistenceReady = false;
    }

    m_hadPendingSeed = false;

    UpdateTabsUI();
}

void TabBand::OnRestoreTabSessionRequested(int index) {
    auto savedSessions = SavedTabSessionManager::Instance().GetSavedSessions();
    if (index >= 0 && index < static_cast<int>(savedSessions.size())) {
        const auto& session = savedSessions[index];
        for (const auto& path : session.paths) {
            OnOpenFolderInNewTab(path, false);
        }
    }
}

}  // namespace shelltabs
