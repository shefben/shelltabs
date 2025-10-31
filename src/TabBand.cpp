#include "TabBand.h"

#include <algorithm>
#include <atomic>
#include <deque>
#include <memory>
#include <cstring>
#include <cwchar>
#include <mutex>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <objbase.h>

#include <CommCtrl.h>
#include <ShlObj.h>
#include <Shlwapi.h>
#include <shellapi.h>
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
struct WindowTokenState {
    std::mutex mutex;
    std::unordered_map<HWND, std::wstring> tokens;
};

WindowTokenState& GetWindowTokenState() {
    static auto* state = new WindowTokenState();
    return *state;
}

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
}

TabBand::TabBand() : m_refCount(1), m_processedGroupStoreGeneration(0) {
    ModuleAddRef();
    LogMessage(LogLevel::Info, L"TabBand constructed (this=%p)", this);
}

TabBand::~TabBand() {
    LogMessage(LogLevel::Info, L"TabBand destroyed (this=%p)", this);
    DisconnectSite();
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
            Microsoft::WRL::ComPtr<IDockingWindowSite> dockingSite = m_dockingSite;
            if (!dockingSite && punkToolbarSite) {
                punkToolbarSite->QueryInterface(IID_PPV_ARGS(&dockingSite));
            }
            if (!dockingSite && m_site) {
                m_site.As(&dockingSite);
            }

            if (dockingSite) {
                IUnknown* siteForCall = punkToolbarSite;
                if (!siteForCall && m_site) {
                    siteForCall = m_site.Get();
                }
                const HRESULT hr = dockingSite->ResizeBorderDW(prcBorder, siteForCall, fReserved);
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

IFACEMETHODIMP TabBand::TranslateAcceleratorIO(LPMSG) {
    return GuardExplorerCall(L"TabBand::TranslateAcceleratorIO", []() -> HRESULT { return S_FALSE; },
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
    m_internalNavigation = false;
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
    if (location.IsValid()) {
        QueueNavigateTo(location);
    }
    return true;
}

void TabBand::OnTabSelected(TabLocation location) {
    const auto current = m_tabs.SelectedLocation();
    if (current.groupIndex == location.groupIndex && current.tabIndex == location.tabIndex) {
        return;
    }
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
            store.Load();
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

void TabBand::OnReopenClosedTabRequested() {
    if (m_closedTabHistory.empty()) {
        return;
    }

    ClosedTabSet set = std::move(m_closedTabHistory.back());
    m_closedTabHistory.pop_back();

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
}

void TabBand::OnToggleGroupCollapsed(int groupIndex) {
    m_tabs.ToggleGroupCollapsed(groupIndex);
    UpdateTabsUI();
}

void TabBand::OnUnhideAllInGroup(int groupIndex) {
    m_tabs.UnhideAllInGroup(groupIndex);
    UpdateTabsUI();
}

void TabBand::OnCreateIslandAfter(int groupIndex) {
    m_tabs.CreateGroupAfter(groupIndex);
    UpdateTabsUI();
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
        GroupStore::Instance().UpdateTabs(removed->savedGroupId, {});
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

    if (!group->savedGroupId.empty() && colorChanged) {
        auto& store = GroupStore::Instance();
        store.Load();
        store.UpdateColor(group->savedGroupId, color);
    }
}

void TabBand::OnDetachGroupRequested(int groupIndex) {
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
}

void TabBand::OnMoveTabRequested(TabLocation from, TabLocation to) {
    m_tabs.MoveTab(from, to);
    UpdateTabsUI();
    SyncAllSavedGroups();
}

void TabBand::OnMoveGroupRequested(int fromGroup, int toGroup) {
    m_tabs.MoveGroup(fromGroup, toGroup);
    UpdateTabsUI();
}

void TabBand::OnMoveTabToNewGroup(TabLocation from, int insertIndex, bool headerVisible) {
    m_tabs.MoveTabToNewGroup(from, insertIndex, headerVisible);
    UpdateTabsUI();
    SyncAllSavedGroups();
}

std::optional<TabInfo> TabBand::DetachTabForTransfer(TabLocation location, bool* wasSelected,
                                                     bool ensurePlaceholderTab, bool* removedLastTab) {
    if (removedLastTab) {
        *removedLastTab = false;
    }

    if (wasSelected) {
        const TabLocation selected = m_tabs.SelectedLocation();
        *wasSelected = (selected.groupIndex == location.groupIndex && selected.tabIndex == location.tabIndex);
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
    if (select && inserted.IsValid()) {
        NavigateToTab(inserted);
    }
    return inserted;
}

std::optional<TabGroup> TabBand::DetachGroupForTransfer(int groupIndex, bool* wasSelected) {
    if (wasSelected) {
        *wasSelected = (m_tabs.SelectedLocation().groupIndex == groupIndex);
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

    return insertIndex;
}

void TabBand::OnSetGroupHeaderVisible(int groupIndex, bool visible) {
    m_tabs.SetGroupHeaderVisible(groupIndex, visible);
    UpdateTabsUI();
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

    const DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
        return;
    }

    UniquePidl pidl = ParseDisplayName(path);
    if (!pidl) {
        return;
    }

    std::wstring name = GetDisplayName(pidl.get());
    if (name.empty()) {
        name = path;
    }

    const TabLocation selected = m_tabs.SelectedLocation();
    const int targetGroup = selected.groupIndex >= 0 ? selected.groupIndex : 0;
    TabLocation location = m_tabs.Add(std::move(pidl), name, name, select, targetGroup);

    UpdateTabsUI();
    SyncAllSavedGroups();

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
    id.hwnd = GetFrameWindow();
    if (m_webBrowser) {
        id.frameCookie = reinterpret_cast<uintptr_t>(m_webBrowser.Get());
    }
    return id;
}

std::wstring TabBand::ResolveWindowToken() {
    if (!m_windowToken.empty()) {
        return m_windowToken;
    }

    HWND frame = GetFrameWindow();
    if (!frame) {
        return {};
    }

    {
        auto& state = GetWindowTokenState();
        std::scoped_lock lock(state.mutex);
        const auto existing = state.tokens.find(frame);
        if (existing != state.tokens.end()) {
            m_windowToken = existing->second;
            return m_windowToken;
        }

        GUID guid{};
        std::wstring token;
        if (SUCCEEDED(CoCreateGuid(&guid))) {
            wchar_t buffer[64];
            if (StringFromGUID2(guid, buffer, ARRAYSIZE(buffer)) > 0) {
                token.assign(buffer);
                token.erase(std::remove(token.begin(), token.end(), L'{'), token.end());
                token.erase(std::remove(token.begin(), token.end(), L'}'), token.end());
            }
        }
        if (token.empty()) {
            token = std::to_wstring(reinterpret_cast<uintptr_t>(frame));
        }
        state.tokens.emplace(frame, token);
        m_windowToken = token;
    }

    return m_windowToken;
}

void TabBand::ReleaseWindowToken() {
    HWND frame = GetFrameWindow();
    if (!frame) {
        m_windowToken.clear();
        return;
    }

    auto& state = GetWindowTokenState();
    std::scoped_lock lock(state.mutex);
    state.tokens.erase(frame);
    m_windowToken.clear();
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

    PreviewCache::Instance().StorePreviewFromWindow(tab->pidl.get(), viewWindow, kPreviewImageSize,
                                                    ResolveWindowToken());
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
        EnsureOptionsLoaded();
        TabBandDockMode preferred = m_requestedDockMode;
        if (preferred == TabBandDockMode::kAutomatic) {
            preferred = m_options.tabDockMode;
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
    store.Load();
    m_options = store.Get();
    m_optionsLoaded = true;
}

void TabBand::DisconnectSite() {
    LogMessage(LogLevel::Info, L"TabBand::DisconnectSite (this=%p)", this);
    SaveSession();
    StopSessionFlushTimer();
    if (m_sessionMarkerActive && m_sessionStore) {
        m_sessionStore->ClearSessionMarker();
    }
    m_sessionMarkerActive = false;
    m_tabs.ClearWindowId();
    for (int groupIndex = 0; groupIndex < m_tabs.GroupCount(); ++groupIndex) {
        if (const TabGroup* group = m_tabs.GetGroup(groupIndex)) {
            CancelPendingPreviewForGroup(*group);
        }
    }
    if (!m_windowToken.empty()) {
        PreviewCache::Instance().CancelPendingCapturesForOwner(m_windowToken);
    }
    ReleaseWindowToken();

    if (m_browserEvents) {
        m_browserEvents->Disconnect();
        m_browserEvents.reset();
    }

    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();
    m_siteOleWindow.Reset();
    m_dockingSite.Reset();
    m_dockingFrame.Reset();

    if (m_window) {
        m_window->SetSite(nullptr);
        m_window->Destroy();
        m_window.reset();
    }

    m_tabs.Clear();
    m_internalNavigation = false;
    m_allowExternalNewWindows = 0;
    m_sessionStore.reset();
}

void TabBand::InitializeTabs() {
    LogScope scope(L"TabBand::InitializeTabs");
    for (int groupIndex = 0; groupIndex < m_tabs.GroupCount(); ++groupIndex) {
        if (const TabGroup* group = m_tabs.GetGroup(groupIndex)) {
            CancelPendingPreviewForGroup(*group);
        }
    }
    if (!m_windowToken.empty()) {
        PreviewCache::Instance().CancelPendingCapturesForOwner(m_windowToken);
    }
    m_tabs.Clear();

    std::shared_ptr<PendingWindowSeed> pendingSeed = DequeuePendingWindowSeed();

    GroupStore::Instance().Load();
    EnsureSessionStore();
    EnsureOptionsLoaded();
    if (m_requestedDockMode == TabBandDockMode::kAutomatic) {
        m_requestedDockMode = m_options.tabDockMode;
    }
    if (m_sessionStore) {
        m_lastSessionUnclean = m_sessionStore->WasPreviousSessionUnclean();
        m_sessionStore->MarkSessionActive();
        m_sessionMarkerActive = true;
        StartSessionFlushTimer();
    } else {
        m_lastSessionUnclean = false;
    }

    bool shouldRestore = true;
    if (m_lastSessionUnclean && !m_options.reopenOnCrash) {
        shouldRestore = false;
    }
    if (pendingSeed) {
        shouldRestore = false;
    }

    bool restored = false;
    if (shouldRestore) {
        restored = RestoreSession();
    }

    bool handledPendingSeed = false;
    if (pendingSeed && pendingSeed->type == WindowSeedType::Group) {
        if (!pendingSeed->group.tabs.empty()) {
            std::vector<TabGroup> groups;
            groups.emplace_back(std::move(pendingSeed->group));
            m_tabs.Restore(std::move(groups), 0, 0, m_tabs.NextGroupSequence());
            handledPendingSeed = true;
        }
    }

    if (!handledPendingSeed) {
        if (!restored || m_tabs.TotalTabCount() == 0) {
            UniquePidl pidl = QueryCurrentFolder();
            if (pidl) {
                std::wstring name = GetDisplayName(pidl.get());
                if (name.empty()) {
                    name = L"Tab";
                }
                TabLocation location = m_tabs.Add(std::move(pidl), name, name, true);
                if (location.IsValid()) {
                    if (!pendingSeed || pendingSeed->type == WindowSeedType::StandaloneTab) {
                        m_tabs.SetGroupHeaderVisible(location.groupIndex, false);
                    }
                }
                LogMessage(LogLevel::Info, L"TabBand::InitializeTabs seeded new tab %ls", name.c_str());
            }
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

    SyncAllSavedGroups();
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

    LogMessage(LogLevel::Info, L"TabBand::EnsureSessionStore resolving storage (this=%p)", this);
    const std::wstring token = ResolveWindowToken();
    if (token.empty()) {
        LogMessage(LogLevel::Info,
                   L"TabBand::EnsureSessionStore window token unavailable; deferring initialization");
        return;
    }

    LogMessage(LogLevel::Info, L"TabBand::EnsureSessionStore window token %ls", token.c_str());
    std::wstring storagePath = SessionStore::BuildPathForToken(token);
    if (storagePath.empty()) {
        LogMessage(LogLevel::Warning,
                   L"TabBand::EnsureSessionStore failed to build storage path for token %ls",
                   token.c_str());
        return;
    }

    LogMessage(LogLevel::Info, L"TabBand::EnsureSessionStore using storage path %ls", storagePath.c_str());
    m_sessionStore = std::make_unique<SessionStore>(std::move(storagePath));
    StartSessionFlushTimer();
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

    LogMessage(LogLevel::Info, L"TabBand::RestoreSession loaded %llu groups",
               static_cast<unsigned long long>(data.groups.size()));

    m_dockMode = data.dockMode;
    if (m_dockMode == TabBandDockMode::kAutomatic) {
        EnsureOptionsLoaded();
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
        return;
    }

    EnsureSessionStore();
    if (!m_sessionStore) {
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
        return;
    }

    m_sessionStore->Save(data);
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

    constexpr UINT kSessionFlushIntervalMs = 15000;
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
    if (!tab.path.empty()) {
        return;
    }
    tab.path = GetParsingName(tab.pidl.get());
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
    if (backgroundChanged || fontChanged || backgroundTransparencyChanged || fontBrightnessChanged ||
        backgroundColorsChanged || fontColorsChanged || tabColorsChanged) {
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

void TabBand::NavigateToTab(TabLocation location) {
    if (!m_shellBrowser) {
        return;
    }
    auto* tab = m_tabs.Get(location);
    if (!tab || !tab->pidl) {
        return;
    }

    const auto current = m_tabs.SelectedLocation();
    if (current.groupIndex != location.groupIndex || current.tabIndex != location.tabIndex) {
        CaptureActiveTabPreview();
    }

    m_tabs.SetGroupCollapsed(location.groupIndex, false);
    m_tabs.SetSelectedLocation(location);
    SaveSession();
    m_internalNavigation = true;
    EnsureFtpNamespaceBinding(tab->pidl.get());
    const HRESULT hr = m_shellBrowser->BrowseObject(tab->pidl.get(), SBSP_SAMEBROWSER);
    if (FAILED(hr)) {
        m_internalNavigation = false;
    }
}

void TabBand::EnsureTabForCurrentFolder() {
    UniquePidl current = QueryCurrentFolder();
    if (!current) {
        return;
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

    const TabLocation selected = m_tabs.SelectedLocation();
    if (selected.IsValid()) {
        if (auto* tab = m_tabs.Get(selected)) {
            const std::wstring oldKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            tab->pidl = ClonePidl(current.get());
            tab->name = name;
            tab->tooltip = name;
            tab->hidden = false;
            tab->path = !parsingName.empty() ? parsingName : computeNormalizedPath(tab->pidl.get());
            const std::wstring newKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            if (!oldKey.empty() && oldKey != newKey) {
                IconCache::Instance().InvalidateFamily(oldKey);
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
            const std::wstring newKey = BuildIconCacheFamilyKey(tab->pidl.get(), tab->path);
            if (!oldKey.empty() && oldKey != newKey) {
                IconCache::Instance().InvalidateFamily(oldKey);
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

std::vector<std::wstring> TabBand::GetSavedGroupNames() const {
    auto& store = GroupStore::Instance();
    store.Load();
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
    store.Load();
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
    store.Load();
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
        store.Load();
        std::vector<SavedGroup> updatedGroups = dialog.savedGroups;
        if (updatedGroups.empty()) {
            updatedGroups = store.Groups();
        }

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
    store.Load();

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
    EnsureOptionsLoaded();
    if (!m_options.persistGroupPaths) {
        return;
    }
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

}  // namespace shelltabs

