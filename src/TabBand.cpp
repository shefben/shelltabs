#include "TabBand.h"

#include <algorithm>
#include <atomic>
#include <deque>
#include <memory>
#include <cstring>
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
#include "TabBandWindow.h"
#include "Utilities.h"

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
}

TabBand::TabBand() : m_refCount(1) {
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

IFACEMETHODIMP TabBand::ContextSensitiveHelp(BOOL) {
    return E_NOTIMPL;
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

IFACEMETHODIMP TabBand::ResizeBorderDW(const RECT*, IUnknown*, BOOL) {
    return E_NOTIMPL;
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

void TabBand::OnNewThisPCInGroupRequested(int groupIndex) {
	// Parse the shell namespace path for "This PC"
	UniquePidl pidl = ParseExplorerUrl(L"shell:MyComputerFolder");
	if (!pidl) {
		// Fallback for older systems
		LPITEMIDLIST raw = nullptr;
		if (SUCCEEDED(SHGetSpecialFolderLocation(nullptr, CSIDL_DRIVES, &raw)) && raw) {
			pidl.reset(raw);
		}
	}
	if (!pidl) return;

	std::wstring name = GetDisplayName(pidl.get());
	if (name.empty()) name = L"This PC";
	std::wstring tooltip = GetParsingName(pidl.get());
	if (tooltip.empty()) tooltip = name;

	if (groupIndex < 0) groupIndex = 0;

	TabLocation loc = m_tabs.Add(std::move(pidl), name, tooltip, true, groupIndex);
	UpdateTabsUI();
	SyncAllSavedGroups();
	if (loc.IsValid()) {
		NavigateToTab(loc);
	}
}

void TabBand::OnBrowserNavigate() {
    EnsureTabForCurrentFolder();
    UpdateTabsUI();
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

void TabBand::OnNewTabRequested() {
    UniquePidl pidl;
    std::wstring name;

    const TabLocation current = m_tabs.SelectedLocation();
    if (const auto* tab = m_tabs.Get(current)) {
        pidl = ClonePidl(tab->pidl.get());
        name = tab->name;
    }

    if (!pidl) {
        pidl = QueryCurrentFolder();
        name = GetDisplayName(pidl.get());
    }

    if (!pidl) {
        return;
    }

    if (name.empty()) {
        name = L"Tab";
    }

    const int targetGroup = current.groupIndex >= 0 ? current.groupIndex : 0;
    TabLocation location = m_tabs.Add(std::move(pidl), name, name, true, targetGroup);
    UpdateTabsUI();
    SyncAllSavedGroups();
    if (location.IsValid()) {
        NavigateToTab(location);
    }
}

void TabBand::OnCloseTabRequested(TabLocation location) {
    const auto selected = m_tabs.SelectedLocation();
    const bool wasSelected = (selected.groupIndex == location.groupIndex && selected.tabIndex == location.tabIndex);

    std::wstring removedGroupId;
    if (const auto* group = m_tabs.GetGroup(location.groupIndex)) {
        if (!group->savedGroupId.empty() && group->tabs.size() == 1) {
            removedGroupId = group->savedGroupId;
        }
    }

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

    TabLocation newLocation = m_tabs.Add(std::move(clone), name, tooltip, true, location.groupIndex);
    UpdateTabsUI();
    SyncAllSavedGroups();
    if (newLocation.IsValid()) {
        NavigateToTab(newLocation);
    }
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
}

void TabBand::OnEditGroupProperties(int groupIndex) {
    auto* group = m_tabs.GetGroup(groupIndex);
    if (!group) {
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
    if (path.empty()) {
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
    if (path.empty()) {
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

void TabBand::OnOpenFolderInNewTab(const std::wstring& path) {
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
    TabLocation location = m_tabs.Add(std::move(pidl), name, name, true, targetGroup);

    UpdateTabsUI();
    SyncAllSavedGroups();

    if (location.IsValid()) {
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
        LogMessage(LogLevel::Info, L"TabBand::EnsureWindow created window hwnd=%p",
                   m_window ? m_window->GetHwnd() : nullptr);
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
    if (m_sessionMarkerActive) {
        SessionStore::ClearSessionMarker();
        m_sessionMarkerActive = false;
    }
    ReleaseWindowToken();

    if (m_browserEvents) {
        m_browserEvents->Disconnect();
        m_browserEvents.reset();
    }

    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();

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
    m_tabs.Clear();

    std::shared_ptr<PendingWindowSeed> pendingSeed = DequeuePendingWindowSeed();

    GroupStore::Instance().Load();
    EnsureSessionStore();
    EnsureOptionsLoaded();
    m_lastSessionUnclean = SessionStore::WasPreviousSessionUnclean();
    SessionStore::MarkSessionActive();
    m_sessionMarkerActive = true;

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

    std::vector<TabGroup> groups;
    groups.reserve(data.groups.size());
    for (const auto& groupData : data.groups) {
        TabGroup group;
        group.name = groupData.name.empty() ? L"Island" : groupData.name;
        group.collapsed = groupData.collapsed;
        group.headerVisible = groupData.headerVisible;
        group.hasCustomOutline = groupData.hasOutline;
        group.outlineColor = groupData.outlineColor;
        group.savedGroupId = groupData.savedGroupId;
        for (const auto& tabData : groupData.tabs) {
            UniquePidl pidl = ParseDisplayName(tabData.path);
            if (!pidl) {
                continue;
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
            tab.path = tabData.path;
            group.tabs.emplace_back(std::move(tab));
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

    SessionData data;
    const TabLocation selected = m_tabs.SelectedLocation();
    data.selectedGroup = selected.groupIndex;
    data.selectedTab = selected.tabIndex;
    data.groupSequence = m_tabs.NextGroupSequence();

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
        storedGroup.savedGroupId = group->savedGroupId;

        for (const auto& tab : group->tabs) {
            SessionTab storedTab;
            storedTab.path = tab.path;
            if (storedTab.path.empty()) {
                storedTab.path = GetParsingName(tab.pidl.get());
            }
            if (storedTab.path.empty()) {
                continue;
            }
            storedTab.name = tab.name;
            storedTab.tooltip = tab.tooltip;
            storedTab.hidden = tab.hidden;
            storedGroup.tabs.emplace_back(std::move(storedTab));
        }

        if (!storedGroup.tabs.empty() || !storedGroup.savedGroupId.empty()) {
            data.groups.emplace_back(std::move(storedGroup));
        }
    }

    if (data.groups.empty()) {
        return;
    }

    m_sessionStore->Save(data);
}

void TabBand::ApplyOptionsChanges(const ShellTabsOptions& previousOptions) {
    if (!previousOptions.persistGroupPaths && m_options.persistGroupPaths) {
        SyncAllSavedGroups();
    }
}

UniquePidl TabBand::QueryCurrentFolder() const {
    return GetCurrentFolderPidL(m_shellBrowser, m_webBrowser);
}

void TabBand::NavigateToTab(TabLocation location) {
    if (!m_shellBrowser) {
        return;
    }
    auto* tab = m_tabs.Get(location);
    if (!tab || !tab->pidl) {
        return;
    }

    m_tabs.SetGroupCollapsed(location.groupIndex, false);
    m_tabs.SetSelectedLocation(location);
    SaveSession();
    m_internalNavigation = true;
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
    const std::wstring parsingName = GetParsingName(current.get());

    const TabLocation selected = m_tabs.SelectedLocation();
    if (selected.IsValid()) {
        if (auto* tab = m_tabs.Get(selected)) {
            tab->pidl = ClonePidl(current.get());
            tab->name = name;
            tab->tooltip = name;
            tab->hidden = false;
            tab->path = !parsingName.empty() ? parsingName : GetParsingName(tab->pidl.get());
            m_tabs.SetGroupCollapsed(selected.groupIndex, false);
            SyncSavedGroup(selected.groupIndex);
            return;
        }
    }

    TabLocation existing = m_tabs.Find(current.get());
    if (existing.IsValid()) {
        if (auto* tab = m_tabs.Get(existing)) {
            tab->hidden = false;
            tab->name = name;
            tab->tooltip = name;
            tab->path = !parsingName.empty() ? parsingName : GetParsingName(tab->pidl.get());
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
    TabLocation location = m_tabs.Add(std::move(clone), name, name, true);
    if (location.IsValid()) {
        SyncSavedGroup(location.groupIndex);
    }
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
    if (!directory.empty()) {
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
    if (destinationPath.empty()) {
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
    store.Upsert(saved);

    const int groupIndex = m_tabs.CreateGroupAfter(afterGroup, name, true);
    if (auto* group = m_tabs.GetGroup(groupIndex)) {
        group->savedGroupId = name;
        group->hasCustomOutline = true;
        group->outlineColor = color;
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

void TabBand::OnShowOptionsDialog(int initialTab) {
    EnsureOptionsLoaded();
    ShellTabsOptions previousOptions = m_options;

    HWND owner = nullptr;
    if (m_window) {
        owner = m_window->GetHwnd();
        if (owner) {
            owner = GetAncestor(owner, GA_ROOT);
        }
    }

    OptionsDialogResult dialog = ShowOptionsDialog(owner, initialTab);
    if (!dialog.saved) {
        return;
    }

    m_optionsLoaded = false;
    EnsureOptionsLoaded();
    if (dialog.optionsChanged) {
        ApplyOptionsChanges(previousOptions);
    }
    if (dialog.groupsChanged) {
        GroupStore::Instance().Load();
        SyncAllSavedGroups();
        UpdateTabsUI();
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
        GroupStore::Instance().Upsert(std::move(saved));
    }
}

void TabBand::SyncAllSavedGroups() const {
    const int groupCount = m_tabs.GroupCount();
    for (int i = 0; i < groupCount; ++i) {
        SyncSavedGroup(i);
    }
}

}  // namespace shelltabs

