#include "TabBand.h"

#include <algorithm>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include <CommCtrl.h>
#include <ShlObj.h>
#include <Shlwapi.h>
#include <shellapi.h>
#include <shobjidl_core.h>

#ifndef SBSP_EXPLORE
#define SBSP_EXPLORE 0x00000004
#endif

#include "Guids.h"
#include "Module.h"
#include "TabBandWindow.h"
#include "Utilities.h"

namespace shelltabs {

using Microsoft::WRL::ComPtr;

TabBand::TabBand() : m_refCount(1) {
    ModuleAddRef();
}

TabBand::~TabBand() {
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
    if (!phwnd) {
        return E_POINTER;
    }
    EnsureWindow();
    *phwnd = m_window ? m_window->GetHwnd() : nullptr;
    return *phwnd ? S_OK : E_FAIL;
}

IFACEMETHODIMP TabBand::ContextSensitiveHelp(BOOL) {
    return E_NOTIMPL;
}

IFACEMETHODIMP TabBand::ShowDW(BOOL fShow) {
    EnsureWindow();
    if (m_window) {
        m_window->Show(fShow != FALSE);
    }
    return S_OK;
}

IFACEMETHODIMP TabBand::CloseDW(DWORD) {
    if (m_window) {
        m_window->Show(false);
    }
    return S_OK;
}

IFACEMETHODIMP TabBand::ResizeBorderDW(const RECT*, IUnknown*, BOOL) {
    return E_NOTIMPL;
}

IFACEMETHODIMP TabBand::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi) {
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
        constexpr wchar_t kTitle[] = L"Shell Tabs";
        lstrcpynW(pdbi->wszTitle, kTitle, ARRAYSIZE(pdbi->wszTitle));
    }
    if (pdbi->dwMask & DBIM_MODEFLAGS) {
        pdbi->dwModeFlags = DBIMF_VARIABLEHEIGHT | DBIMF_NORMAL;
    }
    if (pdbi->dwMask & DBIM_BKCOLOR) {
        pdbi->dwMask &= ~DBIM_BKCOLOR;
    }

    return S_OK;
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
    if (fActivate) {
        EnsureWindow();
        if (m_window) {
            m_window->FocusTab();
        }
    }
    return S_OK;
}

IFACEMETHODIMP TabBand::HasFocusIO() {
    if (!m_window) {
        return S_FALSE;
    }
    return m_window->HasFocus() ? S_OK : S_FALSE;
}

IFACEMETHODIMP TabBand::TranslateAcceleratorIO(LPMSG) {
    return S_FALSE;
}

IFACEMETHODIMP TabBand::SetSite(IUnknown* pUnkSite) {
    if (!pUnkSite) {
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
        serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_webBrowser));
    }

    if ((!m_shellBrowser || !m_webBrowser) && site) {
        serviceProvider.Reset();
        if (SUCCEEDED(site.As(&serviceProvider)) && serviceProvider) {
            serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_shellBrowser));
            serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_webBrowser));
        }
    }

    if (!m_shellBrowser || !m_webBrowser) {
        DisconnectSite();
        return E_FAIL;
    }

    EnsureSessionStore();

    EnsureWindow();
    if (!m_window) {
        DisconnectSite();
        return E_FAIL;
    }

    m_browserEvents = std::make_unique<BrowserEvents>(this);
    if (m_browserEvents) {
        hr = m_browserEvents->Connect(m_webBrowser);
        if (FAILED(hr)) {
            m_browserEvents.reset();
        }
    }

    InitializeTabs();
    UpdateTabsUI();

    if (!m_viewColorizer) {
        m_viewColorizer = std::make_unique<FolderViewColorizer>();
    }
    if (m_viewColorizer) {
        m_viewColorizer->Attach(m_shellBrowser);
    }

    return S_OK;
}

IFACEMETHODIMP TabBand::GetSite(REFIID riid, void** ppvSite) {
    if (!ppvSite) {
        return E_POINTER;
    }
    if (!m_site) {
        *ppvSite = nullptr;
        return E_FAIL;
    }
    return m_site->QueryInterface(riid, ppvSite);
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
    if (m_viewColorizer) {
        m_viewColorizer->Refresh();
    }
    m_internalNavigation = false;
}

void TabBand::OnBrowserQuit() {
    DisconnectSite();
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
    if (location.IsValid()) {
        NavigateToTab(location);
    }
}

void TabBand::OnCloseTabRequested(TabLocation location) {
    const auto selected = m_tabs.SelectedLocation();
    const bool wasSelected = (selected.groupIndex == location.groupIndex && selected.tabIndex == location.tabIndex);

    m_tabs.Remove(location);

    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }

    UpdateTabsUI();

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

    OpenTabInNewWindow(*tab);
    m_tabs.Remove(location);
    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }
    UpdateTabsUI();
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
}

void TabBand::OnDetachGroupRequested(int groupIndex) {
    auto* group = m_tabs.GetGroup(groupIndex);
    if (!group) {
        return;
    }

    std::vector<TabLocation> tabsToDetach;
    tabsToDetach.reserve(group->tabs.size());
    for (size_t i = 0; i < group->tabs.size(); ++i) {
        tabsToDetach.emplace_back(TabLocation{groupIndex, static_cast<int>(i)});
    }

    for (auto it = tabsToDetach.rbegin(); it != tabsToDetach.rend(); ++it) {
        if (const auto* tab = m_tabs.Get(*it)) {
            OpenTabInNewWindow(*tab);
        }
        m_tabs.Remove(*it);
    }

    if (m_tabs.TotalTabCount() == 0) {
        EnsureTabForCurrentFolder();
    }
    UpdateTabsUI();
}

void TabBand::OnMoveTabRequested(TabLocation from, TabLocation to) {
    m_tabs.MoveTab(from, to);
    UpdateTabsUI();
}

void TabBand::OnMoveGroupRequested(int fromGroup, int toGroup) {
    m_tabs.MoveGroup(fromGroup, toGroup);
    UpdateTabsUI();
}

void TabBand::OnToggleSplitView(int groupIndex) {
    m_tabs.ToggleSplitView(groupIndex);
    UpdateTabsUI();
    if (m_tabs.IsSplitViewEnabled(groupIndex)) {
        EnsureSplitViewWindows(groupIndex);
    }
}

void TabBand::OnPromoteSplitSecondary(TabLocation location) {
    m_tabs.SetSplitSecondary(location);
    UpdateTabsUI();
    EnsureSplitViewWindows(location.groupIndex);
}

void TabBand::OnClearSplitSecondary(int groupIndex) {
    m_tabs.ClearSplitSecondary(groupIndex);
    UpdateTabsUI();
}

void TabBand::OnSwapSplitPanes(int groupIndex) {
    m_tabs.SwapSplitSelection(groupIndex);
    UpdateTabsUI();
    const TabLocation selection = m_tabs.SelectedLocation();
    if (selection.IsValid()) {
        NavigateToTab(selection);
    }
    EnsureSplitViewWindows(groupIndex);
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

void TabBand::EnsureWindow() {
    if (m_window) {
        return;
    }

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
    }
}

void TabBand::DisconnectSite() {
    SaveSession();

    if (m_browserEvents) {
        m_browserEvents->Disconnect();
        m_browserEvents.reset();
    }

    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();

    if (m_window) {
        m_window->Destroy();
        m_window.reset();
    }

    if (m_viewColorizer) {
        m_viewColorizer->Detach();
    }

    m_tabs.Clear();
    m_internalNavigation = false;
    m_viewColorizer.reset();
}

void TabBand::InitializeTabs() {
    m_tabs.Clear();

    EnsureSessionStore();
    const bool restored = RestoreSession();

    if (!restored || m_tabs.TotalTabCount() == 0) {
        UniquePidl pidl = QueryCurrentFolder();
        if (pidl) {
            std::wstring name = GetDisplayName(pidl.get());
            if (name.empty()) {
                name = L"Tab";
            }
            m_tabs.Add(std::move(pidl), name, name, true);
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

void TabBand::UpdateTabsUI() {
    const auto items = m_tabs.BuildView();
    if (m_window) {
        m_window->SetTabs(items);
    }
    SaveSession();
}

void TabBand::EnsureSessionStore() {
    if (!m_sessionStore) {
        m_sessionStore = std::make_unique<SessionStore>();
    }
}

bool TabBand::RestoreSession() {
    if (!m_sessionStore) {
        return false;
    }

    SessionData data;
    if (!m_sessionStore->Load(data)) {
        return false;
    }

    std::vector<TabGroup> groups;
    groups.reserve(data.groups.size());
    for (const auto& groupData : data.groups) {
        TabGroup group;
        group.name = groupData.name.empty() ? L"Island" : groupData.name;
        group.collapsed = groupData.collapsed;
        group.splitView = groupData.splitView;
        group.splitPrimary = groupData.splitPrimary;
        group.splitSecondary = groupData.splitSecondary;
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
        if (!group.tabs.empty()) {
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
        storedGroup.splitView = group->splitView;
        storedGroup.splitPrimary = group->splitPrimary;
        storedGroup.splitSecondary = group->splitSecondary;

        for (const auto& tab : group->tabs) {
            SessionTab storedTab;
            storedTab.path = GetParsingName(tab.pidl.get());
            if (storedTab.path.empty()) {
                continue;
            }
            storedTab.name = tab.name;
            storedTab.tooltip = tab.tooltip;
            storedTab.hidden = tab.hidden;
            storedGroup.tabs.emplace_back(std::move(storedTab));
        }

        if (!storedGroup.tabs.empty()) {
            data.groups.emplace_back(std::move(storedGroup));
        }
    }

    if (data.groups.empty()) {
        return;
    }

    m_sessionStore->Save(data);
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
    EnsureSplitViewWindows(location.groupIndex);
}

void TabBand::EnsureTabForCurrentFolder() {
    UniquePidl current = QueryCurrentFolder();
    if (!current) {
        return;
    }

    TabLocation existing = m_tabs.Find(current.get());
    std::wstring name = GetDisplayName(current.get());
    if (name.empty()) {
        name = L"Tab";
    }

    if (existing.IsValid()) {
        if (auto* tab = m_tabs.Get(existing)) {
            tab->name = name;
            tab->tooltip = name;
            tab->hidden = false;
            tab->path = GetParsingName(tab->pidl.get());
        }
        m_tabs.SetGroupCollapsed(existing.groupIndex, false);
        m_tabs.SetSelectedLocation(existing);
    } else {
        UniquePidl clone = ClonePidl(current.get());
        if (!clone) {
            return;
        }
        m_tabs.Add(std::move(clone), name, name, true);
    }
}

void TabBand::OpenTabInNewWindow(const TabInfo& tab) const {
    if (!m_shellBrowser || !tab.pidl) {
        return;
    }
    m_shellBrowser->BrowseObject(tab.pidl.get(), SBSP_NEWBROWSER | SBSP_EXPLORE | SBSP_ABSOLUTE);
}

std::vector<std::pair<TabLocation, std::wstring>> TabBand::GetHiddenTabs(int groupIndex) const {
    return m_tabs.GetHiddenTabs(groupIndex);
}

void TabBand::EnsureSplitViewWindows(int groupIndex) {
    if (!m_tabs.IsSplitViewEnabled(groupIndex)) {
        return;
    }
    TabLocation secondary = m_tabs.GetSplitSecondary(groupIndex);
    if (!secondary.IsValid()) {
        return;
    }
    if (const auto* tab = m_tabs.Get(secondary)) {
        OpenTabInNewWindow(*tab);
    }
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

}  // namespace shelltabs

