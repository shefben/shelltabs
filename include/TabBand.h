#pragma once

#include <windows.h>

#include <atomic>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include <ShObjIdl_core.h>
#include <exdisp.h>
#include <shlobj.h>
#include <wrl/client.h>

#include "BrowserEvents.h"
#include "TabManager.h"
#include "SessionStore.h"
#include "OptionsStore.h"

namespace shelltabs {

class TabBandWindow;
class TabBand : public IDeskBand2,
                public IObjectWithSite,
                public IInputObject,
                public IPersistStream {
    friend class TabBandWindow;

public:
    TabBand();
    ~TabBand();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IOleWindow (through IDeskBand)
    IFACEMETHODIMP GetWindow(HWND* phwnd) override;
    IFACEMETHODIMP ContextSensitiveHelp(BOOL) override;

    // IDockingWindow
    IFACEMETHODIMP ShowDW(BOOL fShow) override;
    IFACEMETHODIMP CloseDW(DWORD dwReserved) override;
    IFACEMETHODIMP ResizeBorderDW(const RECT*, IUnknown*, BOOL) override;

    // IDeskBand
    IFACEMETHODIMP GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi) override;

    // IDeskBand2
    IFACEMETHODIMP CanRenderComposited(BOOL* pfCanRenderComposited) override;
    IFACEMETHODIMP SetCompositionState(BOOL fCompositionEnabled) override;
    IFACEMETHODIMP GetCompositionState(BOOL* pfCompositionEnabled) override;

    // IInputObject
    IFACEMETHODIMP UIActivateIO(BOOL fActivate, LPMSG lpMsg) override;
    IFACEMETHODIMP HasFocusIO() override;
    IFACEMETHODIMP TranslateAcceleratorIO(LPMSG lpMsg) override;

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* pUnkSite) override;
    IFACEMETHODIMP GetSite(REFIID riid, void** ppvSite) override;

    // IPersist
    IFACEMETHODIMP GetClassID(CLSID* pClassID) override;

    // IPersistStream
    IFACEMETHODIMP IsDirty() override;
    IFACEMETHODIMP Load(IStream* pStm) override;
    IFACEMETHODIMP Save(IStream* pStm, BOOL fClearDirty) override;
    IFACEMETHODIMP GetSizeMax(ULARGE_INTEGER* pcbSize) override;

    void OnBrowserNavigate();
    void OnBrowserQuit();
    bool OnBrowserNewWindow(const std::wstring& targetUrl);
    bool OnCtrlBeforeNavigate(const std::wstring& url);
    void OnTabSelected(TabLocation location);
    void OnNewTabRequested(int targetGroup = -1);
    void OnCloseTabRequested(TabLocation location);
    void OnCloseOtherTabsRequested(TabLocation location);
    void OnCloseTabsToRightRequested(TabLocation location);
    void OnCloseTabsToLeftRequested(TabLocation location);
    void OnReopenClosedTabRequested();
    void OnHideTabRequested(TabLocation location);
    void OnUnhideTabRequested(TabLocation location);
    void OnDetachTabRequested(TabLocation location);
    void OnCloneTabRequested(TabLocation location);
    void OnToggleTabPinned(TabLocation location);
    void OnToggleGroupCollapsed(int groupIndex);
    void OnUnhideAllInGroup(int groupIndex);
    void OnCreateIslandAfter(int groupIndex);
    void OnCloseIslandRequested(int groupIndex);
    void OnEditGroupProperties(int groupIndex);
    void OnDetachGroupRequested(int groupIndex);
    void OnMoveTabRequested(TabLocation from, TabLocation to);
    void OnMoveGroupRequested(int fromGroup, int toGroup);
    void OnMoveTabToNewGroup(TabLocation from, int insertIndex, bool headerVisible);
    std::optional<TabInfo> DetachTabForTransfer(TabLocation location, bool* wasSelected,
                                                bool ensurePlaceholderTab = true,
                                                bool* removedLastTab = nullptr);
    TabLocation InsertTransferredTab(TabInfo tab, int groupIndex, int tabIndex, bool createGroup, bool headerVisible,
                                     bool select);
    std::optional<TabGroup> DetachGroupForTransfer(int groupIndex, bool* wasSelected);
    int InsertTransferredGroup(TabGroup group, int insertIndex, bool select);
    void OnSetGroupHeaderVisible(int groupIndex, bool visible);
    void OnOpenTerminal(TabLocation location);
    void OnOpenVSCode(TabLocation location);
    void OnCopyPath(TabLocation location);
    void OnFilesDropped(TabLocation location, const std::vector<std::wstring>& paths, bool move);
    void OnOpenFolderInNewTab(const std::wstring& path, bool select = true);
    void CloseFrameWindowAsync();
    void EnsureTabPreview(TabLocation location);

    TabManager& GetTabManager() noexcept { return m_tabs; }
    const TabManager& GetTabManager() const noexcept { return m_tabs; }

    bool CanCloseOtherTabs(TabLocation location) const;
    bool CanCloseTabsToRight(TabLocation location) const;
    bool CanCloseTabsToLeft(TabLocation location) const;
    bool CanReopenClosedTabs() const;

    std::vector<std::pair<TabLocation, std::wstring>> GetHiddenTabs(int groupIndex) const;
    int GetGroupCount() const noexcept;
    bool IsGroupHeaderVisible(int groupIndex) const;
    bool BuildExplorerContextMenu(TabLocation location, HMENU menu, UINT idFirst, UINT idLast,
                                  Microsoft::WRL::ComPtr<IContextMenu>* menuOut,
                                  Microsoft::WRL::ComPtr<IContextMenu2>* menu2Out,
                                  Microsoft::WRL::ComPtr<IContextMenu3>* menu3Out,
                                  UINT* usedLast) const;
    bool InvokeExplorerContextCommand(TabLocation location, IContextMenu* menu, UINT commandId,
                                      UINT idFirst, const POINT& ptInvoke) const;
    std::vector<std::wstring> GetSavedGroupNames() const;
    void OnCreateSavedGroup(int afterGroup);
    void OnLoadSavedGroup(const std::wstring& name, int afterGroup);
    void OnShowOptionsDialog(int initialTab = 0, const std::wstring& focusGroupId = std::wstring(),
                             bool editFocusedGroup = false);
    void OnSavedGroupsChanged();
    void OnDeferredNavigate();
    void OnDockingModeChanged(TabBandDockMode mode);
    std::wstring GetSavedGroupId(int groupIndex) const;

private:
    std::atomic<long> m_refCount;
    DWORD m_bandId = 0;
    DWORD m_viewMode = 0;
    bool m_isComposited = false;

    Microsoft::WRL::ComPtr<IInputObjectSite> m_site;
    Microsoft::WRL::ComPtr<IOleWindow> m_siteOleWindow;
    Microsoft::WRL::ComPtr<IDockingWindowSite> m_dockingSite;
    Microsoft::WRL::ComPtr<IShellBrowser> m_shellBrowser;
    Microsoft::WRL::ComPtr<IWebBrowser2> m_webBrowser;

    std::unique_ptr<TabBandWindow> m_window;
    TabManager m_tabs;
    std::unique_ptr<SessionStore> m_sessionStore;
    bool m_restoringSession = false;
    std::wstring m_windowToken;
    mutable ShellTabsOptions m_options{};
    mutable bool m_optionsLoaded = false;
    bool m_sessionMarkerActive = false;
    bool m_lastSessionUnclean = false;
    bool m_sessionFlushTimerActive = false;
    bool m_sessionFlushTimerPending = false;

    std::unique_ptr<BrowserEvents> m_browserEvents;
    DWORD m_browserCookie = 0;
    bool m_internalNavigation = false;
    int m_allowExternalNewWindows = 0;
    TabLocation m_pendingNavigation;
    bool m_deferredNavigationPosted = false;
    TabBandDockMode m_dockMode = TabBandDockMode::kAutomatic;
    TabBandDockMode m_requestedDockMode = TabBandDockMode::kAutomatic;
    mutable bool m_skipSavedGroupSync = false;
    uint64_t m_processedGroupStoreGeneration = 0;

    struct ClosedGroupMetadata {
        std::wstring name;
        bool collapsed = false;
        bool headerVisible = true;
        bool hasOutline = false;
        COLORREF outlineColor = RGB(0, 120, 215);
        TabGroupOutlineStyle outlineStyle = TabGroupOutlineStyle::kSolid;
        std::wstring savedGroupId;
    };

    struct ClosedTabEntry {
        int originalIndex = -1;
        TabInfo tab;
    };

    struct ClosedTabSet {
        int groupIndex = -1;
        std::vector<ClosedTabEntry> entries;
        std::optional<ClosedGroupMetadata> groupInfo;
        bool groupRemoved = false;
        int selectionOriginalIndex = -1;
    };

    std::vector<ClosedTabSet> m_closedTabHistory;

    ClosedGroupMetadata CaptureGroupMetadata(const TabGroup& group) const;
    void EnsureTabPath(TabInfo& tab) const;
    void PushClosedSet(ClosedTabSet set);
    std::optional<ClosedTabSet> BuildClosedSetFromSession(const SessionClosedSet& stored) const;
    std::optional<SessionClosedSet> BuildSessionClosedSet(const ClosedTabSet& set) const;

    void EnsureWindow();
    void EnsureOptionsLoaded() const;
    void DisconnectSite();
    void InitializeTabs();
    void UpdateTabsUI();
    void EnsureSessionStore();
    bool RestoreSession();
    void SaveSession();
    void StartSessionFlushTimer();
    void StopSessionFlushTimer();
    void OnPeriodicSessionFlush();
    void ApplyOptionsChanges(const ShellTabsOptions& previousOptions);
    UniquePidl QueryCurrentFolder() const;
    void CancelPendingPreviewForTab(const TabInfo& tab) const;
    void CancelPendingPreviewForGroup(const TabGroup& group) const;
    void NavigateToTab(TabLocation location);
    void EnsureTabForCurrentFolder();
    void OpenTabInNewWindow(const TabInfo& tab);
    bool LaunchShellExecute(const std::wstring& application, const std::wstring& parameters,
                            const std::wstring& workingDirectory) const;
    std::wstring GetTabPath(TabLocation location) const;
    void PerformFileOperation(TabLocation location, const std::vector<std::wstring>& paths, bool move);
    bool HandleNewWindowRequest(const std::wstring& targetUrl);
    void QueueNavigateTo(TabLocation location);
    void SyncSavedGroup(int groupIndex) const;
    void SyncAllSavedGroups() const;
    bool ApplySavedGroupMetadata(const std::vector<SavedGroup>& savedGroups,
                                 const std::vector<std::pair<std::wstring, std::wstring>>& renamedGroups,
                                 const std::vector<std::wstring>& removedGroupIds);
    HWND GetFrameWindow() const;
    TabManager::ExplorerWindowId BuildWindowId() const;
    std::wstring ResolveWindowToken();
    void ReleaseWindowToken();
    void CaptureActiveTabPreview();
};

}  // namespace shelltabs

