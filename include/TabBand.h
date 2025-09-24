#pragma once

#include <windows.h>

#include <atomic>
#include <memory>
#include <vector>

#include <ShObjIdl_core.h>
#include <exdisp.h>
#include <shlobj.h>
#include <wrl/client.h>

#include "BrowserEvents.h"
#include "TabManager.h"
#include "FolderViewColorizer.h"
#include "SessionStore.h"

namespace shelltabs {

class TabBandWindow;

class TabBand : public IDeskBand2,
                public IObjectWithSite,
                public IInputObject,
                public IPersistStream {
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

    void OnTabSelected(TabLocation location);
    void OnNewTabRequested();
    void OnCloseTabRequested(TabLocation location);
    void OnHideTabRequested(TabLocation location);
    void OnUnhideTabRequested(TabLocation location);
    void OnDetachTabRequested(TabLocation location);
    void OnToggleGroupCollapsed(int groupIndex);
    void OnUnhideAllInGroup(int groupIndex);
    void OnCreateIslandAfter(int groupIndex);
    void OnDetachGroupRequested(int groupIndex);
    void OnMoveTabRequested(TabLocation from, TabLocation to);
    void OnMoveGroupRequested(int fromGroup, int toGroup);
    void OnToggleSplitView(int groupIndex);
    void OnPromoteSplitSecondary(TabLocation location);
    void OnClearSplitSecondary(int groupIndex);
    void OnSwapSplitPanes(int groupIndex);
    void OnOpenTerminal(TabLocation location);
    void OnOpenVSCode(TabLocation location);
    void OnCopyPath(TabLocation location);
    void OnFilesDropped(TabLocation location, const std::vector<std::wstring>& paths, bool move);

    std::vector<std::pair<TabLocation, std::wstring>> GetHiddenTabs(int groupIndex) const;

private:
    std::atomic<long> m_refCount;
    DWORD m_bandId = 0;
    DWORD m_viewMode = 0;
    bool m_isComposited = false;

    Microsoft::WRL::ComPtr<IInputObjectSite> m_site;
    Microsoft::WRL::ComPtr<IShellBrowser> m_shellBrowser;
    Microsoft::WRL::ComPtr<IWebBrowser2> m_webBrowser;

    std::unique_ptr<TabBandWindow> m_window;
    TabManager m_tabs;
    std::unique_ptr<FolderViewColorizer> m_viewColorizer;
    std::unique_ptr<SessionStore> m_sessionStore;
    bool m_restoringSession = false;

    std::unique_ptr<BrowserEvents> m_browserEvents;
    DWORD m_browserCookie = 0;
    bool m_internalNavigation = false;

    void EnsureWindow();
    void DisconnectSite();
    void InitializeTabs();
    void UpdateTabsUI();
    void EnsureSessionStore();
    bool RestoreSession();
    void SaveSession();
    UniquePidl QueryCurrentFolder() const;
    void NavigateToTab(TabLocation location);
    void EnsureTabForCurrentFolder();
    void OpenTabInNewWindow(const TabInfo& tab) const;
    void EnsureSplitViewWindows(int groupIndex);
    bool LaunchShellExecute(const std::wstring& application, const std::wstring& parameters,
                            const std::wstring& workingDirectory) const;
    std::wstring GetTabPath(TabLocation location) const;
    void PerformFileOperation(TabLocation location, const std::vector<std::wstring>& paths, bool move);
};

}  // namespace shelltabs

