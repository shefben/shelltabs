#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <unknwn.h>

#include <atomic>
#include <memory>
#include <string>
#include <vector>

#include <exdisp.h>
#include <ocidl.h>
#include <shlobj.h>

#include <wrl/client.h>

namespace Gdiplus {
class Image;
}
namespace shelltabs {

class CExplorerBHO : public IObjectWithSite, public IDispatch {
public:
    CExplorerBHO();
    ~CExplorerBHO();
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IDispatch
    IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo) override;
    IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
    IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid,
                                 DISPID* rgDispId) override;
    IFACEMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                          DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                          UINT* puArgErr) override;

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* site) override;
    IFACEMETHODIMP GetSite(REFIID riid, void** site) override;

private:
    void Disconnect();
    HRESULT EnsureBandVisible();
    HRESULT ConnectEvents();
    void DisconnectEvents();
    HRESULT ResolveBrowserFromSite(IUnknown* site, IWebBrowser2** browser);
    void UpdateBreadcrumbSubclass();
    void RemoveBreadcrumbSubclass();
    HWND FindBreadcrumbToolbar() const;
    HWND FindBreadcrumbToolbarInWindow(HWND root) const;
    HWND FindProgressWindow() const;
    HWND GetTopLevelExplorerWindow() const;
    bool InstallBreadcrumbSubclass(HWND toolbar);
    bool InstallProgressSubclass(HWND progressWindow);
    void UpdateProgressSubclass();
    void RemoveProgressSubclass();
    void UpdateExplorerViewSubclass();
    void RemoveExplorerViewSubclass();
    bool InstallExplorerViewSubclass(HWND viewWindow, HWND listView, HWND treeView);
    bool HandleExplorerViewMessage(HWND source, UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* result);
    bool HandleListViewMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* result);
    void HandleExplorerContextMenuInit(HWND hwnd, HMENU menu);
    void PrepareContextMenuSelection(HWND sourceWindow, POINT screenPoint);
    void HandleExplorerCommand(UINT commandId);
    void HandleExplorerMenuDismiss(HMENU menu);
    bool CollectSelectedFolderPaths(std::vector<std::wstring>& paths) const;
    bool CollectPathsFromShellViewSelection(std::vector<std::wstring>& paths) const;
    bool CollectPathsFromFolderViewSelection(std::vector<std::wstring>& paths) const;
    bool CollectPathsFromItemArray(IShellItemArray* items, std::vector<std::wstring>& paths) const;
    bool CollectPathsFromListView(std::vector<std::wstring>& paths) const;
    bool CollectPathsFromTreeView(std::vector<std::wstring>& paths) const;
    bool AppendPathFromPidl(PCIDLIST_ABSOLUTE pidl, std::vector<std::wstring>& paths) const;
    void DispatchOpenInNewTab(const std::vector<std::wstring>& paths) const;
    void ClearPendingOpenInNewTabState();
    void EnsureBreadcrumbHook();
    void RemoveBreadcrumbHook();
    bool IsBreadcrumbToolbarCandidate(HWND hwnd) const;
    bool IsBreadcrumbToolbarAncestor(HWND hwnd) const;
    bool IsWindowOwnedByThisExplorer(HWND hwnd) const;
    bool HandleBreadcrumbPaint(HWND hwnd);
    bool HandleProgressPaint(HWND hwnd);
    enum class BreadcrumbDiscoveryStage {
        None,
        ServiceUnavailable,
        ServiceWindowMissing,
        ServiceToolbarMissing,
        FrameMissing,
        RebarMissing,
        ParentMissing,
        ToolbarMissing,
        Discovered,
    };
    void LogBreadcrumbStage(BreadcrumbDiscoveryStage stage, const wchar_t* format, ...) const;
    void UpdateFolderBackgroundOptions(const ShellTabsOptions& options);
    void ReloadFolderBackgroundImage();
    void ApplyListViewBackgroundMode();
    static LRESULT CALLBACK BreadcrumbCbtProc(int code, WPARAM wParam, LPARAM lParam);
    static LRESULT CALLBACK BreadcrumbSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR subclassId, DWORD_PTR refData);
    static LRESULT CALLBACK ProgressSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                 UINT_PTR subclassId, DWORD_PTR refData);
    static LRESULT CALLBACK ExplorerViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                     UINT_PTR subclassId, DWORD_PTR refData);

    std::atomic<long> m_refCount;
    Microsoft::WRL::ComPtr<IUnknown> m_site;
    Microsoft::WRL::ComPtr<IWebBrowser2> m_webBrowser;
    Microsoft::WRL::ComPtr<IShellBrowser> m_shellBrowser;
    Microsoft::WRL::ComPtr<IConnectionPoint> m_connectionPoint;
    DWORD m_connectionCookie = 0;
    bool m_bandVisible = false;
    bool m_shouldRetryEnsure = true;
    HWND m_breadcrumbToolbar = nullptr;
    bool m_breadcrumbSubclassInstalled = false;
    bool m_breadcrumbGradientEnabled = false;
    bool m_breadcrumbFontGradientEnabled = false;
    int m_breadcrumbGradientTransparency = 45;
    int m_breadcrumbFontBrightness = 85;
    bool m_useCustomBreadcrumbGradientColors = false;
    COLORREF m_breadcrumbGradientStartColor = RGB(255, 59, 48);
    COLORREF m_breadcrumbGradientEndColor = RGB(175, 82, 222);
    bool m_useCustomBreadcrumbFontColors = false;
    COLORREF m_breadcrumbFontGradientStartColor = RGB(255, 255, 255);
    COLORREF m_breadcrumbFontGradientEndColor = RGB(255, 255, 255);
    bool m_useCustomProgressGradientColors = false;
    COLORREF m_progressGradientStartColor = RGB(0, 120, 215);
    COLORREF m_progressGradientEndColor = RGB(0, 153, 255);
    HWND m_progressWindow = nullptr;
    bool m_progressSubclassInstalled = false;
    bool m_breadcrumbHookRegistered = false;
    enum class BreadcrumbLogState {
        Unknown,
        Disabled,
        Searching,
    };
    BreadcrumbLogState m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    bool m_loggedBreadcrumbToolbarMissing = false;
    bool m_bufferedPaintInitialized = false;
    bool m_gdiplusInitialized = false;
    ULONG_PTR m_gdiplusToken = 0;
    mutable BreadcrumbDiscoveryStage m_lastBreadcrumbStage = BreadcrumbDiscoveryStage::None;
    Microsoft::WRL::ComPtr<IShellView> m_shellView;
    HWND m_shellViewWindow = nullptr;
    bool m_shellViewWindowSubclassInstalled = false;
    HWND m_frameWindow = nullptr;
    bool m_frameSubclassInstalled = false;
    HWND m_listView = nullptr;
    HWND m_treeView = nullptr;
    bool m_listViewSubclassInstalled = false;
    bool m_treeViewSubclassInstalled = false;
    HMENU m_trackedContextMenu = nullptr;
    std::vector<std::wstring> m_pendingOpenInNewTabPaths;
    bool m_contextMenuInserted = false;
    bool m_folderBackgroundImageEnabled = false;
    bool m_folderBackgroundImageActive = false;
    std::wstring m_folderBackgroundImagePath;
    std::unique_ptr<Gdiplus::Image> m_folderBackgroundImage;
    static constexpr UINT kOpenInNewTabCommandId = 0xE170;
    static constexpr UINT kMaxTrackedSelection = 16;
};

}  // namespace shelltabs

