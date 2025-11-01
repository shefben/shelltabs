#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <CommCtrl.h>
#include <unknwn.h>

#include <atomic>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>
#include <cstdint>

#include <exdisp.h>
#include <ocidl.h>
#include <shlobj.h>
#include <shobjidl_core.h>

#include <wrl/client.h>

#include "ExplorerGlowSurfaces.h"
#include "OptionsStore.h"
#include "PaneHooks.h"
#include "Utilities.h"

namespace Gdiplus {
class Bitmap;
}
namespace shelltabs {

struct ShellTabsOptions;
class NamespaceTreeHost;

        class CExplorerBHO : public IObjectWithSite,
                             public IDispatch,
                             public PaneHighlightProvider {
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

                // PaneHighlightProvider
                bool TryGetListViewHighlight(HWND listView, int itemIndex, PaneHighlight* highlight) override;

        private:
                enum class BandEnsureOutcome {
                        Unknown,
                        Success,
                        PermanentFailure,
                        TemporaryFailure,
                        Throttled,
                };

                struct BandEnsureState {
                        BandEnsureOutcome lastOutcome = BandEnsureOutcome::Unknown;
                        size_t attemptCount = 0;
                        DWORD retryDelayMs = 0;
                        UINT_PTR timerId = 0;
                        bool retryScheduled = false;
                        bool unsupportedHost = false;
                        HRESULT lastHresult = S_OK;
                };

                struct HandleHasher {
                        size_t operator()(HWND hwnd) const noexcept {
                                return reinterpret_cast<size_t>(hwnd);
                        }
                };

                struct ListViewAccentResources {
                        COLORREF accentColor = 0;
                        COLORREF textColor = RGB(255, 255, 255);
                        HBRUSH backgroundBrush = nullptr;
                        HPEN focusPen = nullptr;
                };

                struct ListViewBackgroundSurface {
                        std::unique_ptr<Gdiplus::Bitmap> bitmap;
                        SIZE size{0, 0};
                        std::wstring cacheKey;
                        bool pending = false;
                };

                struct ListViewBackgroundJobState {
                        SIZE size{0, 0};
                        std::wstring cacheKey;
                        uint64_t generation = 0;
                        bool active = false;
                };

                struct TreeItemPidlResolution {
                        UniquePidl owned;
                        PCIDLIST_ABSOLUTE raw = nullptr;

                        [[nodiscard]] bool empty() const noexcept { return raw == nullptr; }
                };

                struct ContextMenuSelectionItem {
                        UniquePidl pidl;
                        PCIDLIST_ABSOLUTE raw = nullptr;
                        DWORD attributes = 0;
                        bool isFolder = false;
                        bool isFileSystem = false;
                        std::wstring path;
                        std::wstring parentPath;
                        std::wstring extension;

                        ContextMenuSelectionItem() = default;

                        ContextMenuSelectionItem(const ContextMenuSelectionItem& other)
                            : pidl(other.pidl ? ClonePidl(other.raw) : nullptr),
                              raw(pidl ? pidl.get() : other.raw),
                              attributes(other.attributes),
                              isFolder(other.isFolder),
                              isFileSystem(other.isFileSystem),
                              path(other.path),
                              parentPath(other.parentPath),
                              extension(other.extension) {}

                        ContextMenuSelectionItem(ContextMenuSelectionItem&& other) noexcept
                            : pidl(std::move(other.pidl)),
                              raw(pidl ? pidl.get() : other.raw),
                              attributes(other.attributes),
                              isFolder(other.isFolder),
                              isFileSystem(other.isFileSystem),
                              path(std::move(other.path)),
                              parentPath(std::move(other.parentPath)),
                              extension(std::move(other.extension)) {
                                other.raw = other.pidl ? other.pidl.get() : nullptr;
                        }

                        ContextMenuSelectionItem& operator=(const ContextMenuSelectionItem& other) {
                                if (this == &other) {
                                        return *this;
                                }

                                pidl = other.pidl ? ClonePidl(other.raw) : nullptr;
                                raw = pidl ? pidl.get() : other.raw;
                                attributes = other.attributes;
                                isFolder = other.isFolder;
                                isFileSystem = other.isFileSystem;
                                path = other.path;
                                parentPath = other.parentPath;
                                extension = other.extension;
                                return *this;
                        }

                        ContextMenuSelectionItem& operator=(ContextMenuSelectionItem&& other) noexcept {
                                if (this == &other) {
                                        return *this;
                                }

                                pidl = std::move(other.pidl);
                                raw = pidl ? pidl.get() : other.raw;
                                attributes = other.attributes;
                                isFolder = other.isFolder;
                                isFileSystem = other.isFileSystem;
                                path = std::move(other.path);
                                parentPath = std::move(other.parentPath);
                                extension = std::move(other.extension);

                                other.raw = other.pidl ? other.pidl.get() : nullptr;
                                return *this;
                        }
                };

                struct ContextMenuSelectionSnapshot {
                        std::vector<ContextMenuSelectionItem> items;
                        size_t fileCount = 0;
                        size_t folderCount = 0;

                        void Clear() {
                                items.clear();
                                fileCount = 0;
                                folderCount = 0;
                        }
                };

                struct PreparedMenuItem {
                        const ContextMenuItem* definition = nullptr;
                        ContextMenuItemType type = ContextMenuItemType::kCommand;
                        ContextMenuInsertionAnchor anchor = ContextMenuInsertionAnchor::kDefault;
                        HMENU submenu = nullptr;
                        UINT commandId = 0;
                        HBITMAP bitmap = nullptr;
                        bool enabled = true;
                        std::wstring label;
                };

                void Disconnect();
                HRESULT EnsureBandVisible();
                HRESULT ConnectEvents();
                void DisconnectEvents();
                HRESULT ResolveBrowserFromSite(IUnknown* site, IWebBrowser2** browser);
                void ScheduleEnsureRetry(HWND hostWindow, BandEnsureState& state, HRESULT lastHr,
                                         BandEnsureOutcome outcome, const wchar_t* reason = nullptr);
                void CancelEnsureRetry(BandEnsureState& state);
                void CancelAllEnsureRetries();
                void HandleEnsureBandTimer(UINT_PTR timerId);
                static void CALLBACK EnsureBandTimerProc(HWND hwnd, UINT msg, UINT_PTR timerId, DWORD tickCount);
                void UpdateBreadcrumbSubclass();
		void RemoveBreadcrumbSubclass();
		HWND FindBreadcrumbToolbar() const;
		HWND FindBreadcrumbToolbarInWindow(HWND root) const;
		HWND FindProgressWindow() const;
		HWND FindAddressEditControl() const;
		HWND GetTopLevelExplorerWindow() const;
		bool InstallBreadcrumbSubclass(HWND toolbar);
                bool InstallProgressSubclass(HWND progressWindow);
                bool InstallAddressEditSubclass(HWND editWindow);
                void UpdateProgressSubclass();
                void RemoveProgressSubclass();
                void UpdateAddressEditSubclass();
                void RemoveAddressEditSubclass();
                void RequestAddressEditRedraw(HWND hwnd) const;
                void ResetAddressEditStateCache();
                bool RefreshAddressEditState(HWND hwnd, bool updateText, bool updateSelection,
                                             bool updateFocus, bool updateTheme);
                bool RefreshAddressEditText(HWND hwnd);
                bool RefreshAddressEditSelection(HWND hwnd);
                bool RefreshAddressEditFocus(HWND hwnd);
                bool RefreshAddressEditTheme();
                bool RefreshAddressEditFont(HWND hwnd);
                bool EnsureProgressGradientResources();
                void DestroyProgressGradientResources();
                void UpdateExplorerViewSubclass();
                void RemoveExplorerViewSubclass();
                bool InstallExplorerViewSubclass(HWND viewWindow);
                bool TryResolveExplorerPanes();
                void HandleExplorerPaneCandidate(HWND candidate);
                void UpdateExplorerPaneCreationWatch(bool watchListView, bool watchTreeView);
                void ScheduleExplorerPaneRetry();
                void CancelExplorerPaneRetry(bool resetAttemptState = true);
                void ScheduleExplorerPaneFallback();
                void CancelExplorerPaneFallback();
                //bool InstallExplorerViewSubclass(HWND viewWindow, HWND listView, HWND treeView, HWND directUiHost);
                void TryAttachNamespaceTreeControl(IShellView* shellView);
                void ResetNamespaceTreeControl();
                void InvalidateNamespaceTreeControl() const;
                bool HandleExplorerViewMessage(HWND source, UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* result);
                void HandleExplorerPostPaint(HWND hwnd, UINT msg, WPARAM wParam);
                void ReloadFolderBackgrounds(const ShellTabsOptions& options);
                void ClearFolderBackgrounds();
                std::wstring NormalizeBackgroundKey(const std::wstring& path) const;
                bool EnsureFolderBackgroundBitmap(const std::wstring& key) const;
                bool EnsureUniversalBackgroundBitmap() const;
                Gdiplus::Bitmap* ResolveCurrentFolderBackground() const;
                bool DrawFolderBackground(HWND hwnd, HDC dc);
                void UpdateCurrentFolderBackground();
                void InvalidateFolderBackgroundTargets() const;
                bool EnsureListViewBackgroundSurface(HWND listView);
                bool PaintListViewBackground(HWND hwnd, HDC dc);
                void ClearListViewBackgroundImage();
                void StartListViewBackgroundJob(SIZE size, std::wstring cacheKey,
                        std::unique_ptr<Gdiplus::Bitmap> source);
                void CancelListViewBackgroundJob();
                std::wstring ResolveBackgroundCacheKey() const;
                void HandleExplorerContextMenuInit(HWND hwnd, HMENU menu);
                void PrepareContextMenuSelection(HWND sourceWindow, POINT screenPoint);
                void HandleExplorerCommand(UINT commandId);
                void HandleExplorerMenuDismiss(HMENU menu);
                bool CollectSelectedFolderPaths(std::vector<std::wstring>& paths) const;
                bool CollectContextMenuSelection(ContextMenuSelectionSnapshot& selection) const;
                bool CollectContextSelectionFromShellView(ContextMenuSelectionSnapshot& selection) const;
                bool CollectContextSelectionFromFolderView(ContextMenuSelectionSnapshot& selection) const;
                bool CollectContextSelectionFromItemArray(IShellItemArray* items,
                        ContextMenuSelectionSnapshot& selection) const;
                bool CollectContextSelectionFromListView(ContextMenuSelectionSnapshot& selection) const;
                bool CollectContextSelectionFromTreeView(ContextMenuSelectionSnapshot& selection) const;
                bool AppendSelectionItemFromShellItem(IShellItem* item, ContextMenuSelectionSnapshot& selection) const;
                bool AppendSelectionItemFromPidl(PCIDLIST_ABSOLUTE pidl,
                        ContextMenuSelectionSnapshot& selection) const;
                bool PopulateCustomContextMenus(HMENU menu, const ContextMenuSelectionSnapshot& selection,
                        bool anchorFound, UINT anchorPosition);
                bool PopulateCustomSubmenu(HMENU submenu, const std::vector<ContextMenuItem>& items,
                        const ContextMenuSelectionSnapshot& selection);
                std::optional<PreparedMenuItem> PrepareMenuItem(const ContextMenuItem& item,
                        const ContextMenuSelectionSnapshot& selection, bool allowSubmenuAnchors);
                bool ShouldDisplayMenuItem(const ContextMenuItem& item,
                        const ContextMenuSelectionSnapshot& selection) const;
                bool DoesSelectionMatchScope(const ContextMenuItemScope& scope,
                        const ContextMenuSelectionSnapshot& selection) const;
                bool IsSelectionCountAllowed(const ContextMenuSelectionRule& rule, size_t selectionCount) const;
                UINT AllocateContextMenuCommandId(HMENU menu);
                void TrackContextCommand(UINT commandId, const ContextMenuItem* item);
                void ExecuteContextMenuCommand(const ContextMenuItem& item) const;
                std::vector<std::wstring> BuildCommandLines(const ContextMenuItem& item) const;
                std::wstring ExpandCommandTemplate(const std::wstring& commandTemplate,
                        const ContextMenuSelectionItem* item) const;
                std::wstring ExpandAggregateTokens(const std::wstring& commandTemplate) const;
                bool ExecuteCommandLine(const std::wstring& commandLine) const;
                HBITMAP CreateBitmapFromIcon(HICON icon, SIZE desiredSize) const;
                void CleanupContextMenuResources();
                TreeItemPidlResolution ResolveTreeViewItemPidl(HWND treeView, const TVITEMEXW& item) const;
                bool ResolveHighlightFromPidl(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) const;
                bool AppendPathFromPidl(PCIDLIST_ABSOLUTE pidl, std::vector<std::wstring>& paths) const;
                void DispatchOpenInNewTab(const std::vector<std::wstring>& paths);
                void QueueOpenInNewTabRequests(const std::vector<std::wstring>& paths);
                void TryDispatchQueuedOpenInNewTabRequests();
                void HandleOpenInNewTabTimer(UINT_PTR timerId);
                void ScheduleOpenInNewTabRetry();
                void CancelOpenInNewTabRetry();
                static void CALLBACK OpenInNewTabTimerProc(HWND hwnd, UINT msg, UINT_PTR timerId, DWORD tickCount);
                void ClearPendingOpenInNewTabState();
                void EnsureBreadcrumbHook();
                void RemoveBreadcrumbHook();
                bool IsBreadcrumbToolbarCandidate(HWND hwnd) const;
                bool IsBreadcrumbToolbarAncestor(HWND hwnd) const;
                bool IsWindowOwnedByThisExplorer(HWND hwnd) const;
                bool HandleBreadcrumbPaint(HWND hwnd);
                bool HandleProgressPaint(HWND hwnd);
                bool HandleAddressEditPaint(HWND hwnd);
                bool DrawAddressEditContent(HWND hwnd, HDC dc);
                bool ShouldUseListViewAccentColors() const;
                bool ResolveActiveGroupAccent(COLORREF* accent, COLORREF* text) const;
                bool EnsureListViewAccentResources(HWND listView, ListViewAccentResources** resources);
                void ClearListViewAccentResources(HWND listView);
                void ClearListViewAccentResources();
                bool HandleListViewAccentCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result);
                void RefreshListViewAccentState();
                void EnsureListViewSubclass();
                void EnsureListViewHostSubclass(HWND hostWindow);
                bool TryAttachListViewFromFolderView();
                HWND ResolveListViewFromFolderView();
                bool AttachTreeView(HWND treeView);
                bool RegisterGlowSurface(HWND hwnd, ExplorerSurfaceKind kind, bool ensureSubclass);
                void UnregisterGlowSurface(HWND hwnd);
                void UpdateGlowSurfaceTargets();
                void UpdateStatusBarTheme();
                void ResetStatusBarTheme(HWND statusBar = nullptr);
                void PruneGlowSurfaces(const std::unordered_set<HWND, HandleHasher>& active);
                void ResetGlowSurfaces();
                bool AttachListView(HWND listView);
                void DetachListView();
                void DetachListViewHosts();
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
		static LRESULT CALLBACK BreadcrumbCbtProc(int code, WPARAM wParam, LPARAM lParam);
		static LRESULT CALLBACK BreadcrumbSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
			UINT_PTR subclassId, DWORD_PTR refData);
		static LRESULT CALLBACK ProgressSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
			UINT_PTR subclassId, DWORD_PTR refData);
		static LRESULT CALLBACK AddressEditSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
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
                int m_breadcrumbHighlightAlphaMultiplier = 100;
                int m_breadcrumbDropdownAlphaMultiplier = 100;
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
                HBITMAP m_progressGradientBitmap = nullptr;
                void* m_progressGradientBits = nullptr;
                BITMAPINFO m_progressGradientInfo{};
                COLORREF m_progressGradientBitmapStartColor = 0;
                COLORREF m_progressGradientBitmapEndColor = 0;
                HWND m_addressEditWindow = nullptr;
                bool m_addressEditSubclassInstalled = false;
                mutable bool m_addressEditRedrawPending = false;
                mutable bool m_addressEditRedrawTimerActive = false;
                std::wstring m_addressEditCachedText;
                DWORD m_addressEditCachedSelStart = 0;
                DWORD m_addressEditCachedSelEnd = 0;
                bool m_addressEditCachedHasFocus = false;
                bool m_addressEditCachedThemeActive = false;
                HFONT m_addressEditCachedFont = nullptr;
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
                Microsoft::WRL::ComPtr<IFolderView2> m_folderView2;
                HWND m_shellViewWindow = nullptr;
                bool m_shellViewWindowSubclassInstalled = false;
                HWND m_frameWindow = nullptr;
                bool m_frameSubclassInstalled = false;
                HWND m_directUiView = nullptr;
                bool m_directUiSubclassInstalled = false;
                HWND m_listView = nullptr;
                HWND m_treeView = nullptr;
                bool m_listViewSubclassInstalled = false;
                bool m_treeViewSubclassInstalled = false;
                std::unordered_set<HWND, HandleHasher> m_listViewHostSubclassed;
                std::unordered_map<HWND, std::unique_ptr<ExplorerGlowSurface>, HandleHasher> m_glowSurfaces;
                bool m_watchListViewCreation = false;
                bool m_watchTreeViewCreation = false;
                HWND m_statusBar = nullptr;
                COLORREF m_statusBarBackgroundColor = CLR_DEFAULT;
                COLORREF m_statusBarTextColor = CLR_DEFAULT;
                bool m_statusBarThemeValid = false;
                bool m_statusBarBackgroundApplied = false;
                bool m_explorerPaneRetryPending = false;
                UINT_PTR m_explorerPaneRetryTimerId = 0;
                DWORD m_explorerPaneRetryDelayMs = 0;
                size_t m_explorerPaneRetryAttempts = 0;
                bool m_explorerPaneFallbackPending = false;
                bool m_explorerPaneFallbackUsed = false;
                UINT_PTR m_explorerPaneFallbackTimerId = 0;
                bool m_loggedExplorerPanesReady = false;
                bool m_loggedListViewMissing = false;
                bool m_loggedTreeViewMissing = false;
                PaneHookRouter m_paneHooks;
                ExplorerGlowCoordinator m_glowCoordinator;
                Microsoft::WRL::ComPtr<INameSpaceTreeControl> m_namespaceTreeControl;
                std::unique_ptr<NamespaceTreeHost> m_namespaceTreeHost;
                struct FolderBackgroundEntryData {
                        std::wstring imagePath;
                        std::wstring folderDisplayPath;
                };
                bool m_folderBackgroundsEnabled = false;
                std::unordered_map<std::wstring, FolderBackgroundEntryData> m_folderBackgroundEntries;
                mutable std::unordered_map<std::wstring, std::unique_ptr<Gdiplus::Bitmap>> m_folderBackgroundBitmaps;
                mutable std::wstring m_universalBackgroundImagePath;
                mutable std::unique_ptr<Gdiplus::Bitmap> m_universalBackgroundBitmap;
                mutable std::unordered_set<std::wstring> m_failedBackgroundKeys;
                std::wstring m_currentFolderKey;
                ListViewBackgroundSurface m_listViewBackgroundSurface;
                ListViewBackgroundJobState m_listViewBackgroundJobState;
                std::jthread m_listViewBackgroundWorker;
                std::atomic<uint64_t> m_listViewBackgroundGeneration{0};
                HMENU m_trackedContextMenu = nullptr;
                std::vector<std::wstring> m_pendingOpenInNewTabPaths;
                std::vector<std::wstring> m_openInNewTabQueue;
                std::unordered_map<HWND, BandEnsureState, HandleHasher> m_bandEnsureStates;
                std::unordered_map<HWND, ListViewAccentResources, HandleHasher> m_listViewAccentCache;
                bool m_useExplorerAccentColors = true;
                std::vector<ContextMenuItem> m_cachedContextMenuItems;
                ContextMenuSelectionSnapshot m_contextMenuSelection;
                std::unordered_map<UINT, const ContextMenuItem*> m_contextMenuCommandMap;
                std::vector<IconCache::Reference> m_contextMenuIconRefs;
                std::vector<HBITMAP> m_contextMenuBitmaps;
                std::vector<HMENU> m_contextMenuSubmenus;
                UINT m_nextContextCommandId = 0;

                static std::mutex s_ensureTimerLock;
                static std::unordered_map<UINT_PTR, CExplorerBHO*> s_ensureTimers;
                static std::mutex s_openInNewTabTimerLock;
                static std::unordered_map<UINT_PTR, CExplorerBHO*> s_openInNewTabTimers;
                UINT_PTR m_openInNewTabTimerId = 0;
                bool m_openInNewTabRetryScheduled = false;
                bool m_contextMenuInserted = false;
                static constexpr UINT kOpenInNewTabCommandId = 0xE170;
                static constexpr UINT kCustomCommandIdBase = 0xE200;
        };

}  // namespace shelltabs

