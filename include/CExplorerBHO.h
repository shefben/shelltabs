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
#include "EditGradientRenderer.h"
#include "ExplorerThemeUtils.h"
#include "OptionsStore.h"
#include "Utilities.h"
#include "VisualPropertiesInterop.h"
#include "NameColorView.h"
#include "ShellTabsMessages.h"

namespace Gdiplus {
class Bitmap;
}

namespace shelltabs {

struct ShellTabsOptions;

        class CExplorerBHO : public IObjectWithSite,
                             public IDispatch {
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
                void ScheduleDeferredBreadcrumbUpdate();
                void CancelDeferredBreadcrumbUpdate();
                static void CALLBACK DeferredBreadcrumbTimerProc(HWND hwnd, UINT msg, UINT_PTR timerId, DWORD tickCount);
		void RemoveBreadcrumbSubclass();
		HWND FindBreadcrumbToolbar() const;
		HWND FindBreadcrumbToolbarInWindow(HWND root) const;
                HWND FindProgressWindow() const;
                HWND FindAddressEditControl() const;
                HWND GetTopLevelExplorerWindow() const;
                HWND GetShellTabsBandWindow() const;
                bool PostTravelToolbarNavigationMessage(bool navigateBack) const;
		bool InstallBreadcrumbSubclass(HWND toolbar);
                bool InstallProgressSubclass(HWND progressWindow);
                bool InstallAddressEditSubclass(HWND editWindow);
                void UpdateProgressSubclass();
                void UpdateTravelBandSubclass();
                bool InstallTravelBandSubclass(HWND travelBand, HWND toolbar);
                void RemoveTravelBandSubclass();
                void ResolveTravelToolbarCommands();
                bool HandleTravelToolbarMouseButton(HWND toolbar, bool buttonUp, WPARAM wParam,
                                                    LPARAM lParam, LRESULT* result);
                bool HandleTravelToolbarMouseActivate(LRESULT* result);
                bool HandleTravelBandNotify(NMHDR* header, LRESULT* result);
                bool HandleTravelBandDropdown(const NMTOOLBARW& info, LRESULT* result);
                bool ShowTravelHistoryMenu(HistoryMenuKind kind, const RECT& buttonRect, LRESULT* result);
                void ResetTravelToolbarButtonState();
                void SetTravelToolbarButtonPressed(UINT commandId, bool pressed);
                void SetTravelToolbarButtonEnabled(UINT commandId, bool enabled);
                bool IsTravelToolbarButtonEnabled(UINT commandId) const;
                void UpdateTravelToolbarButtonStates();
                void BeginTravelToolbarCapture(HWND toolbar);
                void ReleaseTravelToolbarCapture();
                void RemoveProgressSubclass();
                void UpdateAddressEditSubclass();
                void RemoveAddressEditSubclass();
                std::vector<HWND> FindAllStatusBars() const;
                void UpdateStatusBarTheme();
                void RemoveStatusBarSubclass();
                bool HandleStatusBarPaint(HWND hwnd);
                bool HandleRebarPaint(HWND hwnd);
                void RequestAddressEditRedraw(HWND hwnd) const;
                void ResetAddressEditStateCache();
                bool RefreshAddressEditState(HWND hwnd, bool updateText, bool updateSelection,
                                             bool updateFocus, bool updateTheme);
                bool RefreshAddressEditText(HWND hwnd);
                bool RefreshAddressEditSelection(HWND hwnd);
                bool RefreshAddressEditFocus(HWND hwnd);
                bool RefreshAddressEditTheme();
                bool RefreshAddressEditFont(HWND hwnd);
                bool IsExplorerEditAncestor(HWND hwnd) const;
                void MaybeAddExplorerEdit(HWND candidate,
                                           std::unordered_set<HWND, HandleHasher>& seen,
                                           std::vector<HWND>& edits) const;
                bool EnsureProgressGradientResources();
                void DestroyProgressGradientResources();
                void ReloadFolderBackgrounds(const ShellTabsOptions& options);
                void ClearFolderBackgrounds();
                std::wstring NormalizeBackgroundKey(const std::wstring& path) const;
                bool EnsureFolderBackgroundBitmap(const std::wstring& key) const;
                bool EnsureUniversalBackgroundBitmap() const;
                void UpdateCurrentFolderBackground();
                void ApplyCurrentFolderViewRules(const ShellTabsOptions& options);
                void InvalidateFolderBackgroundTargets() const;
                void UpdateNameColorView();
                void EnsureBreadcrumbHook();
                void RemoveBreadcrumbHook();
                bool IsBreadcrumbToolbarCandidate(HWND hwnd) const;
                bool IsBreadcrumbToolbarAncestor(HWND hwnd) const;
                bool IsWindowOwnedByThisExplorer(HWND hwnd) const;
                bool HandleBreadcrumbPaint(HWND hwnd);
                bool HandleProgressPaint(HWND hwnd);
                void PaintAddressEditOverlay(HWND hwnd, HDC dc, const RECT* clip = nullptr);
                bool DrawAddressEditContent(HWND hwnd, HDC dc, GradientEditRenderOptions options = {});
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
                static LRESULT CALLBACK TravelBandSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                        UINT_PTR subclassId, DWORD_PTR refData);
                static LRESULT CALLBACK TravelToolbarSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                        UINT_PTR subclassId, DWORD_PTR refData);
                static LRESULT CALLBACK StatusBarSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                        UINT_PTR subclassId, DWORD_PTR refData);
                static LRESULT CALLBACK ExplorerFrameSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                        UINT_PTR subclassId, DWORD_PTR refData);
                static LRESULT CALLBACK DirectUISubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                        UINT_PTR subclassId, DWORD_PTR refData);
                static bool ParseAddressEditPath(HWND editControl, std::wstring* outPath);
                void InstallExplorerFrameSubclass();
                void RemoveExplorerFrameSubclass();
                void InstallDirectUISubclass();
                void RemoveDirectUISubclass();

                std::atomic<long> m_refCount;
		Microsoft::WRL::ComPtr<IUnknown> m_site;
		Microsoft::WRL::ComPtr<IWebBrowser2> m_webBrowser;
		Microsoft::WRL::ComPtr<IShellBrowser> m_shellBrowser;
		Microsoft::WRL::ComPtr<IConnectionPoint> m_connectionPoint;
		DWORD m_connectionCookie = 0;
		bool m_bandVisible = false;
		bool m_shouldRetryEnsure = true;
		bool m_isDisconnecting = false;
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
                HWND m_travelBand = nullptr;
                HWND m_travelToolbar = nullptr;
                bool m_travelBandSubclassInstalled = false;
                bool m_travelToolbarSubclassInstalled = false;
                bool m_travelToolbarMouseCaptured = false;
                bool m_travelHistoryMenuVisible = false;
                int m_travelToolbarPressedButton = -1;
                UINT m_travelBackCommandId = 0;
                UINT m_travelForwardCommandId = 0;
                UINT m_travelHistoryDropdownCommandId = 0;
                HWND m_explorerFrame = nullptr;
                bool m_explorerFrameSubclassInstalled = false;
                HWND m_directUIHwnd = nullptr;
                std::unique_ptr<NameColorView> m_nameColorView;
                bool m_directUISubclassInstalled = false;
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
                HWND m_statusBar = nullptr;
                bool m_statusBarSubclassInstalled = false;
                bool m_statusBarDarkModeApplied = false;
                std::vector<HWND> m_statusBars;       // all msctls_statusbar32 in the frame
                std::vector<HWND> m_statusBarRebars;  // rebars / containers above each status bar
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
                ExplorerGlowCoordinator m_glowCoordinator;
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
                HBITMAP m_currentBackgroundBitmap = nullptr;
                std::unordered_map<HWND, BandEnsureState, HandleHasher> m_bandEnsureStates;
                bool m_useExplorerAccentColors = true;
                std::vector<ContextMenuItem> m_cachedContextMenuItems;

                static std::mutex s_ensureTimerLock;
                static std::unordered_map<UINT_PTR, CExplorerBHO*> s_ensureTimers;
                static std::mutex s_deferredBreadcrumbTimerLock;
                static std::unordered_map<UINT_PTR, CExplorerBHO*> s_deferredBreadcrumbTimers;
                UINT_PTR m_deferredBreadcrumbTimerId = 0;
                static constexpr ULONGLONG kCustomDrawTimeoutMs = 2000;
        };

}  // namespace shelltabs
