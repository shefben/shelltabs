#include "TaskbarPreviewHook.h"

#include "TaskbarTabListPopup.h"
#include "Utilities.h"
#include "Logging.h"
#include "Module.h"

#include <wrl/client.h>

namespace shelltabs {

TaskbarPreviewHook& TaskbarPreviewHook::Instance() {
    static TaskbarPreviewHook instance;
    return instance;
}

TaskbarPreviewHook::~TaskbarPreviewHook() {
    Shutdown();
}

void TaskbarPreviewHook::Initialize(std::unique_ptr<ITaskbarTabProvider> provider) {
    if (m_initialized) return;
    m_provider = std::move(provider);
    m_popup = std::make_unique<TaskbarTabListPopup>();
    m_popup->Create();
    m_popup->SetTabActivatedCallback(
        [this](HWND explorerHwnd, int groupIndex, int tabIndex) {
            if (m_provider) {
                m_provider->ActivateTab(explorerHwnd, groupIndex, tabIndex);
            }
            if (m_popup) {
                m_popup->Hide();
            }
        });

    m_eventHook = SetWinEventHook(
        EVENT_OBJECT_SHOW, EVENT_OBJECT_HIDE,
        GetModuleHandleInstance(),
        WinEventProc,
        GetCurrentProcessId(), 0,
        WINEVENT_INCONTEXT);

    if (!m_eventHook) {
        LogLastError(L"TaskbarPreviewHook::SetWinEventHook", GetLastError());
    }

    m_initialized = true;
    LogMessage(LogLevel::Info, L"TaskbarPreviewHook initialized");
}

void TaskbarPreviewHook::Shutdown() {
    if (!m_initialized) return;

    if (m_pollTimer) {
        KillTimer(nullptr, m_pollTimer);
        m_pollTimer = 0;
    }
    if (m_hideTimer) {
        KillTimer(nullptr, m_hideTimer);
        m_hideTimer = 0;
    }
    if (m_eventHook) {
        UnhookWinEvent(m_eventHook);
        m_eventHook = nullptr;
    }
    if (m_popup) {
        m_popup->Destroy();
        m_popup.reset();
    }
    if (m_provider) {
        m_provider->ClearCachedResources();
        m_provider.reset();
    }
    m_activeThumbnailWnd = nullptr;
    m_currentExplorerHwnd = nullptr;
    m_explorerWindows.clear();
    m_initialized = false;

    LogMessage(LogLevel::Info, L"TaskbarPreviewHook shut down");
}

void CALLBACK TaskbarPreviewHook::WinEventProc(
    HWINEVENTHOOK /*hook*/, DWORD event, HWND hwnd,
    LONG idObject, LONG /*idChild*/,
    DWORD /*eventThread*/, DWORD /*eventTime*/) {
    if (idObject != OBJID_WINDOW || !hwnd) return;

    if (!MatchesClass(hwnd, L"TaskListThumbnailWnd")) return;

    auto& self = Instance();
    if (event == EVENT_OBJECT_SHOW) {
        self.OnPreviewPopupShown(hwnd);
    } else if (event == EVENT_OBJECT_HIDE) {
        self.OnPreviewPopupHidden(hwnd);
    }
}

void CALLBACK TaskbarPreviewHook::PollTimerProc(
    HWND /*hwnd*/, UINT /*msg*/, UINT_PTR /*id*/, DWORD /*time*/) {
    Instance().OnPollTick();
}

void CALLBACK TaskbarPreviewHook::HideDebounceProc(
    HWND /*hwnd*/, UINT /*msg*/, UINT_PTR /*id*/, DWORD /*time*/) {
    Instance().OnHideDebounce();
}

void TaskbarPreviewHook::OnPreviewPopupShown(HWND thumbnailWnd) {
    m_activeThumbnailWnd = thumbnailWnd;
    RefreshExplorerWindowList();

    // Kill any pending hide debounce
    if (m_hideTimer) {
        KillTimer(nullptr, m_hideTimer);
        m_hideTimer = 0;
    }

    // Start polling
    if (!m_pollTimer) {
        m_pollTimer = SetTimer(nullptr, kPollTimerId, kPollIntervalMs, PollTimerProc);
    }
}

void TaskbarPreviewHook::OnPreviewPopupHidden(HWND thumbnailWnd) {
    if (m_activeThumbnailWnd != thumbnailWnd) return;

    if (m_pollTimer) {
        KillTimer(nullptr, m_pollTimer);
        m_pollTimer = 0;
    }
    if (m_hideTimer) {
        KillTimer(nullptr, m_hideTimer);
        m_hideTimer = 0;
    }
    if (m_popup) {
        m_popup->Hide();
    }
    m_activeThumbnailWnd = nullptr;
    m_currentExplorerHwnd = nullptr;
}

void TaskbarPreviewHook::OnPollTick() {
    if (!m_activeThumbnailWnd || !IsWindowVisible(m_activeThumbnailWnd)) {
        // Thumbnail window gone — clean up
        OnPreviewPopupHidden(m_activeThumbnailWnd);
        return;
    }

    POINT cursorPos;
    GetCursorPos(&cursorPos);

    // If cursor is over our popup, keep it visible
    if (m_popup && m_popup->IsVisible()) {
        HWND popupHwnd = m_popup->GetHwnd();
        if (popupHwnd) {
            RECT popupRect;
            GetWindowRect(popupHwnd, &popupRect);
            if (PtInRect(&popupRect, cursorPos)) {
                // Kill any pending hide debounce — cursor is on popup
                if (m_hideTimer) {
                    KillTimer(nullptr, m_hideTimer);
                    m_hideTimer = 0;
                }
                return;
            }
        }
    }

    // Check if cursor is over the thumbnail window at all
    RECT thumbnailRect;
    GetWindowRect(m_activeThumbnailWnd, &thumbnailRect);
    if (!PtInRect(&thumbnailRect, cursorPos)) {
        // Cursor left thumbnail area — start hide debounce (if popup is visible)
        if (m_popup && m_popup->IsVisible() && !m_hideTimer) {
            m_hideTimer = SetTimer(nullptr, kHideTimerId, kHideDebounceMs, HideDebounceProc);
        }
        return;
    }

    // Kill hide debounce — cursor is back over thumbnails
    if (m_hideTimer) {
        KillTimer(nullptr, m_hideTimer);
        m_hideTimer = 0;
    }

    auto info = HitTestThumbnail(m_activeThumbnailWnd, cursorPos);
    if (!info.explorerHwnd) {
        // Between thumbnails — start hide debounce
        if (m_popup && m_popup->IsVisible() && !m_hideTimer) {
            m_hideTimer = SetTimer(nullptr, kHideTimerId, kHideDebounceMs, HideDebounceProc);
        }
        return;
    }

    // Check if explorer window is still valid
    if (!IsWindow(info.explorerHwnd)) {
        if (m_popup) m_popup->Hide();
        m_currentExplorerHwnd = nullptr;
        return;
    }

    // Same window as before — no update needed
    if (info.explorerHwnd == m_currentExplorerHwnd && m_popup && m_popup->IsVisible()) {
        return;
    }

    // Check if this is a multi-tab window
    if (!m_provider || !m_provider->IsMultiTabWindow(info.explorerHwnd)) {
        if (m_popup) m_popup->Hide();
        m_currentExplorerHwnd = nullptr;
        return;
    }

    // Query tabs and show popup
    auto snapshot = m_provider->QueryTabs(info.explorerHwnd);
    if (snapshot.tabs.size() <= 1) {
        if (m_popup) m_popup->Hide();
        m_currentExplorerHwnd = nullptr;
        return;
    }

    m_currentExplorerHwnd = info.explorerHwnd;

    // Find the taskbar window for positioning
    HWND taskbarHwnd = FindWindowW(L"Shell_TrayWnd", nullptr);

    if (m_popup) {
        m_popup->ShowForWindow(snapshot, info.thumbnailRect, taskbarHwnd);
    }
}

void TaskbarPreviewHook::OnHideDebounce() {
    if (m_hideTimer) {
        KillTimer(nullptr, m_hideTimer);
        m_hideTimer = 0;
    }

    // Double-check: if cursor moved back over popup or thumbnail, don't hide
    POINT cursorPos;
    GetCursorPos(&cursorPos);

    if (m_popup && m_popup->IsVisible()) {
        HWND popupHwnd = m_popup->GetHwnd();
        if (popupHwnd) {
            RECT popupRect;
            GetWindowRect(popupHwnd, &popupRect);
            if (PtInRect(&popupRect, cursorPos)) return;
        }
    }

    if (m_activeThumbnailWnd && IsWindowVisible(m_activeThumbnailWnd)) {
        RECT thumbnailRect;
        GetWindowRect(m_activeThumbnailWnd, &thumbnailRect);
        if (PtInRect(&thumbnailRect, cursorPos)) return;
    }

    if (m_popup) m_popup->Hide();
    m_currentExplorerHwnd = nullptr;
}

TaskbarPreviewHook::ThumbnailInfo TaskbarPreviewHook::HitTestThumbnail(
    HWND thumbnailWnd, POINT screenPt) {
    Microsoft::WRL::ComPtr<IAccessible> acc;
    HRESULT hr = AccessibleObjectFromWindow(
        thumbnailWnd, static_cast<DWORD>(OBJID_CLIENT), IID_PPV_ARGS(&acc));
    if (FAILED(hr) || !acc) return {};

    VARIANT varChild{};
    hr = acc->accHitTest(screenPt.x, screenPt.y, &varChild);
    if (FAILED(hr) || varChild.vt != VT_I4 || varChild.lVal == CHILDID_SELF) {
        VariantClear(&varChild);
        return {};
    }

    VARIANT varId;
    varId.vt = VT_I4;
    varId.lVal = varChild.lVal;

    BSTR name = nullptr;
    hr = acc->get_accName(varId, &name);
    if (FAILED(hr) || !name) return {};
    std::wstring title(name);
    SysFreeString(name);

    LONG left = 0, top = 0, width = 0, height = 0;
    hr = acc->accLocation(&left, &top, &width, &height, varId);
    if (FAILED(hr)) return {};

    ThumbnailInfo info;
    info.windowTitle = std::move(title);
    info.thumbnailRect = {left, top, left + width, top + height};
    info.explorerHwnd = MatchTitleToExplorerWindow(info.windowTitle);
    return info;
}

HWND TaskbarPreviewHook::MatchTitleToExplorerWindow(const std::wstring& title) {
    // Search cached list first (case-insensitive)
    for (const auto& entry : m_explorerWindows) {
        if (_wcsicmp(entry.title.c_str(), title.c_str()) == 0) {
            return entry.hwnd;
        }
    }

    // Fallback: direct FindWindow
    HWND hwnd = FindWindowW(L"CabinetWClass", title.c_str());
    return hwnd;
}

void TaskbarPreviewHook::RefreshExplorerWindowList() {
    m_explorerWindows.clear();
    DWORD pid = GetCurrentProcessId();

    EnumWindows(
        [](HWND hwnd, LPARAM lParam) -> BOOL {
            auto* self = reinterpret_cast<TaskbarPreviewHook*>(lParam);
            DWORD windowPid = 0;
            GetWindowThreadProcessId(hwnd, &windowPid);
            if (windowPid != GetCurrentProcessId()) return TRUE;
            if (!MatchesClass(hwnd, L"CabinetWClass")) return TRUE;
            if (!IsWindowVisible(hwnd)) return TRUE;

            wchar_t title[256] = {};
            GetWindowTextW(hwnd, title, 256);
            self->m_explorerWindows.push_back({hwnd, title});
            return TRUE;
        },
        reinterpret_cast<LPARAM>(this));

    (void)pid;  // used inside lambda via GetCurrentProcessId()
}

}  // namespace shelltabs
