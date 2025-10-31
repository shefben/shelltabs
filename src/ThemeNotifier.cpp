#include "ThemeNotifier.h"

#include <windows.ui.viewmanagement.h>
#include <windows.foundation.h>
#include <windows.foundation.collections.h>
#include <eventtoken.h>
#include <inspectable.h>
#include <roapi.h>
#include <roerrorapi.h>
#include <wrl/client.h>
#include <wrl/event.h>
#include <wrl/wrappers/corewrappers.h>
#include <wtsapi32.h>
#include <utility>

#include "Logging.h"

#pragma comment(lib, "runtimeobject.lib")
#pragma comment(lib, "wtsapi32.lib")

using Microsoft::WRL::ComPtr;
using Microsoft::WRL::Wrappers::HStringReference;

using ABI::Windows::Foundation::IInspectable;
using ABI::Windows::Foundation::ITypedEventHandler;
using ABI::Windows::UI::Color;
using ABI::Windows::UI::ViewManagement::IUISettings3;
using ABI::Windows::UI::ViewManagement::UIColorType;
using ABI::Windows::UI::ViewManagement::UISettings;

namespace shelltabs {

namespace {

class UiSettingsEventHandler
    : public Microsoft::WRL::RuntimeClass<Microsoft::WRL::RuntimeClassFlags<Microsoft::WRL::ClassicCom>,
                                          ITypedEventHandler<UISettings*, IInspectable*>> {
public:
    using CallbackType = std::function<void()>;

    explicit UiSettingsEventHandler(CallbackType callback) : m_callback(std::move(callback)) {}

    HRESULT RuntimeClassInitialize(CallbackType callback) {
        m_callback = std::move(callback);
        return S_OK;
    }

    IFACEMETHODIMP Invoke(UISettings*, IInspectable*) override {
        if (m_callback) {
            m_callback();
        }
        return S_OK;
    }

private:
    CallbackType m_callback;
};

bool ShouldHandleSessionEvent(DWORD sessionEvent) {
    switch (sessionEvent) {
        case WTS_SESSION_LOGON:
        case WTS_SESSION_UNLOCK:
        case WTS_SESSION_REMOTE_CONTROL:
#ifdef WTS_SESSION_REMOTE_CONNECT
        case WTS_SESSION_REMOTE_CONNECT:
#endif
#ifdef WTS_SESSION_REMOTE_DISCONNECT
        case WTS_SESSION_REMOTE_DISCONNECT:
#endif
        case WTS_CONSOLE_CONNECT:
        case WTS_CONSOLE_DISCONNECT:
            return true;
        default:
            return false;
    }
}

}  // namespace

struct ThemeNotifier::UiSettingsState {
    ComPtr<IUISettings3> uiSettings3;
    ComPtr<UiSettingsEventHandler> handler;
    EventRegistrationToken colorToken{};
    bool winrtInitialized = false;
};

ThemeNotifier::ThemeNotifier() = default;

ThemeNotifier::~ThemeNotifier() { Shutdown(); }

bool ThemeNotifier::Initialize(HWND window, std::function<void()> callback) {
    Shutdown();

    m_window = window;
    m_callback = std::move(callback);

    if (m_window && !m_callback) {
        return false;
    }

    if (m_window) {
        if (WTSRegisterSessionNotification(m_window, NOTIFY_FOR_THIS_SESSION)) {
            m_wtsRegistered = true;
        } else {
            LogLastError(L"WTSRegisterSessionNotification", GetLastError());
        }
    }

    auto state = std::make_unique<UiSettingsState>();
    const HRESULT initResult = RoInitialize(RO_INIT_MULTITHREADED);
    if (SUCCEEDED(initResult) || initResult == S_FALSE || initResult == RO_E_ALREADYINITIALIZED) {
        state->winrtInitialized = SUCCEEDED(initResult) || initResult == S_FALSE;

        HStringReference classId(RuntimeClass_Windows_UI_ViewManagement_UISettings);
        ComPtr<IInspectable> inspectable;
        HRESULT hr = RoActivateInstance(classId.Get(), &inspectable);
        if (SUCCEEDED(hr)) {
            hr = inspectable.As(&state->uiSettings3);
            if (FAILED(hr)) {
                LogMessage(LogLevel::Warning,
                           L"ThemeNotifier: UISettings3 not available (hr=0x%08X)", hr);
            }
        } else {
            LogMessage(LogLevel::Warning, L"ThemeNotifier: failed to activate UISettings (hr=0x%08X)", hr);
        }

        if (state->uiSettings3) {
            state->handler = Microsoft::WRL::Make<UiSettingsEventHandler>([this]() {
                UpdateColorSnapshot();
                NotifyThemeChanged();
            });
            if (state->handler) {
                hr = state->uiSettings3->add_ColorValuesChanged(state->handler.Get(), &state->colorToken);
                if (FAILED(hr)) {
                    LogMessage(LogLevel::Warning,
                               L"ThemeNotifier: failed to subscribe to UISettings3 (hr=0x%08X)", hr);
                    state->handler.Reset();
                }
            }
            UpdateColorSnapshot();
        }
    } else {
        LogMessage(LogLevel::Warning, L"ThemeNotifier: RoInitialize failed (hr=0x%08X)", initResult);
    }

    m_uiSettings = std::move(state);

    if (!m_cachedColors.valid) {
        UpdateColorSnapshot();
    }

    if (m_callback) {
        NotifyThemeChanged();
    }

    return true;
}

void ThemeNotifier::Shutdown() {
    if (m_wtsRegistered && m_window) {
        WTSUnRegisterSessionNotification(m_window);
    }
    m_wtsRegistered = false;

    if (m_uiSettings) {
        if (m_uiSettings->uiSettings3 && m_uiSettings->handler) {
            m_uiSettings->uiSettings3->remove_ColorValuesChanged(m_uiSettings->colorToken);
        }
        m_uiSettings->handler.Reset();
        m_uiSettings->uiSettings3.Reset();
        if (m_uiSettings->winrtInitialized) {
            RoUninitialize();
        }
    }

    m_uiSettings.reset();
    m_callback = nullptr;
    m_window = nullptr;

    std::lock_guard<std::mutex> guard(m_mutex);
    m_cachedColors = {};
}

ThemeNotifier::ThemeColors ThemeNotifier::GetThemeColors() const {
    std::lock_guard<std::mutex> guard(m_mutex);
    return m_cachedColors;
}

bool ThemeNotifier::HandleSessionChange(WPARAM sessionEvent, LPARAM) {
    if (!ShouldHandleSessionEvent(static_cast<DWORD>(sessionEvent))) {
        return false;
    }
    UpdateColorSnapshot();
    NotifyThemeChanged();
    return true;
}

void ThemeNotifier::NotifyThemeChanged() {
    if (m_callback) {
        m_callback();
    }
}

void ThemeNotifier::UpdateColorSnapshot() {
    ThemeColors snapshot;

    if (m_uiSettings && m_uiSettings->uiSettings3) {
        Color color{};
        if (SUCCEEDED(m_uiSettings->uiSettings3->GetColorValue(UIColorType_Background, &color))) {
            snapshot.background = RGB(color.R, color.G, color.B);
            snapshot.valid = true;
        }
        if (SUCCEEDED(m_uiSettings->uiSettings3->GetColorValue(UIColorType_Foreground, &color))) {
            snapshot.foreground = RGB(color.R, color.G, color.B);
            snapshot.valid = true;
        }
    }

    if (!snapshot.valid) {
        snapshot.background = GetSysColor(COLOR_WINDOW);
        snapshot.foreground = GetSysColor(COLOR_WINDOWTEXT);
    }

    std::lock_guard<std::mutex> guard(m_mutex);
    m_cachedColors = snapshot;
}

#if defined(SHELLTABS_ENABLE_THEME_TEST_HOOKS)
void ThemeNotifier::SimulateColorChangeForTest() {
    UpdateColorSnapshot();
    NotifyThemeChanged();
}

void ThemeNotifier::SimulateSessionEventForTest(DWORD sessionEvent) {
    if (ShouldHandleSessionEvent(sessionEvent)) {
        UpdateColorSnapshot();
        NotifyThemeChanged();
    }
}
#endif

}  // namespace shelltabs

