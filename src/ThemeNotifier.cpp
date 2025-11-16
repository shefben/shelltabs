#include "ThemeNotifier.h"

#include <windows.ui.viewmanagement.h>
#include <windows.foundation.h>
#include <windows.foundation.collections.h>
#include <eventtoken.h>
#include <inspectable.h>
#include <roapi.h>
#include <winerror.h>
#include <wrl/client.h>
#include <wrl/event.h>
#include <wrl/wrappers/corewrappers.h>
#include <wtsapi32.h>
#include <functional>
#include <utility>

#include "Logging.h"

#pragma comment(lib, "runtimeobject.lib")
#pragma comment(lib, "wtsapi32.lib")

using Microsoft::WRL::ComPtr;
using Microsoft::WRL::Wrappers::HStringReference;

using ABI::Windows::UI::Color;
using ABI::Windows::UI::ViewManagement::IUISettings;
using ABI::Windows::UI::ViewManagement::IUISettings3;
using UiSettingsColorChangedHandler =
    ABI::Windows::Foundation::__FITypedEventHandler_2_Windows__CUI__CViewManagement__CUISettings_IInspectable_t;

namespace shelltabs {

namespace {

class UiSettingsEventHandler
    : public Microsoft::WRL::RuntimeClass<Microsoft::WRL::RuntimeClassFlags<Microsoft::WRL::ClassicCom>,
                                          UiSettingsColorChangedHandler> {
public:
    using CallbackType = std::function<void()>;

    explicit UiSettingsEventHandler(CallbackType callback) : m_callback(std::move(callback)) {}

    HRESULT RuntimeClassInitialize(CallbackType callback) {
        m_callback = std::move(callback);
        return S_OK;
    }

    IFACEMETHODIMP Invoke(IUISettings*, ::IInspectable*) override {
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
    enum class ApartmentModel {
        None,
        Multithreaded,
        SingleThreaded,
    };

    ComPtr<IUISettings3> uiSettings3;
    ComPtr<UiSettingsEventHandler> handler;
    EventRegistrationToken colorToken{};
    ApartmentModel apartmentModel = ApartmentModel::None;
};

ThemeNotifier::ThemeNotifier() = default;

ThemeNotifier::~ThemeNotifier() { Shutdown(); }

bool ThemeNotifier::Initialize(HWND window, std::function<void()> callback) {
    Shutdown();

    if (window && !callback) {
        return false;
    }

    auto state = std::make_unique<UiSettingsState>();

    const auto cleanupState = [&]() {
        if (state && state->apartmentModel != UiSettingsState::ApartmentModel::None) {
            RoUninitialize();
            state->apartmentModel = UiSettingsState::ApartmentModel::None;
        }
    };

    const auto tryInitialize = [&](RO_INIT_TYPE type, UiSettingsState::ApartmentModel model) -> HRESULT {
        const HRESULT hr = RoInitialize(type);
        if (SUCCEEDED(hr) || hr == S_FALSE || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED)) {
            state->apartmentModel = model;
        }
        return hr;
    };

    HRESULT initResult = tryInitialize(RO_INIT_MULTITHREADED, UiSettingsState::ApartmentModel::Multithreaded);
    if (state->apartmentModel == UiSettingsState::ApartmentModel::None && initResult == RPC_E_CHANGED_MODE) {
        initResult = tryInitialize(RO_INIT_SINGLETHREADED, UiSettingsState::ApartmentModel::SingleThreaded);
    }

    if (state->apartmentModel == UiSettingsState::ApartmentModel::None) {
        LogMessage(LogLevel::Warning,
                   L"ThemeNotifier: RoInitialize failed (hr=0x%08X)",
                   initResult);
        cleanupState();
        return false;
    }

    HStringReference classId(RuntimeClass_Windows_UI_ViewManagement_UISettings);
    ComPtr<::IInspectable> inspectable;
    HRESULT hr = RoActivateInstance(classId.Get(), &inspectable);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"ThemeNotifier: failed to activate UISettings (hr=0x%08X)", hr);
        cleanupState();
        return false;
    }

    hr = inspectable.As(&state->uiSettings3);
    if (FAILED(hr) || !state->uiSettings3) {
        LogMessage(LogLevel::Warning, L"ThemeNotifier: UISettings3 not available (hr=0x%08X)", hr);
        cleanupState();
        return false;
    }

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

    m_window = window;
    m_callback = std::move(callback);

    if (m_window) {
        if (WTSRegisterSessionNotification(m_window, NOTIFY_FOR_THIS_SESSION)) {
            m_wtsRegistered = true;
        } else {
            LogLastError(L"WTSRegisterSessionNotification", GetLastError());
        }
    }

    m_uiSettings = std::move(state);

    UpdateColorSnapshot();

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
        if (m_uiSettings->apartmentModel != UiSettingsState::ApartmentModel::None) {
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

void ThemeNotifier::RefreshColorsFromSystem() {
    UpdateColorSnapshot();
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
        constexpr auto kBackgroundColorType = ABI::Windows::UI::ViewManagement::UIColorType_Background;
        constexpr auto kForegroundColorType = ABI::Windows::UI::ViewManagement::UIColorType_Foreground;

        if (SUCCEEDED(m_uiSettings->uiSettings3->GetColorValue(kBackgroundColorType, &color))) {
            snapshot.background = RGB(color.R, color.G, color.B);
            snapshot.valid = true;
        }
        if (SUCCEEDED(m_uiSettings->uiSettings3->GetColorValue(kForegroundColorType, &color))) {
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

