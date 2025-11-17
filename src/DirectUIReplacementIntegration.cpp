#include "DirectUIReplacementIntegration.h"
#include "CustomFileListView.h"
#include "Logging.h"
#include "OptionsStore.h"
#include "StringUtils.h"
#include "Utilities.h"
#include <mutex>
#include <optional>
#include <string>
#include <winreg.h>

namespace shelltabs {

// Static member initialization
bool DirectUIReplacementIntegration::s_initialized = false;
bool DirectUIReplacementIntegration::s_enabled = false;
bool DirectUIReplacementIntegration::s_featureOptedIn = false;
unsigned int DirectUIReplacementIntegration::s_activeHostCount = 0;
void (*DirectUIReplacementIntegration::s_viewCreatedCallback)(ShellTabs::CustomFileListView*, HWND, void*) = nullptr;
void* DirectUIReplacementIntegration::s_viewCreatedContext = nullptr;

namespace {
    std::mutex g_initMutex;
    HINSTANCE g_hInstance = nullptr;
    std::once_flag g_faultMitigationRegistration;
    std::mutex g_hostMutex;
    constexpr wchar_t kOptInEnvironmentVariable[] = L"SHELLTABS_ENABLE_DIRECTUI_REPLACEMENT";
    constexpr wchar_t kOptInRegistryPath[] = L"Software\\ShellTabs";
    constexpr wchar_t kOptInRegistryValue[] = L"EnableDirectUIReplacement";

    void DisableDirectUIReplacementAfterFault(const FaultMitigationDetails& details, void*) {
        LogMessage(LogLevel::Warning,
                   L"DirectUIReplacementIntegration: disabling hooks due to ShellTabs fault %llu (code=0x%08X)",
                   details.faultId,
                   details.exceptionCode);
        DirectUIReplacementIntegration::SetEnabled(false);
    }

    void EnsureFaultMitigationRegistration() {
        std::call_once(g_faultMitigationRegistration, [] {
            RegisterFaultMitigationHandler(&DisableDirectUIReplacementAfterFault, nullptr);
        });
    }

    std::optional<bool> ReadEnvironmentOptInOverride() {
        DWORD required = GetEnvironmentVariableW(kOptInEnvironmentVariable, nullptr, 0);
        if (required == 0) {
            return std::nullopt;
        }

        std::wstring buffer(required, L'\0');
        DWORD copied = GetEnvironmentVariableW(kOptInEnvironmentVariable, buffer.data(), required);
        if (copied == 0 || copied >= required) {
            return std::nullopt;
        }

        buffer.resize(copied);
        std::wstring_view trimmed = TrimView(buffer);
        return ParseBool(trimmed);
    }

    std::optional<bool> ReadRegistryOptInOverride() {
        DWORD value = 0;
        DWORD size = sizeof(value);
        const LONG status = RegGetValueW(HKEY_CURRENT_USER,
                                         kOptInRegistryPath,
                                         kOptInRegistryValue,
                                         RRF_RT_REG_DWORD,
                                         nullptr,
                                         &value,
                                         &size);
        if (status != ERROR_SUCCESS) {
            return std::nullopt;
        }
        return value != 0;
    }

    bool EvaluateDirectUiReplacementOptIn() {
        bool enabled = OptionsStore::Instance().Get().enableDirectUiReplacement;

        if (auto envOverride = ReadEnvironmentOptInOverride()) {
            LogMessage(LogLevel::Info,
                       L"DirectUIReplacementIntegration: environment override %ls=%d",
                       kOptInEnvironmentVariable,
                       *envOverride ? 1 : 0);
            return *envOverride;
        }

        if (auto registryOverride = ReadRegistryOptInOverride()) {
            LogMessage(LogLevel::Info,
                       L"DirectUIReplacementIntegration: registry override %ls\\%ls=%d",
                       kOptInRegistryPath,
                       kOptInRegistryValue,
                       *registryOverride ? 1 : 0);
            enabled = *registryOverride;
        }

        return enabled;
    }
}

bool DirectUIReplacementIntegration::Initialize() {
    std::lock_guard<std::mutex> lock(g_initMutex);

    if (s_initialized) {
        return true;
    }

    EnsureFaultMitigationRegistration();

    s_featureOptedIn = EvaluateDirectUiReplacementOptIn();
    if (!s_featureOptedIn) {
        LogMessage(LogLevel::Info,
                   L"DirectUIReplacementIntegration: replacement disabled (opt-in not enabled)");
        s_enabled = false;
        s_initialized = true;
        return true;
    }

    s_enabled = true;

    // Get module instance
    g_hInstance = GetModuleHandleW(nullptr);
    if (!g_hInstance) {
        g_hInstance = GetModuleHandleW(L"ShellTabs.dll");
    }

    // Register custom window class
    if (!ShellTabs::CustomFileListView::RegisterWindowClass(g_hInstance)) {
        return false;
    }

    // Initialize the hook system (only if enabled)
    if (s_enabled) {
        if (!ShellTabs::DirectUIReplacementHook::Initialize()) {
            return false;
        }
    }

    s_initialized = true;
    return true;
}

void DirectUIReplacementIntegration::Shutdown() {
    std::lock_guard<std::mutex> lock(g_initMutex);

    if (!s_initialized) {
        return;
    }

    if (s_featureOptedIn) {
        ShellTabs::DirectUIReplacementHook::Shutdown();
    }

    s_initialized = false;
    s_viewCreatedCallback = nullptr;
    s_viewCreatedContext = nullptr;
}

bool DirectUIReplacementIntegration::IsEnabled() {
    return s_featureOptedIn && s_enabled && s_initialized;
}

void DirectUIReplacementIntegration::SetEnabled(bool enabled) {
    std::lock_guard<std::mutex> lock(g_initMutex);

    if (enabled && !s_featureOptedIn) {
        LogMessage(LogLevel::Info,
                   L"DirectUIReplacementIntegration: ignoring enable request (feature not opted in)");
        return;
    }

    if (s_enabled == enabled) {
        return;
    }

    if (enabled && HasFaultMitigationTriggered()) {
        LogMessage(LogLevel::Warning,
                   L"DirectUIReplacementIntegration: refusing to re-enable after crash mitigation was triggered");
        return;
    }

    s_enabled = enabled;

    // If we're already initialized, update hook state
    if (s_initialized) {
        if (enabled) {
            ShellTabs::DirectUIReplacementHook::Initialize();
        } else {
            ShellTabs::DirectUIReplacementHook::Shutdown();
        }
    }
}

void DirectUIReplacementIntegration::RegisterHost(void* context) {
    UNREFERENCED_PARAMETER(context);
    std::lock_guard<std::mutex> lock(g_hostMutex);
    ++s_activeHostCount;
}

void DirectUIReplacementIntegration::UnregisterHost(void* context) {
    UNREFERENCED_PARAMETER(context);
    std::lock_guard<std::mutex> lock(g_hostMutex);
    if (s_activeHostCount > 0) {
        --s_activeHostCount;
    }
}

ShellTabs::CustomFileListView* DirectUIReplacementIntegration::GetCustomViewForWindow(HWND hwnd) {
    if (!s_initialized || !hwnd) {
        return nullptr;
    }

    return ShellTabs::DirectUIReplacementHook::GetInstance(hwnd);
}

void DirectUIReplacementIntegration::SetCustomViewCreatedCallback(
    void (*callback)(ShellTabs::CustomFileListView* view, HWND hwnd, void* context),
    void* context) {
    std::lock_guard<std::mutex> lock(g_initMutex);
    s_viewCreatedCallback = callback;
    s_viewCreatedContext = context;
}

void DirectUIReplacementIntegration::ClearCustomViewCreatedCallback(void* context) {
    std::lock_guard<std::mutex> lock(g_initMutex);
    if (s_viewCreatedContext == context) {
        s_viewCreatedCallback = nullptr;
        s_viewCreatedContext = nullptr;
    }
}

bool DirectUIReplacementIntegration::CanCreateCustomView() {
    if (!s_initialized || !s_enabled || !s_viewCreatedCallback || !s_featureOptedIn) {
        return false;
    }

    std::lock_guard<std::mutex> lock(g_hostMutex);
    return s_activeHostCount > 0;
}

void DirectUIReplacementIntegration::NotifyViewCreated(
    ShellTabs::CustomFileListView* view, HWND hwnd) {
    if (s_viewCreatedCallback) {
        s_viewCreatedCallback(view, hwnd, s_viewCreatedContext);
    }
}

bool IsDirectUIWindow(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    wchar_t className[256] = {};
    if (GetClassNameW(hwnd, className, _countof(className)) == 0) {
        return false;
    }

    // Check for native DirectUIHWND or our custom replacement
    return wcscmp(className, L"DirectUIHWND") == 0 ||
           wcscmp(className, L"ShellTabsFileListView") == 0;
}

HWND FindDirectUIWindow(HWND shellViewWindow) {
    if (!shellViewWindow || !IsWindow(shellViewWindow)) {
        return nullptr;
    }

    // First try to find our custom window
    HWND customView = FindDescendantWindow(shellViewWindow, L"ShellTabsFileListView");
    if (customView) {
        return customView;
    }

    // Fallback to finding native DirectUIHWND
    return FindDescendantWindow(shellViewWindow, L"DirectUIHWND");
}

} // namespace shelltabs
