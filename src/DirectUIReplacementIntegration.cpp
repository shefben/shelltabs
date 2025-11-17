#include "DirectUIReplacementIntegration.h"
#include "CustomFileListView.h"
#include "Logging.h"
#include "Utilities.h"
#include <mutex>

namespace shelltabs {

// Static member initialization
bool DirectUIReplacementIntegration::s_initialized = false;
bool DirectUIReplacementIntegration::s_enabled = true;  // Enabled by default
unsigned int DirectUIReplacementIntegration::s_activeHostCount = 0;
void (*DirectUIReplacementIntegration::s_viewCreatedCallback)(ShellTabs::CustomFileListView*, HWND, void*) = nullptr;
void* DirectUIReplacementIntegration::s_viewCreatedContext = nullptr;

namespace {
    std::mutex g_initMutex;
    HINSTANCE g_hInstance = nullptr;
    std::once_flag g_faultMitigationRegistration;
    std::mutex g_hostMutex;

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
}

bool DirectUIReplacementIntegration::Initialize() {
    std::lock_guard<std::mutex> lock(g_initMutex);

    if (s_initialized) {
        return true;
    }

    EnsureFaultMitigationRegistration();

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

    ShellTabs::DirectUIReplacementHook::Shutdown();

    s_initialized = false;
    s_viewCreatedCallback = nullptr;
    s_viewCreatedContext = nullptr;
}

bool DirectUIReplacementIntegration::IsEnabled() {
    return s_enabled && s_initialized;
}

void DirectUIReplacementIntegration::SetEnabled(bool enabled) {
    std::lock_guard<std::mutex> lock(g_initMutex);

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
    if (!s_initialized || !s_enabled || !s_viewCreatedCallback) {
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
