#include "MinHookSupport.h"

#include "Logging.h"
#include "MinHook.h"

#include <mutex>

namespace shelltabs {

namespace {

std::mutex g_minHookMutex;
unsigned int g_minHookRefCount = 0;

const wchar_t* NormalizeContext(const wchar_t* context) noexcept {
    return context ? context : L"(unknown)";
}

}  // namespace

bool AcquireMinHook(const wchar_t* context) noexcept {
    const wchar_t* normalized = NormalizeContext(context);
    std::lock_guard<std::mutex> guard(g_minHookMutex);

    if (g_minHookRefCount > 0) {
        ++g_minHookRefCount;
        LogMessage(LogLevel::Verbose,
                   L"MinHookSupport: Reusing initialized MinHook for %ls (refCount=%u)",
                   normalized, g_minHookRefCount);
        return true;
    }

    MH_STATUS status = MH_Initialize();
    if (status == MH_OK || status == MH_ERROR_ALREADY_INITIALIZED) {
        g_minHookRefCount = 1;
        if (status == MH_ERROR_ALREADY_INITIALIZED) {
            LogMessage(LogLevel::Warning,
                       L"MinHookSupport: MH_Initialize reported already initialized while acquiring for %ls (status=%d)",
                       normalized, status);
        } else {
            LogMessage(LogLevel::Verbose,
                       L"MinHookSupport: Initialized MinHook for %ls", normalized);
        }
        return true;
    }

    LogMessage(LogLevel::Error,
               L"MinHookSupport: Failed to initialize MinHook for %ls (status=%d)",
               normalized, status);
    return false;
}

void ReleaseMinHook(const wchar_t* context) noexcept {
    const wchar_t* normalized = NormalizeContext(context);
    std::lock_guard<std::mutex> guard(g_minHookMutex);

    if (g_minHookRefCount == 0) {
        LogMessage(LogLevel::Warning,
                   L"MinHookSupport: ReleaseMinHook called with zero ref count from %ls",
                   normalized);
        return;
    }

    --g_minHookRefCount;
    if (g_minHookRefCount > 0) {
        LogMessage(LogLevel::Verbose,
                   L"MinHookSupport: Retained MinHook for %ls (refCount=%u)",
                   normalized, g_minHookRefCount);
        return;
    }

    MH_STATUS status = MH_Uninitialize();
    if (status != MH_OK && status != MH_ERROR_NOT_INITIALIZED) {
        LogMessage(LogLevel::Warning,
                   L"MinHookSupport: MH_Uninitialize failed while releasing for %ls (status=%d)",
                   normalized, status);
    } else {
        LogMessage(LogLevel::Verbose,
                   L"MinHookSupport: Uninitialized MinHook after release for %ls",
                   normalized);
    }
}

MinHookScopedAcquire::MinHookScopedAcquire(const wchar_t* context) noexcept
    : m_context(NormalizeContext(context)), m_acquired(AcquireMinHook(m_context)) {}

MinHookScopedAcquire::~MinHookScopedAcquire() {
    Release();
}

void MinHookScopedAcquire::Release() noexcept {
    if (!m_acquired) {
        return;
    }

    ReleaseMinHook(m_context);
    m_acquired = false;
}

}  // namespace shelltabs

