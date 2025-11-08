#include "ComVTableHook.h"
#include "Logging.h"
#include "MinHookSupport.h"
#include <mutex>
#include <combaseapi.h>

#pragma comment(lib, "ole32.lib")

namespace shelltabs {

using shelltabs::LogLevel;
using shelltabs::LogMessage;

namespace {
    std::mutex g_hookMutex;
}

std::vector<ComVTableHook::VTableEntry> ComVTableHook::s_hookedMethods;
std::unordered_map<std::wstring, ComVTableHook::CreateCallback> ComVTableHook::s_classHooks;
void* ComVTableHook::s_originalCoCreateInstance = nullptr;
bool ComVTableHook::s_coCreateHooked = false;
bool ComVTableHook::s_minHookAcquired = false;

//=============================================================================
// VTable Manipulation
//=============================================================================

void** ComVTableHook::GetVTable(void* pInterface) {
    if (!pInterface) return nullptr;

    // The first pointer in a COM object points to the vtable
    return *reinterpret_cast<void***>(pInterface);
}

bool ComVTableHook::MakeVTableWritable(void** vtable, SIZE_T count, DWORD* pOldProtect) {
    if (!vtable || !pOldProtect) return false;

    SIZE_T size = count * sizeof(void*);
    return VirtualProtect(vtable, size, PAGE_READWRITE, pOldProtect) != FALSE;
}

bool ComVTableHook::RestoreVTableProtection(void** vtable, SIZE_T count, DWORD oldProtect) {
    if (!vtable) return false;

    SIZE_T size = count * sizeof(void*);
    DWORD dummy;
    return VirtualProtect(vtable, size, oldProtect, &dummy) != FALSE;
}

//=============================================================================
// Method Hooking
//=============================================================================

bool ComVTableHook::HookMethod(void* pInterface, UINT vtableIndex, void* pDetour, void** ppOriginal) {
    std::lock_guard<std::mutex> lock(g_hookMutex);

    if (!pInterface || !pDetour || !ppOriginal) {
        LogMessage(LogLevel::Error, L"ComVTableHook::HookMethod: Invalid parameters");
        return false;
    }

    void** vtable = GetVTable(pInterface);
    if (!vtable) {
        LogMessage(LogLevel::Error, L"ComVTableHook::HookMethod: Failed to get vtable");
        return false;
    }

    // Save original function pointer
    *ppOriginal = vtable[vtableIndex];

    // Check if already hooked
    for (const auto& entry : s_hookedMethods) {
        if (entry.pInterface == pInterface && entry.index == vtableIndex) {
            LogMessage(LogLevel::Warning, L"ComVTableHook::HookMethod: Method already hooked at index %u", vtableIndex);
            *ppOriginal = entry.pOriginal;
            return true;
        }
    }

    // Make vtable writable
    DWORD oldProtect;
    if (!MakeVTableWritable(&vtable[vtableIndex], 1, &oldProtect)) {
        LogMessage(LogLevel::Error, L"ComVTableHook::HookMethod: Failed to make vtable writable");
        return false;
    }

    // Replace function pointer
    vtable[vtableIndex] = pDetour;

    // Restore protection
    if (!RestoreVTableProtection(&vtable[vtableIndex], 1, oldProtect)) {
        LogMessage(LogLevel::Warning, L"ComVTableHook::HookMethod: Failed to restore vtable protection");
    }

    // Record the hook
    VTableEntry entry;
    entry.pInterface = pInterface;
    entry.index = vtableIndex;
    entry.pOriginal = *ppOriginal;
    entry.pDetour = pDetour;
    s_hookedMethods.push_back(entry);

    LogMessage(LogLevel::Info, L"ComVTableHook::HookMethod: Successfully hooked vtable[%u] for interface %p",
               vtableIndex, pInterface);

    return true;
}

bool ComVTableHook::UnhookMethod(void* pInterface, UINT vtableIndex) {
    std::lock_guard<std::mutex> lock(g_hookMutex);

    if (!pInterface) {
        LogMessage(LogLevel::Error, L"ComVTableHook::UnhookMethod: Invalid interface pointer");
        return false;
    }

    // Find the hook
    for (auto it = s_hookedMethods.begin(); it != s_hookedMethods.end(); ++it) {
        if (it->pInterface == pInterface && it->index == vtableIndex) {
            void** vtable = GetVTable(pInterface);
            if (!vtable) {
                LogMessage(LogLevel::Error, L"ComVTableHook::UnhookMethod: Failed to get vtable");
                return false;
            }

            // Make vtable writable
            DWORD oldProtect;
            if (!MakeVTableWritable(&vtable[vtableIndex], 1, &oldProtect)) {
                LogMessage(LogLevel::Error, L"ComVTableHook::UnhookMethod: Failed to make vtable writable");
                return false;
            }

            // Restore original function pointer
            vtable[vtableIndex] = it->pOriginal;

            // Restore protection
            RestoreVTableProtection(&vtable[vtableIndex], 1, oldProtect);

            // Remove from list
            s_hookedMethods.erase(it);

            LogMessage(LogLevel::Info, L"ComVTableHook::UnhookMethod: Successfully unhooked vtable[%u] for interface %p",
                       vtableIndex, pInterface);

            return true;
        }
    }

    LogMessage(LogLevel::Warning, L"ComVTableHook::UnhookMethod: Hook not found for interface %p, index %u",
               pInterface, vtableIndex);
    return false;
}

//=============================================================================
// CoCreateInstance Hooking
//=============================================================================

HRESULT WINAPI ComVTableHook::CoCreateInstance_Hook(
    REFCLSID rclsid,
    LPUNKNOWN pUnkOuter,
    DWORD dwClsContext,
    REFIID riid,
    LPVOID* ppv) {

    // Call original CoCreateInstance
    auto originalFunc = reinterpret_cast<decltype(&::CoCreateInstance)>(s_originalCoCreateInstance);
    HRESULT hr = originalFunc(rclsid, pUnkOuter, dwClsContext, riid, ppv);

    if (SUCCEEDED(hr) && ppv && *ppv) {
        std::lock_guard<std::mutex> lock(g_hookMutex);

        // Convert CLSID to string for lookup
        wchar_t clsidStr[40];
        StringFromGUID2(rclsid, clsidStr, ARRAYSIZE(clsidStr));

        // Check if we have a hook registered for this CLSID
        auto it = s_classHooks.find(clsidStr);
        if (it != s_classHooks.end()) {
            LogMessage(LogLevel::Info, L"ComVTableHook: Intercepted creation of %s", clsidStr);

            // Call the registered callback
            try {
                it->second(static_cast<IUnknown*>(*ppv), riid);
            }
            catch (...) {
                LogMessage(LogLevel::Error, L"ComVTableHook: Exception in create callback for %s", clsidStr);
            }
        }
    }

    return hr;
}

bool ComVTableHook::HookCoCreateInstance() {
    if (s_coCreateHooked) {
        LogMessage(LogLevel::Warning, L"ComVTableHook: CoCreateInstance already hooked");
        return true;
    }

    MinHookScopedAcquire minHookGuard(L"ComVTableHook::HookCoCreateInstance");
    if (!minHookGuard.IsAcquired()) {
        return false;
    }

    // Hook CoCreateInstance
    if (MH_CreateHook(&CoCreateInstance, &CoCreateInstance_Hook,
                      &s_originalCoCreateInstance) != MH_OK) {
        LogMessage(LogLevel::Error, L"ComVTableHook: Failed to create CoCreateInstance hook");
        return false;
    }

    if (MH_EnableHook(&CoCreateInstance) != MH_OK) {
        LogMessage(LogLevel::Error, L"ComVTableHook: Failed to enable CoCreateInstance hook");
        MH_RemoveHook(&CoCreateInstance);
        return false;
    }

    s_coCreateHooked = true;
    s_minHookAcquired = true;
    minHookGuard.Dismiss();
    LogMessage(LogLevel::Info, L"ComVTableHook: Successfully hooked CoCreateInstance");

    return true;
}

void ComVTableHook::UnhookCoCreateInstance() {
    if (!s_coCreateHooked) return;

    MH_DisableHook(&CoCreateInstance);
    MH_RemoveHook(&CoCreateInstance);

    s_coCreateHooked = false;
    if (s_minHookAcquired) {
        ReleaseMinHook(L"ComVTableHook::UnhookCoCreateInstance");
        s_minHookAcquired = false;
    }
    LogMessage(LogLevel::Info, L"ComVTableHook: Unhooked CoCreateInstance");
}

void ComVTableHook::RegisterClassHook(REFCLSID clsid, CreateCallback callback) {
    std::lock_guard<std::mutex> lock(g_hookMutex);

    wchar_t clsidStr[40];
    StringFromGUID2(clsid, clsidStr, ARRAYSIZE(clsidStr));

    s_classHooks[clsidStr] = callback;
    LogMessage(LogLevel::Info, L"ComVTableHook: Registered class hook for %s", clsidStr);
}

void ComVTableHook::UnregisterClassHook(REFCLSID clsid) {
    std::lock_guard<std::mutex> lock(g_hookMutex);

    wchar_t clsidStr[40];
    StringFromGUID2(clsid, clsidStr, ARRAYSIZE(clsidStr));

    auto it = s_classHooks.find(clsidStr);
    if (it != s_classHooks.end()) {
        s_classHooks.erase(it);
        LogMessage(LogLevel::Info, L"ComVTableHook: Unregistered class hook for %s", clsidStr);
    }
}

} // namespace shelltabs
