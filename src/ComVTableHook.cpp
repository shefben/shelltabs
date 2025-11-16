#include "ComVTableHook.h"
#include "Logging.h"
#include "MinHook.h"
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
bool ComVTableHook::s_initializedMinHook = false;

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

    // Check if this vtable slot is already hooked. Many COM objects share the same
    // vtable pointer, so patching it more than once would overwrite the stored
    // original function with our detour and cause infinite recursion.
    for (const auto& entry : s_hookedMethods) {
        if (entry.vtable == vtable && entry.index == vtableIndex) {
            LogMessage(LogLevel::Warning,
                       L"ComVTableHook::HookMethod: Method already hooked at index %u for vtable %p",
                       vtableIndex, vtable);
            *ppOriginal = entry.pOriginal;
            return true;
        }
    }

    // Save original function pointer
    *ppOriginal = vtable[vtableIndex];

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
    entry.vtable = vtable;
    entry.index = vtableIndex;
    entry.pOriginal = *ppOriginal;
    entry.pDetour = pDetour;
    s_hookedMethods.push_back(entry);

    LogMessage(LogLevel::Info, L"ComVTableHook::HookMethod: Successfully hooked vtable[%u] for vtable %p",
               vtableIndex, vtable);

    return true;
}

bool ComVTableHook::UnhookMethod(void* pInterface, UINT vtableIndex) {
    std::lock_guard<std::mutex> lock(g_hookMutex);

    if (!pInterface) {
        LogMessage(LogLevel::Error, L"ComVTableHook::UnhookMethod: Invalid interface pointer");
        return false;
    }

    void** vtable = GetVTable(pInterface);
    if (!vtable) {
        LogMessage(LogLevel::Error, L"ComVTableHook::UnhookMethod: Failed to get vtable");
        return false;
    }

    // Find the hook
    for (auto it = s_hookedMethods.begin(); it != s_hookedMethods.end(); ++it) {
        if (it->vtable == vtable && it->index == vtableIndex) {
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

            LogMessage(LogLevel::Info, L"ComVTableHook::UnhookMethod: Successfully unhooked vtable[%u] for vtable %p",
                       vtableIndex, vtable);

            return true;
        }
    }

    LogMessage(LogLevel::Warning, L"ComVTableHook::UnhookMethod: Hook not found for vtable %p, index %u",
               vtable, vtableIndex);
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
        // Convert CLSID to string for lookup
        wchar_t clsidStr[40];
        StringFromGUID2(rclsid, clsidStr, ARRAYSIZE(clsidStr));
        std::wstring clsidKey(clsidStr);

        CreateCallback callback;
        {
            std::lock_guard<std::mutex> lock(g_hookMutex);
            auto it = s_classHooks.find(clsidKey);
            if (it != s_classHooks.end()) {
                callback = it->second;
            }
        }

        if (callback) {
            LogMessage(LogLevel::Info, L"ComVTableHook: Intercepted creation of %s", clsidKey.c_str());

            // Call the registered callback outside the global hook lock so callbacks
            // can safely invoke HookMethod/UnhookMethod which also lock g_hookMutex.
            try {
                callback(static_cast<IUnknown*>(*ppv), riid);
            }
            catch (...) {
                LogMessage(LogLevel::Error, L"ComVTableHook: Exception in create callback for %s", clsidKey.c_str());
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

    bool initializedHere = false;
    if (!s_initializedMinHook) {
        MH_STATUS initStatus = MH_Initialize();
        if (initStatus == MH_OK) {
            initializedHere = true;
            s_initializedMinHook = true;
            LogMessage(LogLevel::Info, L"ComVTableHook: MinHook initialized for CoCreateInstance hook");
        }
        else if (initStatus == MH_ERROR_ALREADY_INITIALIZED) {
            LogMessage(LogLevel::Info, L"ComVTableHook: MinHook already initialized by another subsystem; reusing state");
        }
        else {
            LogMessage(LogLevel::Error, L"ComVTableHook: Failed to initialize MinHook (status %d)", initStatus);
            return false;
        }
    }

    const auto cleanupOnFailure = [initializedHere]() {
        MH_DisableHook(&CoCreateInstance);
        MH_RemoveHook(&CoCreateInstance);
        if (initializedHere) {
            MH_STATUS uninitStatus = MH_Uninitialize();
            if (uninitStatus == MH_OK) {
                ComVTableHook::s_initializedMinHook = false;
                LogMessage(LogLevel::Info, L"ComVTableHook: MinHook uninitialized after failed hook setup");
            }
            else {
                LogMessage(LogLevel::Warning, L"ComVTableHook: Failed to uninitialize MinHook after hook setup failure (status %d)", uninitStatus);
            }
        }
    };

    // Hook CoCreateInstance
    MH_STATUS createStatus = MH_CreateHook(&CoCreateInstance, &CoCreateInstance_Hook,
                                           &s_originalCoCreateInstance);
    if (createStatus != MH_OK) {
        LogMessage(LogLevel::Error, L"ComVTableHook: Failed to create CoCreateInstance hook (status %d)", createStatus);
        cleanupOnFailure();
        return false;
    }

    MH_STATUS enableStatus = MH_EnableHook(&CoCreateInstance);
    if (enableStatus != MH_OK) {
        LogMessage(LogLevel::Error, L"ComVTableHook: Failed to enable CoCreateInstance hook (status %d)", enableStatus);
        cleanupOnFailure();
        return false;
    }

    s_coCreateHooked = true;
    LogMessage(LogLevel::Info, L"ComVTableHook: Successfully hooked CoCreateInstance");

    return true;
}

void ComVTableHook::UnhookCoCreateInstance() {
    if (!s_coCreateHooked) return;

    MH_STATUS disableStatus = MH_DisableHook(&CoCreateInstance);
    if (disableStatus != MH_OK) {
        LogMessage(LogLevel::Warning, L"ComVTableHook: Failed to disable CoCreateInstance hook (status %d)", disableStatus);
    }

    MH_STATUS removeStatus = MH_RemoveHook(&CoCreateInstance);
    if (removeStatus != MH_OK) {
        LogMessage(LogLevel::Warning, L"ComVTableHook: Failed to remove CoCreateInstance hook (status %d)", removeStatus);
    }

    s_coCreateHooked = false;
    LogMessage(LogLevel::Info, L"ComVTableHook: Unhooked CoCreateInstance");

    if (s_initializedMinHook) {
        MH_STATUS uninitStatus = MH_Uninitialize();
        if (uninitStatus == MH_OK) {
            s_initializedMinHook = false;
            LogMessage(LogLevel::Info, L"ComVTableHook: MinHook uninitialized after CoCreateInstance unhook");
        }
        else {
            LogMessage(LogLevel::Warning, L"ComVTableHook: Failed to uninitialize MinHook during teardown (status %d)", uninitStatus);
        }
    }
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
