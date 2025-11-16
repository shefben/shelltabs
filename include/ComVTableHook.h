#pragma once

#include <Windows.h>
#include <unknwn.h>
#include <vector>
#include <unordered_map>
#include <functional>
#include <string>

namespace shelltabs {

//=============================================================================
// COM VTable Hooking Utility
//=============================================================================
// This utility provides safe hooking of COM interface vtable methods.
// It works by:
// 1. Intercepting CoCreateInstance to catch COM object creation
// 2. Replacing vtable function pointers with our hook functions
// 3. Maintaining original function pointers for call-through
//
// WARNING: This is an advanced, undocumented technique that:
// - May break with Windows updates
// - Requires careful memory management
// - Can cause crashes if not implemented correctly
// - Should be used only when no official API exists
//=============================================================================

class ComVTableHook {
public:
    // Hook a specific method in a COM interface vtable
    // vtableIndex: The index of the method in the vtable (0-based, after IUnknown methods)
    // pInterface: Pointer to the COM interface instance
    // pDetour: Pointer to your hook function
    // ppOriginal: Receives pointer to the original function
    static bool HookMethod(void* pInterface, UINT vtableIndex, void* pDetour, void** ppOriginal);

    // Unhook a previously hooked method
    static bool UnhookMethod(void* pInterface, UINT vtableIndex);

    // Hook CoCreateInstance to intercept specific COM object creation
    static bool HookCoCreateInstance();

    // Unhook CoCreateInstance
    static void UnhookCoCreateInstance();

    // Register a COM class to intercept during creation
    // clsid: The CLSID to intercept
    // callback: Function called when object is created, receives IUnknown*
    using CreateCallback = std::function<void(IUnknown* pObject, REFIID riid)>;
    static void RegisterClassHook(REFCLSID clsid, CreateCallback callback);

    // Unregister a COM class hook
    static void UnregisterClassHook(REFCLSID clsid);

private:
    struct VTableEntry {
        void** vtable;   // Address of the COM vtable that owns the slot we patched
        UINT index;      // Slot index within the vtable
        void* pOriginal; // Original function pointer we replaced
        void* pDetour;   // Our detour that currently lives in the slot
    };

    // CoCreateInstance hook
    static HRESULT WINAPI CoCreateInstance_Hook(
        REFCLSID rclsid,
        LPUNKNOWN pUnkOuter,
        DWORD dwClsContext,
        REFIID riid,
        LPVOID* ppv);

    // Get vtable pointer from interface
    static void** GetVTable(void* pInterface);

    // Change memory protection to allow vtable modification
    static bool MakeVTableWritable(void** vtable, SIZE_T count, DWORD* pOldProtect);

    // Restore memory protection
    static bool RestoreVTableProtection(void** vtable, SIZE_T count, DWORD oldProtect);

    static std::vector<VTableEntry> s_hookedMethods;
    static std::unordered_map<std::wstring, CreateCallback> s_classHooks;
    static void* s_originalCoCreateInstance;
    static bool s_coCreateHooked;
    static bool s_initializedMinHook;
};

//=============================================================================
// Helper Macros for COM Vtable Hooking
//=============================================================================

// Calculate vtable index for a method
// IUnknown has 3 methods (QueryInterface, AddRef, Release)
// So the first custom method is at index 3

#define VTABLE_INDEX_IUNKNOWN_QUERYINTERFACE 0
#define VTABLE_INDEX_IUNKNOWN_ADDREF 1
#define VTABLE_INDEX_IUNKNOWN_RELEASE 2

// Example for IUIFramework (inherits from IUnknown):
// Initialize = 3, Destroy = 4, LoadUI = 5, GetView = 6, etc.
#define VTABLE_INDEX_IUIFRAMEWORK_INITIALIZE 3
#define VTABLE_INDEX_IUIFRAMEWORK_DESTROY 4
#define VTABLE_INDEX_IUIFRAMEWORK_LOADUI 5
#define VTABLE_INDEX_IUIFRAMEWORK_GETVIEW 6
#define VTABLE_INDEX_IUIFRAMEWORK_GETUICOMMANDPROPERTY 7
#define VTABLE_INDEX_IUIFRAMEWORK_SETUICOMMANDPROPERTY 8
#define VTABLE_INDEX_IUIFRAMEWORK_INVALIDATEUICOMMAND 9
#define VTABLE_INDEX_IUIFRAMEWORK_FLUSHPENDINGINVALIDATIONS 10
#define VTABLE_INDEX_IUIFRAMEWORK_SETMODES 11

} // namespace shelltabs
