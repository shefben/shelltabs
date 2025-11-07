#include "ExplorerRibbonHook.h"
#include "Logging.h"
#include "MinHook.h"
#include <propvarutil.h>
#include <propkey.h>
#include <shlwapi.h>
#include <mutex>

#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "shlwapi.lib")

namespace shelltabs {

using shelltabs::LogLevel;
using shelltabs::LogMessage;

namespace {
    std::mutex g_ribbonMutex;
}

//=============================================================================
// RibbonCommandHandler Implementation
//=============================================================================

RibbonCommandHandler::RibbonCommandHandler()
    : m_refCount(1) {
}

RibbonCommandHandler::~RibbonCommandHandler() {
}

STDMETHODIMP RibbonCommandHandler::QueryInterface(REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;

    if (riid == IID_IUnknown || riid == IID_IUICommandHandler) {
        *ppv = static_cast<IUICommandHandler*>(this);
        AddRef();
        return S_OK;
    }

    *ppv = nullptr;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) RibbonCommandHandler::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

STDMETHODIMP_(ULONG) RibbonCommandHandler::Release() {
    LONG refCount = InterlockedDecrement(&m_refCount);
    if (refCount == 0) {
        delete this;
    }
    return refCount;
}

STDMETHODIMP RibbonCommandHandler::Execute(
    UINT32 commandId,
    UI_EXECUTIONVERB verb,
    const PROPERTYKEY* key,
    const PROPVARIANT* currentValue,
    IUISimplePropertySet* commandExecutionProperties) {

    if (verb == UI_EXECUTIONVERB_EXECUTE) {
        LogMessage(LogLevel::Info, L"RibbonCommandHandler: Executing command %u", commandId);

        // Find and execute registered callback
        auto it = m_buttonCallbacks.find(commandId);
        if (it != m_buttonCallbacks.end()) {
            // Get the current Explorer window
            HWND hwnd = GetForegroundWindow();
            it->second(hwnd);
            return S_OK;
        } else {
            LogMessage(LogLevel::Warning, L"RibbonCommandHandler: No callback registered for command %u", commandId);
        }
    }

    return E_NOTIMPL;
}

STDMETHODIMP RibbonCommandHandler::UpdateProperty(
    UINT32 commandId,
    REFPROPERTYKEY key,
    const PROPVARIANT* currentValue,
    PROPVARIANT* newValue) {

    // Provide property values for our custom commands
    if (key == UI_PKEY_Label) {
        switch (commandId) {
        case cmdCustomTab:
            return InitPropVariantFromString(L"Custom", newValue);

        case cmdCustomGroup1:
            return InitPropVariantFromString(L"Actions", newValue);

        case cmdCustomButton1:
            return InitPropVariantFromString(L"Button 1", newValue);

        case cmdCustomButton2:
            return InitPropVariantFromString(L"Button 2", newValue);

        case cmdCustomButton3:
            return InitPropVariantFromString(L"Button 3", newValue);

        case cmdCustomButton4:
            return InitPropVariantFromString(L"Button 4", newValue);

        case cmdCustomButton5:
            return InitPropVariantFromString(L"Button 5", newValue);

        default:
            break;
        }
    }
    else if (key == UI_PKEY_TooltipTitle) {
        switch (commandId) {
        case cmdCustomButton1:
            return InitPropVariantFromString(L"Custom Button 1", newValue);

        case cmdCustomButton2:
            return InitPropVariantFromString(L"Custom Button 2", newValue);

        case cmdCustomButton3:
            return InitPropVariantFromString(L"Custom Button 3", newValue);

        case cmdCustomButton4:
            return InitPropVariantFromString(L"Custom Button 4", newValue);

        case cmdCustomButton5:
            return InitPropVariantFromString(L"Custom Button 5", newValue);

        default:
            break;
        }
    }
    else if (key == UI_PKEY_TooltipDescription) {
        switch (commandId) {
        case cmdCustomButton1:
            return InitPropVariantFromString(L"Execute custom action 1", newValue);

        case cmdCustomButton2:
            return InitPropVariantFromString(L"Execute custom action 2", newValue);

        case cmdCustomButton3:
            return InitPropVariantFromString(L"Execute custom action 3", newValue);

        case cmdCustomButton4:
            return InitPropVariantFromString(L"Execute custom action 4", newValue);

        case cmdCustomButton5:
            return InitPropVariantFromString(L"Execute custom action 5", newValue);

        default:
            break;
        }
    }
    else if (key == UI_PKEY_Enabled) {
        // All our buttons are always enabled
        if (commandId >= cmdCustomButton1 && commandId <= cmdCustomButton5) {
            return InitPropVariantFromBoolean(TRUE, newValue);
        }
    }

    return E_NOTIMPL;
}

void RibbonCommandHandler::RegisterButtonCallback(UINT32 commandId, ButtonCallback callback) {
    m_buttonCallbacks[commandId] = callback;
    LogMessage(LogLevel::Info, L"RibbonCommandHandler: Registered callback for command %u", commandId);
}

//=============================================================================
// RibbonApplicationHandler Implementation
//=============================================================================

RibbonApplicationHandler::RibbonApplicationHandler()
    : m_refCount(1) {
}

RibbonApplicationHandler::~RibbonApplicationHandler() {
}

STDMETHODIMP RibbonApplicationHandler::QueryInterface(REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;

    if (riid == IID_IUnknown || riid == IID_IUIApplication) {
        *ppv = static_cast<IUIApplication*>(this);
        AddRef();
        return S_OK;
    }

    *ppv = nullptr;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) RibbonApplicationHandler::AddRef() {
    return InterlockedIncrement(&m_refCount);
}

STDMETHODIMP_(ULONG) RibbonApplicationHandler::Release() {
    LONG refCount = InterlockedDecrement(&m_refCount);
    if (refCount == 0) {
        delete this;
    }
    return refCount;
}

STDMETHODIMP RibbonApplicationHandler::OnViewChanged(
    UINT32 viewId,
    UI_VIEWTYPE typeId,
    IUnknown* view,
    UI_VIEWVERB verb,
    INT32 reasonCode) {

    LogMessage(LogLevel::Verbose, L"RibbonApplicationHandler: OnViewChanged(viewId=%u, verb=%d)", viewId, verb);
    return S_OK;
}

STDMETHODIMP RibbonApplicationHandler::OnCreateUICommand(
    UINT32 commandId,
    UI_COMMANDTYPE typeId,
    IUICommandHandler** commandHandler) {

    // Return our command handler for custom commands
    if (commandId >= cmdCustomTab && commandId < 60000) {
        if (m_commandHandler) {
            *commandHandler = m_commandHandler.Get();
            (*commandHandler)->AddRef();
            LogMessage(LogLevel::Info, L"RibbonApplicationHandler: Created command handler for %u", commandId);
            return S_OK;
        }
    }

    return E_NOTIMPL;
}

STDMETHODIMP RibbonApplicationHandler::OnDestroyUICommand(
    UINT32 commandId,
    UI_COMMANDTYPE typeId,
    IUICommandHandler* commandHandler) {

    LogMessage(LogLevel::Verbose, L"RibbonApplicationHandler: OnDestroyUICommand(%u)", commandId);
    return S_OK;
}

void RibbonApplicationHandler::SetCommandHandler(RibbonCommandHandler* handler) {
    m_commandHandler = handler;
}

//=============================================================================
// ExplorerRibbonHook Implementation
//=============================================================================

bool ExplorerRibbonHook::s_enabled = false;
void* ExplorerRibbonHook::s_originalLoadUI = nullptr;
void* ExplorerRibbonHook::s_originalInitialize = nullptr;
Microsoft::WRL::ComPtr<RibbonCommandHandler> ExplorerRibbonHook::s_commandHandler;
Microsoft::WRL::ComPtr<RibbonApplicationHandler> ExplorerRibbonHook::s_appHandler;
std::unordered_map<HWND, Microsoft::WRL::ComPtr<IUIFramework>> ExplorerRibbonHook::s_ribbonInstances;

bool ExplorerRibbonHook::Initialize() {
    if (s_enabled) return true;

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Initializing ribbon hooks...");

    // Create shared command and application handlers
    s_commandHandler = new RibbonCommandHandler();
    s_appHandler = new RibbonApplicationHandler();
    s_appHandler->SetCommandHandler(s_commandHandler.Get());

    // Register default button callbacks (users can override these)
    RegisterButtonAction(cmdCustomButton1, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Custom Button 1 clicked!");
        MessageBoxW(hwnd, L"This is Custom Button 1.\n\nYou can register your own callback for this button.",
                    L"Custom Ribbon Button", MB_OK | MB_ICONINFORMATION);
    });

    RegisterButtonAction(cmdCustomButton2, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Custom Button 2 clicked!");
        MessageBoxW(hwnd, L"This is Custom Button 2.\n\nCustomize this action as needed.",
                    L"Custom Ribbon Button", MB_OK | MB_ICONINFORMATION);
    });

    RegisterButtonAction(cmdCustomButton3, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Custom Button 3 clicked!");
        MessageBoxW(hwnd, L"This is Custom Button 3.",
                    L"Custom Ribbon Button", MB_OK | MB_ICONINFORMATION);
    });

    RegisterButtonAction(cmdCustomButton4, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Custom Button 4 clicked!");
        MessageBoxW(hwnd, L"This is Custom Button 4.",
                    L"Custom Ribbon Button", MB_OK | MB_ICONINFORMATION);
    });

    RegisterButtonAction(cmdCustomButton5, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Custom Button 5 clicked!");
        MessageBoxW(hwnd, L"This is Custom Button 5.",
                    L"Custom Ribbon Button", MB_OK | MB_ICONINFORMATION);
    });

    // Note: Hooking IUIFramework COM methods requires a different approach than
    // standard Win32 API hooking. We would need to:
    // 1. Hook CoCreateInstance to intercept CLSID_UIRibbonFramework creation
    // 2. Replace the vtable entries for LoadUI and Initialize
    //
    // This is an advanced technique and may be fragile across Windows updates.
    // For a production implementation, consider:
    // - Using Detours instead of MinHook for COM vtable hooking
    // - Implementing a COM wrapper proxy that sits between Explorer and the framework
    // - Using DLL injection to load before Explorer initializes its ribbon

    LogMessage(LogLevel::Warning, L"ExplorerRibbonHook: COM interface hooking not yet implemented.");
    LogMessage(LogLevel::Warning, L"ExplorerRibbonHook: This requires vtable hooking or CoCreateInstance interception.");

    s_enabled = true;
    return true;
}

void ExplorerRibbonHook::Shutdown() {
    if (!s_enabled) return;

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Shutting down ribbon hooks...");

    std::lock_guard<std::mutex> lock(g_ribbonMutex);

    // Release all ribbon instances
    s_ribbonInstances.clear();

    // Release handlers
    s_commandHandler.Reset();
    s_appHandler.Reset();

    s_enabled = false;
}

void ExplorerRibbonHook::RegisterButtonAction(
    UINT32 commandId,
    RibbonCommandHandler::ButtonCallback callback) {

    if (s_commandHandler) {
        s_commandHandler->RegisterButtonCallback(commandId, callback);
    }
}

IUIFramework* ExplorerRibbonHook::GetRibbonFramework(HWND explorerWindow) {
    std::lock_guard<std::mutex> lock(g_ribbonMutex);

    auto it = s_ribbonInstances.find(explorerWindow);
    if (it != s_ribbonInstances.end()) {
        return it->second.Get();
    }

    return nullptr;
}

HRESULT ExplorerRibbonHook::InjectCustomRibbonTab(IUIFramework* framework, HWND hwnd) {
    // This function would programmatically add ribbon elements
    // In a real implementation, you would:
    // 1. Get the IUIRibbon interface from the framework
    // 2. Use InvalidateUICommand to refresh the ribbon
    // 3. Return updated tab collections when queried

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: InjectCustomRibbonTab called for hwnd=%p", hwnd);

    // TODO: Implement actual ribbon tab injection
    // This requires either:
    // - Modifying the ribbon markup binary before it's loaded
    // - Implementing IUICollection to provide additional tabs dynamically
    // - Using undocumented ribbon internal APIs

    return S_OK;
}

} // namespace shelltabs
