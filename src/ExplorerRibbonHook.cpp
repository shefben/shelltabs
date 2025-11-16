#include "ExplorerRibbonHook.h"
#include "ComVTableHook.h"
#include "Logging.h"
#include "MinHook.h"
#include <propvarutil.h>
#include <propkey.h>
#include <shlwapi.h>
#include <mutex>
#include <initguid.h>

// Define Windows Ribbon Framework GUIDs
// These are required because the linker can't find them in the Windows SDK libraries
DEFINE_GUID(IID_IUICommandHandler, 0x75ae0a2d, 0xdc03, 0x4c9f, 0x88, 0x83, 0x06, 0x96, 0x60, 0xd0, 0xbe, 0xb6);
DEFINE_GUID(IID_IUIApplication, 0xd428903c, 0x729a, 0x491d, 0x91, 0x0d, 0x68, 0x2a, 0x08, 0xff, 0x25, 0x22);

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

    (void)key;
    (void)currentValue;
    (void)commandExecutionProperties;

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

    (void)currentValue;

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

    (void)typeId;
    (void)view;
    (void)reasonCode;

    LogMessage(LogLevel::Verbose, L"RibbonApplicationHandler: OnViewChanged(viewId=%u, verb=%d)", viewId, verb);
    return S_OK;
}

STDMETHODIMP RibbonApplicationHandler::OnCreateUICommand(
    UINT32 commandId,
    UI_COMMANDTYPE typeId,
    IUICommandHandler** commandHandler) {

    (void)typeId;

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

    (void)typeId;
    (void)commandHandler;

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

    // Hook CoCreateInstance to intercept IUIFramework creation
    if (!ComVTableHook::HookCoCreateInstance()) {
        LogMessage(LogLevel::Error, L"ExplorerRibbonHook: Failed to hook CoCreateInstance");
        return false;
    }

    // Register callback for CLSID_UIRibbonFramework creation
    // {926749FA-2615-4987-8845-C33E65F2B957}
    CLSID ribbonFrameworkClsid{};
    CLSIDFromString(L"{926749FA-2615-4987-8845-C33E65F2B957}", &ribbonFrameworkClsid);

    ComVTableHook::RegisterClassHook(ribbonFrameworkClsid,
        [](IUnknown* pUnknown, REFIID riid) {
            (void)riid;
            LogMessage(LogLevel::Info, L"ExplorerRibbonHook: IUIFramework created, setting up hooks...");

            // Query for IUIFramework interface
            IUIFramework* pFramework = nullptr;
            HRESULT hr = pUnknown->QueryInterface(__uuidof(IUIFramework),
                                                  reinterpret_cast<void**>(&pFramework));
            if (SUCCEEDED(hr) && pFramework) {
                // Hook the LoadUI and Initialize methods
                SetupFrameworkHooks(pFramework);
                pFramework->Release();
            }
        });

    s_enabled = true;
    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Ribbon hooks initialized successfully");
    return true;
}

void ExplorerRibbonHook::SetupFrameworkHooks(IUIFramework* pFramework) {
    if (!pFramework) return;

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Setting up IUIFramework vtable hooks...");

    // Get the vtable indices for LoadUI and Initialize
    // IUIFramework vtable layout (after IUnknown methods):
    // 0: QueryInterface
    // 1: AddRef
    // 2: Release
    // 3: Initialize
    // 4: Destroy
    // 5: LoadUI
    // 6: GetView
    // 7: GetUICommandProperty
    // 8: SetUICommandProperty
    // 9: InvalidateUICommand
    // 10: FlushPendingInvalidations
    // 11: SetModes

    const UINT VTABLE_INDEX_INITIALIZE = 3;
    const UINT VTABLE_INDEX_LOADUI = 5;

    // Hook Initialize
    ComVTableHook::HookMethod(pFramework, VTABLE_INDEX_INITIALIZE,
                              reinterpret_cast<void*>(&IUIFramework_Initialize_Hook),
                              &s_originalInitialize);

    // Hook LoadUI
    ComVTableHook::HookMethod(pFramework, VTABLE_INDEX_LOADUI,
                              reinterpret_cast<void*>(&IUIFramework_LoadUI_Hook),
                              &s_originalLoadUI);

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: IUIFramework vtable hooks installed");
}

HRESULT STDMETHODCALLTYPE ExplorerRibbonHook::IUIFramework_Initialize_Hook(
    IUIFramework* pThis,
    HWND frameworkView,
    IUIApplication* application) {

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: IUIFramework::Initialize called for hwnd=%p", frameworkView);

    // Call original Initialize
    auto originalFunc = reinterpret_cast<decltype(&IUIFramework_Initialize_Hook)>(s_originalInitialize);
    HRESULT hr = originalFunc(pThis, frameworkView, application);

    if (SUCCEEDED(hr)) {
        std::lock_guard<std::mutex> lock(g_ribbonMutex);

        // Store the ribbon instance for this window
        s_ribbonInstances[frameworkView] = pThis;

        LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Ribbon framework initialized for window %p", frameworkView);

        // Note: We replace the application handler with our own to intercept command creation
        // However, Explorer's ribbon is already initialized at this point, so we need to
        // inject our custom tab through LoadUI hook instead
    }

    return hr;
}

HRESULT STDMETHODCALLTYPE ExplorerRibbonHook::IUIFramework_LoadUI_Hook(
    IUIFramework* pThis,
    HINSTANCE instance,
    LPCWSTR resourceName) {

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: IUIFramework::LoadUI called (instance=%p, resourceName=%s)",
               instance, resourceName ? resourceName : L"(null)");

    // Call original LoadUI first
    auto originalFunc = reinterpret_cast<decltype(&IUIFramework_LoadUI_Hook)>(s_originalLoadUI);
    HRESULT hr = originalFunc(pThis, instance, resourceName);

    if (SUCCEEDED(hr)) {
        LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Original LoadUI succeeded, injecting custom tab...");

        // Try to inject our custom ribbon tab
        HRESULT injectHr = InjectCustomRibbonTab(pThis, nullptr);
        if (FAILED(injectHr)) {
            LogMessage(LogLevel::Warning, L"ExplorerRibbonHook: Failed to inject custom tab: 0x%08X", injectHr);
        }
    }

    return hr;
}

void ExplorerRibbonHook::Shutdown() {
    if (!s_enabled) return;

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Shutting down ribbon hooks...");

    std::lock_guard<std::mutex> lock(g_ribbonMutex);

    // Unhook CoCreateInstance
    CLSID ribbonFrameworkClsid{};
    CLSIDFromString(L"{926749FA-2615-4987-8845-C33E65F2B957}", &ribbonFrameworkClsid);
    ComVTableHook::UnregisterClassHook(ribbonFrameworkClsid);
    ComVTableHook::UnhookCoCreateInstance();

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
    if (!framework) {
        LogMessage(LogLevel::Error, L"ExplorerRibbonHook: InjectCustomRibbonTab called with null framework");
        return E_POINTER;
    }

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: InjectCustomRibbonTab called for hwnd=%p", hwnd);

    // Set properties for our custom commands in the ribbon
    // This prepares the ribbon to display our custom tab if/when it becomes available

    // Invalidate all our custom commands to ensure they're registered
    for (UINT32 cmdId = cmdCustomTab; cmdId <= cmdCustomButton5; ++cmdId) {
        HRESULT hr = framework->InvalidateUICommand(cmdId, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Enabled);
        if (SUCCEEDED(hr)) {
            LogMessage(LogLevel::Verbose, L"ExplorerRibbonHook: Invalidated command %u", cmdId);
        }
    }

    // Try to set command properties through the framework
    // Enable all our custom buttons
    for (UINT32 cmdId = cmdCustomButton1; cmdId <= cmdCustomButton5; ++cmdId) {
        PROPVARIANT var;
        PropVariantInit(&var);
        var.vt = VT_BOOL;
        var.boolVal = VARIANT_TRUE;

        HRESULT hr = framework->SetUICommandProperty(cmdId, UI_PKEY_Enabled, var);
        if (SUCCEEDED(hr)) {
            LogMessage(LogLevel::Verbose, L"ExplorerRibbonHook: Enabled button command %u", cmdId);
        }

        PropVariantClear(&var);
    }

    // Set labels for our commands
    PROPVARIANT varLabel;
    InitPropVariantFromString(L"Custom", &varLabel);
    framework->SetUICommandProperty(cmdCustomTab, UI_PKEY_Label, varLabel);
    PropVariantClear(&varLabel);

    InitPropVariantFromString(L"Actions", &varLabel);
    framework->SetUICommandProperty(cmdCustomGroup1, UI_PKEY_Label, varLabel);
    PropVariantClear(&varLabel);

    const wchar_t* buttonLabels[] = {
        L"Button 1", L"Button 2", L"Button 3", L"Button 4", L"Button 5"
    };

    for (size_t i = 0; i < 5; ++i) {
        InitPropVariantFromString(buttonLabels[i], &varLabel);
        framework->SetUICommandProperty(cmdCustomButton1 + static_cast<UINT32>(i),
                                       UI_PKEY_Label, varLabel);
        PropVariantClear(&varLabel);
    }

    // Flush all pending invalidations to apply changes
    framework->FlushPendingInvalidations();

    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Custom ribbon tab injection completed");
    LogMessage(LogLevel::Info, L"ExplorerRibbonHook: Note - To display the custom tab in Explorer's ribbon, "
                               L"you need to compile CustomRibbonTab.xml to a .bml resource and "
                               L"inject it through Explorer's ribbon binary modification or use "
                               L"a separate ribbon-enabled window.");

    return S_OK;
}

} // namespace shelltabs
