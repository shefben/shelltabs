#pragma once

#include <Windows.h>
#include <unknwn.h>
#include <UIRibbon.h>
#include <UIRibbonKeydef.h>
#include <wrl/client.h>
#include <vector>
#include <functional>
#include <unordered_map>

namespace shelltabs {

// Forward declarations
class RibbonCommandHandler;

//=============================================================================
// Custom Ribbon Tab Integration
//=============================================================================
// This module hooks into Windows Explorer's ribbon framework to inject
// a custom "Custom" tab with user-defined buttons. Since Microsoft doesn't
// provide official APIs for custom Explorer ribbon tabs, this uses MinHook
// to intercept the IUIFramework COM interface.
//
// Architecture:
// 1. Hook IUIFramework::LoadUI to intercept ribbon initialization
// 2. Inject our custom ribbon markup alongside Explorer's default ribbon
// 3. Provide IUICommandHandler implementation for custom buttons
// 4. Register custom tab with command IDs 50000-59999
//=============================================================================

// Custom command IDs for our ribbon tab (avoid conflicts with Explorer)
enum CustomRibbonCommands : UINT32 {
    // Tab and group commands
    cmdCustomTab = 50000,
    cmdCustomGroup1 = 50001,

    // Button commands (add more as needed)
    cmdCustomButton1 = 50100,
    cmdCustomButton2 = 50101,
    cmdCustomButton3 = 50102,
    cmdCustomButton4 = 50103,
    cmdCustomButton5 = 50104,
};

//=============================================================================
// Ribbon Command Handler
//=============================================================================
class RibbonCommandHandler : public IUICommandHandler {
public:
    RibbonCommandHandler();
    virtual ~RibbonCommandHandler();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // IUICommandHandler
    STDMETHODIMP Execute(
        UINT32 commandId,
        UI_EXECUTIONVERB verb,
        const PROPERTYKEY* key,
        const PROPVARIANT* currentValue,
        IUISimplePropertySet* commandExecutionProperties) override;

    STDMETHODIMP UpdateProperty(
        UINT32 commandId,
        REFPROPERTYKEY key,
        const PROPVARIANT* currentValue,
        PROPVARIANT* newValue) override;

    // Button action registration
    using ButtonCallback = std::function<void(HWND hwnd)>;
    void RegisterButtonCallback(UINT32 commandId, ButtonCallback callback);

private:
    LONG m_refCount;
    std::unordered_map<UINT32, ButtonCallback> m_buttonCallbacks;
};

//=============================================================================
// Ribbon Application Handler
//=============================================================================
class RibbonApplicationHandler : public IUIApplication {
public:
    RibbonApplicationHandler();
    virtual ~RibbonApplicationHandler();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // IUIApplication
    STDMETHODIMP OnViewChanged(
        UINT32 viewId,
        UI_VIEWTYPE typeId,
        IUnknown* view,
        UI_VIEWVERB verb,
        INT32 reasonCode) override;

    STDMETHODIMP OnCreateUICommand(
        UINT32 commandId,
        UI_COMMANDTYPE typeId,
        IUICommandHandler** commandHandler) override;

    STDMETHODIMP OnDestroyUICommand(
        UINT32 commandId,
        UI_COMMANDTYPE typeId,
        IUICommandHandler* commandHandler) override;

    // Set the shared command handler
    void SetCommandHandler(RibbonCommandHandler* handler);

private:
    LONG m_refCount;
    Microsoft::WRL::ComPtr<RibbonCommandHandler> m_commandHandler;
};

//=============================================================================
// Explorer Ribbon Hook Manager
//=============================================================================
class ExplorerRibbonHook {
public:
    // Initialize the ribbon hook system
    static bool Initialize();

    // Shutdown and remove hooks
    static void Shutdown();

    // Check if hooks are active
    static bool IsEnabled() { return s_enabled; }

    // Register a custom button action
    static void RegisterButtonAction(UINT32 commandId,
                                     RibbonCommandHandler::ButtonCallback callback);

    // Get the ribbon framework instance for a given Explorer window
    static IUIFramework* GetRibbonFramework(HWND explorerWindow);

private:
    // Hook implementations
    static HRESULT STDMETHODCALLTYPE IUIFramework_LoadUI_Hook(
        IUIFramework* pThis,
        HINSTANCE instance,
        LPCWSTR resourceName);

    static HRESULT STDMETHODCALLTYPE IUIFramework_Initialize_Hook(
        IUIFramework* pThis,
        HWND frameworkView,
        IUIApplication* application);

    // Helper to set up framework vtable hooks
    static void SetupFrameworkHooks(IUIFramework* pFramework);

    // Helper to inject custom ribbon markup
    static HRESULT InjectCustomRibbonTab(IUIFramework* framework, HWND hwnd);

    static bool s_enabled;
    static Microsoft::WRL::ComPtr<RibbonCommandHandler> s_commandHandler;
    static Microsoft::WRL::ComPtr<RibbonApplicationHandler> s_appHandler;
    static std::unordered_map<HWND, Microsoft::WRL::ComPtr<IUIFramework>> s_ribbonInstances;
};

} // namespace shelltabs
