# Windows Explorer Ribbon Customization

This document explains how to add a custom tab to the Windows File Explorer ribbon, the challenges involved, and the implementation approach used in ShellTabs.

## Table of Contents

1. [Overview](#overview)
2. [Official API Limitations](#official-api-limitations)
3. [Undocumented Approaches](#undocumented-approaches)
4. [Implementation Strategy](#implementation-strategy)
5. [Usage Guide](#usage-guide)
6. [Troubleshooting](#troubleshooting)
7. [Windows 11 Compatibility](#windows-11-compatibility)

---

## Overview

Windows Explorer's ribbon interface (introduced in Windows 8) provides tabs like "Home", "Share", "View", and contextual tabs that appear based on the selected file type. While Microsoft provides APIs for adding items to **context menus** and **specific ribbon buttons** via file type associations, **adding entirely new ribbon tabs is not officially supported**.

### What We Want to Achieve

```
┌────────────────────────────────────────────────────────────┐
│ File  Home  Share  View  [Custom] ◄── Our custom tab       │
├────────────────────────────────────────────────────────────┤
│ [Button 1] [Button 2] [Button 3] [Button 4] [Button 5]     │
└────────────────────────────────────────────────────────────┘
```

---

## Official API Limitations

### What IS Officially Supported

✅ **Context Menu Extensions** (IExplorerCommand)
- Add items to right-click menus
- Works on Windows 7-11
- **ShellTabs already uses this** - see `src/OpenFolderCommand.cpp`

✅ **File Type Association Ribbon Buttons**
- Customize specific buttons for file types (e.g., "Extract All" for ZIP files)
- Register via `HKEY_CLASSES_ROOT\Explorer.AssocActionId.*`
- Limited to predefined action IDs

✅ **Deskband Toolbars** (IDeskBand - **Deprecated**)
- Add custom toolbars below the ribbon
- **ShellTabs uses this approach** on Windows 10
- **Removed in Windows 11** - requires ExplorerPatcher

### What Is NOT Supported

❌ **Custom Ribbon Tabs**
- No public API to add new tabs
- No official way to modify Explorer's ribbon structure
- Microsoft explicitly states: "Use context menus instead"

❌ **Ribbon Tab Reordering**
- Cannot change tab order
- Cannot hide existing tabs

❌ **Dynamic Ribbon Content**
- Limited programmatic access to ribbon state
- No official `IUICollection` extension points for Explorer

---

## Undocumented Approaches

Since official APIs don't exist, we must use **advanced hooking techniques**:

### Approach 1: COM Vtable Hooking ⭐ (Our Implementation)

**How it works:**
1. Hook `CoCreateInstance` to intercept ribbon framework creation
2. When Explorer creates `CLSID_UIRibbonFramework`, intercept it
3. Hook the `IUIFramework::LoadUI` method vtable
4. Modify or extend the ribbon markup before it loads

**Pros:**
- Most flexible approach
- Can inject custom tabs dynamically
- Allows full ribbon customization

**Cons:**
- Very advanced technique
- May break with Windows updates
- Requires careful memory management
- Can crash Explorer if done incorrectly

**Implementation files:**
- `include/ComVTableHook.h` - COM vtable hooking utility
- `include/ExplorerRibbonHook.h` - Ribbon-specific hooks
- `src/ComVTableHook.cpp` - Implementation
- `src/ExplorerRibbonHook.cpp` - Ribbon hook logic

### Approach 2: Ribbon Binary Patching

**How it works:**
1. Extract Explorer's ribbon markup binary from resources
2. Parse the binary ribbon format
3. Inject custom tab XML into the structure
4. Reload the modified ribbon

**Pros:**
- No runtime hooking required

**Cons:**
- Extremely fragile (binary format is undocumented)
- Different across Windows versions
- Very difficult to maintain

### Approach 3: DLL Injection + Markup Interception

**How it works:**
1. Inject DLL before Explorer initializes ribbon
2. Hook file I/O to intercept ribbon markup loading
3. Replace/augment the ribbon XML before it's parsed

**Pros:**
- Works before ribbon initialization
- Cleaner than vtable hooking

**Cons:**
- Requires early injection (before Explorer loads ribbon)
- Still depends on undocumented behavior

---

## Implementation Strategy

Our implementation uses **Approach 1: COM Vtable Hooking** integrated with MinHook (already used extensively in this project).

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        Explorer.exe                         │
└────────────────┬────────────────────────────────────────────┘
                 │
                 │ Creates CLSID_UIRibbonFramework
                 ▼
        ┌──────────────────┐
        │ CoCreateInstance │◄──── Hooked by ComVTableHook
        └────────┬─────────┘
                 │
                 ▼
        ┌────────────────┐
        │  IUIFramework  │
        └────────┬───────┘
                 │
                 │ Vtable methods hooked:
                 │  - LoadUI()  ◄──── Inject custom ribbon XML
                 │  - Initialize()
                 │
                 ▼
      ┌──────────────────────┐
      │ ExplorerRibbonHook   │
      │                      │
      │ - Injects custom tab │
      │ - Provides command   │
      │   handlers           │
      └──────────┬───────────┘
                 │
                 ▼
       ┌────────────────────┐
       │ IUICommandHandler  │  ◄──── Handles button clicks
       │                    │
       │ - Button 1 action  │
       │ - Button 2 action  │
       │ - etc.             │
       └────────────────────┘
```

### Key Components

#### 1. COM Vtable Hooking (`ComVTableHook`)

Provides utilities to hook COM interface methods:

```cpp
// Hook a vtable method
ComVTableHook::HookMethod(pInterface, VTABLE_INDEX_IUIFRAMEWORK_LOADUI,
                          &MyLoadUIHook, &originalLoadUI);

// Intercept COM object creation
ComVTableHook::HookCoCreateInstance();
ComVTableHook::RegisterClassHook(CLSID_UIRibbonFramework, [](IUnknown* pObj) {
    // Called when Explorer creates the ribbon framework
});
```

#### 2. Ribbon Hook Manager (`ExplorerRibbonHook`)

High-level ribbon customization:

```cpp
// Initialize ribbon hooks
ExplorerRibbonHook::Initialize();

// Register button actions
ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton1, [](HWND hwnd) {
    MessageBox(hwnd, L"Button 1 clicked!", L"Custom Action", MB_OK);
});
```

#### 3. Command Handlers (`RibbonCommandHandler`, `RibbonApplicationHandler`)

Implement `IUICommandHandler` and `IUIApplication` to:
- Provide button labels, tooltips, icons
- Handle button click events
- Update ribbon state

#### 4. Ribbon XML Markup (`CustomRibbonTab.xml`)

Defines the custom tab structure:
- Tab definition
- Button groups
- Command IDs
- Images and tooltips

---

## Usage Guide

### Basic Setup

1. **Initialize the ribbon hooks** (in `DllMain` or BHO initialization):

```cpp
#include "ExplorerRibbonHook.h"

// In your DLL initialization
shelltabs::ExplorerRibbonHook::Initialize();
```

2. **Register button callbacks**:

```cpp
using shelltabs::ExplorerRibbonHook;
using shelltabs::cmdCustomButton1;

ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton1, [](HWND hwnd) {
    // Your custom action here
    MessageBox(hwnd, L"Hello from custom ribbon!", L"ShellTabs", MB_OK);
});
```

3. **Cleanup on shutdown**:

```cpp
shelltabs::ExplorerRibbonHook::Shutdown();
```

### Advanced: Custom Button Actions

```cpp
// Example: Open folder in new tab
ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton2, [](HWND hwnd) {
    // Get current folder path
    wchar_t path[MAX_PATH];
    if (SHGetFolderPath(hwnd, CSIDL_DESKTOP, NULL, 0, path) == S_OK) {
        // Open in new tab (integrate with your tab manager)
        TabManager::OpenFolderInNewTab(path);
    }
});

// Example: Toggle hidden files
ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton3, [](HWND hwnd) {
    // Send command to Explorer window
    SendMessage(hwnd, WM_COMMAND, FCIDM_SHVIEW_SHOWHIDDEN, 0);
});
```

### Compiling Ribbon XML

If you modify `CustomRibbonTab.xml`:

1. Install Windows SDK (for `uicc.exe`)

2. Compile the ribbon markup:
```batch
cd src
uicc.exe CustomRibbonTab.xml CustomRibbonTab.bml ^
  /header:CustomRibbonTab.h ^
  /res:CustomRibbonTab.rc
```

> **Reminder:** `CustomRibbonTab.xml` expects the small/large PNG glyphs to be
> present in `src/ribbon_images`. Copy your assets into that directory before
> invoking `uicc.exe` so the compiler can resolve each `<Image
> Source="ribbon_images/...">` reference.

3. Add to your resource script:
```rc
#include "CustomRibbonTab.h"

APPLICATION_RIBBON UIFILE "CustomRibbonTab.bml"
```

4. Include the generated header in your code:
```cpp
#include "CustomRibbonTab.h"
// Command IDs like cmdCustomTab are now available
```

---

## Troubleshooting

### Ribbon Tab Doesn't Appear

**Possible causes:**

1. **Hook failed to install**
   - Check logs for error messages
   - Ensure MinHook initialized successfully
   - Verify `ExplorerRibbonHook::Initialize()` returns `true`

2. **Explorer loaded ribbon before hooks installed**
   - Solution: Hook `CoCreateInstance` earlier in DLL load
   - Consider using DLL injection before Explorer starts

3. **Windows version incompatibility**
   - Ribbon vtable layout may differ across Windows versions
   - Test on target Windows version

### Buttons Don't Respond to Clicks

**Possible causes:**

1. **Command handler not registered**
   ```cpp
   // Verify this was called:
   ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton1, callback);
   ```

2. **Command ID mismatch**
   - Ensure XML command IDs match enum values
   - Check `RibbonCommandHandler::Execute()` receives correct command ID

3. **IUICommandHandler not provided**
   - Verify `OnCreateUICommand()` returns your handler for custom command IDs

### Explorer Crashes

**Likely causes:**

1. **Vtable corruption**
   - Double-check vtable index calculations
   - Ensure memory protection is correctly set/restored

2. **Calling convention mismatch**
   - All COM methods use `STDMETHODCALLTYPE` (__stdcall)
   - Verify hook function signatures match exactly

3. **Incorrect original function call**
   - Always call through original function pointer
   - Don't return without calling original

**Debugging tips:**
- Enable logging to trace execution flow
- Use WinDbg to attach to Explorer and catch crashes
- Test in a VM first to avoid system instability

### Ribbon Disappears After Windows Update

**This is expected** - undocumented hooking may break with updates.

**Solutions:**
- Monitor Windows updates and test promptly
- Consider fallback to context menus if hooks fail
- Implement version detection and adaptive hooking

---

## Windows 11 Compatibility

### Command Bar vs Ribbon

Windows 11 replaced the ribbon with a simplified "command bar":

```
Windows 10:  [File] [Home] [Share] [View]  ← Full ribbon
Windows 11:  [⊙ New] [↻ Sort] [⋯ View] [⋮]  ← Command bar
```

### Impact on Custom Tabs

❌ **Even harder to customize** on Windows 11:
- No ribbon framework (`IUIFramework` not used)
- Custom XAML-based command bar
- Microsoft hasn't exposed APIs yet

### Workarounds for Windows 11

1. **Use ExplorerPatcher**
   - Restores Windows 10 ribbon
   - Our ribbon hooks should work

2. **Context Menu Extensions**
   - Still supported and recommended
   - Use `IExplorerCommand` (already implemented in ShellTabs)

3. **Deskband with ExplorerPatcher**
   - The existing ShellTabs deskband works with ExplorerPatcher

4. **Wait for Official APIs**
   - Microsoft may eventually provide command bar extensibility
   - Monitor Windows SDK updates

---

## References

### Microsoft Documentation

- [Windows Ribbon Framework](https://learn.microsoft.com/en-us/windows/win32/windowsribbon/windowsribbon-entry)
- [IUIFramework Interface](https://learn.microsoft.com/en-us/windows/win32/api/uiribbon/nn-uiribbon-iuiframework)
- [Extending Explorer (Official)](https://learn.microsoft.com/en-us/windows/win32/shell/extending-the-ribbon)

### Related Projects

- [QTTabBar](https://github.com/indiff/qttabbar) - Advanced Explorer extension
- [ExplorerPatcher](https://github.com/valinet/ExplorerPatcher) - Restore Windows 10 features in Windows 11

### ShellTabs Implementation

- `include/ComVTableHook.h` - COM hooking utilities
- `include/ExplorerRibbonHook.h` - Ribbon hook interface
- `src/ComVTableHook.cpp` - Vtable hooking implementation
- `src/ExplorerRibbonHook.cpp` - Ribbon integration
- `src/CustomRibbonTab.xml` - Ribbon markup definition

---

## Conclusion

Adding a custom tab to Explorer's ribbon is **technically possible** but requires:
- Advanced COM hooking techniques
- Maintenance across Windows updates
- Careful error handling to avoid crashes

For most use cases, Microsoft's **recommended approach is context menus** via `IExplorerCommand`.

The ShellTabs implementation provides:
✅ Framework for ribbon customization
✅ Fallback to context menus
✅ Windows 10 & 11 (with ExplorerPatcher) support
✅ Clean API for registering custom actions

**Use this feature with caution** and thoroughly test on your target Windows versions.
