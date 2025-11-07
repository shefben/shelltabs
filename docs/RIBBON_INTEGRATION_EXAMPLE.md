# Integrating Custom Ribbon Tab with CExplorerBHO

This guide shows how to integrate the custom ribbon functionality with ShellTabs' existing Browser Helper Object (BHO).

## Quick Start Integration

### Step 1: Update CExplorerBHO Header

Add the ribbon hook include to `include/CExplorerBHO.h`:

```cpp
// Add this with other includes
#include "ExplorerRibbonHook.h"
```

### Step 2: Initialize Ribbon Hooks in BHO

Modify `src/CExplorerBHO.cpp` to initialize the ribbon hooks when the BHO loads:

```cpp
// In CExplorerBHO::SetSite() method, after existing initialization:

STDMETHODIMP CExplorerBHO::SetSite(IUnknown* pUnkSite) {
    // ... existing code ...

    if (pUnkSite) {
        // Site is being set (BHO is loading)

        // ... existing initialization code ...

        // Initialize custom ribbon hooks
        if (!shelltabs::ExplorerRibbonHook::IsEnabled()) {
            if (shelltabs::ExplorerRibbonHook::Initialize()) {
                LogMessage(LogLevel::Info, L"CExplorerBHO: Ribbon hooks initialized");

                // Register custom button actions
                RegisterCustomRibbonButtons();
            } else {
                LogMessage(LogLevel::Warning, L"CExplorerBHO: Failed to initialize ribbon hooks");
            }
        }
    } else {
        // Site is being cleared (BHO is unloading)

        // ... existing cleanup code ...

        // Shutdown ribbon hooks
        if (shelltabs::ExplorerRibbonHook::IsEnabled()) {
            shelltabs::ExplorerRibbonHook::Shutdown();
            LogMessage(LogLevel::Info, L"CExplorerBHO: Ribbon hooks shut down");
        }
    }

    return S_OK;
}
```

### Step 3: Register Custom Button Actions

Add a helper method to register button callbacks:

```cpp
// In CExplorerBHO class (private methods section):

void CExplorerBHO::RegisterCustomRibbonButtons() {
    using shelltabs::ExplorerRibbonHook;
    using shelltabs::cmdCustomButton1;
    using shelltabs::cmdCustomButton2;
    using shelltabs::cmdCustomButton3;
    using shelltabs::cmdCustomButton4;
    using shelltabs::cmdCustomButton5;

    // Button 1: New Tab
    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton1, [this](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Ribbon: New Tab button clicked");
        // Integrate with existing tab manager
        this->OpenNewTab();
    });

    // Button 2: Refresh Current Tab
    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton2, [this](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Ribbon: Refresh Tab button clicked");
        // Refresh current folder view
        if (m_shellBrowser) {
            Microsoft::WRL::ComPtr<IShellView> shellView;
            if (SUCCEEDED(m_shellBrowser->QueryActiveShellView(&shellView))) {
                shellView->Refresh();
            }
        }
    });

    // Button 3: Close Current Tab
    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton3, [this](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Ribbon: Close Tab button clicked");
        this->CloseCurrentTab();
    });

    // Button 4: Toggle Hidden Files
    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton4, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Ribbon: Toggle Hidden Files button clicked");

        // Get Explorer window
        HWND explorerWnd = GetAncestor(hwnd, GA_ROOT);
        if (explorerWnd) {
            // Send command to toggle hidden files
            SendMessage(explorerWnd, WM_COMMAND,
                       MAKEWPARAM(41024, 0), 0); // Command ID for show hidden files
        }
    });

    // Button 5: Custom Action (Example)
    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton5, [](HWND hwnd) {
        LogMessage(LogLevel::Info, L"Ribbon: Custom Button 5 clicked");

        MessageBox(hwnd,
                  L"This is a placeholder for your custom action.\n\n"
                  L"You can integrate this with any functionality:\n"
                  L"- Open terminal here\n"
                  L"- Copy path to clipboard\n"
                  L"- Bookmark current folder\n"
                  L"- Or anything else you need!",
                  L"Custom Action",
                  MB_OK | MB_ICONINFORMATION);
    });
}
```

### Step 4: Update CMakeLists.txt

Add the new source files to your build:

```cmake
# In CMakeLists.txt, add to sources:

set(SOURCES
    # ... existing sources ...
    src/ComVTableHook.cpp
    src/ExplorerRibbonHook.cpp
)

set(HEADERS
    # ... existing headers ...
    include/ComVTableHook.h
    include/ExplorerRibbonHook.h
)

# Add required libraries
target_link_libraries(ShellTabs
    # ... existing libraries ...
    propsys.lib
    ole32.lib
)
```

---

## Advanced Integration Examples

### Example 1: Dynamic Button State

Update button enabled/disabled state based on context:

```cpp
// In your BHO, periodically update ribbon state
void CExplorerBHO::UpdateRibbonState() {
    IUIFramework* framework = ExplorerRibbonHook::GetRibbonFramework(m_explorerHwnd);
    if (!framework) return;

    // Disable "Close Tab" button if only one tab is open
    PROPVARIANT var;
    InitPropVariantFromBoolean(GetTabCount() > 1, &var);
    framework->SetUICommandProperty(cmdCustomButton3, UI_PKEY_Enabled, var);
    PropVariantClear(&var);

    // Invalidate to refresh UI
    framework->InvalidateUICommand(cmdCustomButton3, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Enabled);
}
```

### Example 2: Context-Aware Actions

Make buttons behave differently based on selected files:

```cpp
ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton2, [this](HWND hwnd) {
    // Get selected items
    Microsoft::WRL::ComPtr<IShellView> shellView;
    if (SUCCEEDED(m_shellBrowser->QueryActiveShellView(&shellView))) {
        Microsoft::WRL::ComPtr<IDataObject> dataObject;
        if (SUCCEEDED(shellView->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&dataObject)))) {
            // Check if selection contains folders
            bool hasFolders = HasFolders(dataObject.Get());

            if (hasFolders) {
                // Open folders in new tabs
                OpenSelectedFoldersInTabs(dataObject.Get());
            } else {
                // Show properties for files
                ShowPropertiesDialog(dataObject.Get());
            }
        }
    }
});
```

### Example 3: Integration with Existing Tab Manager

Connect ribbon buttons to your tab management system:

```cpp
void CExplorerBHO::RegisterCustomRibbonButtons() {
    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton1, [this](HWND hwnd) {
        // Use existing tab manager
        if (m_tabManager) {
            m_tabManager->CreateNewTab();
        }
    });

    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton2, [this](HWND hwnd) {
        // Navigate to parent folder in current tab
        if (m_tabManager) {
            m_tabManager->NavigateToParent();
        }
    });

    ExplorerRibbonHook::RegisterButtonAction(cmdCustomButton3, [this](HWND hwnd) {
        // Close active tab
        if (m_tabManager) {
            m_tabManager->CloseActiveTab();
        }
    });
}
```

---

## Customizing Button Labels and Icons

### Update Button Labels

Modify `src/ExplorerRibbonHook.cpp` in `RibbonCommandHandler::UpdateProperty()`:

```cpp
if (key == UI_PKEY_Label) {
    switch (commandId) {
    case cmdCustomButton1:
        return InitPropVariantFromString(L"New Tab", newValue);  // ← Change this

    case cmdCustomButton2:
        return InitPropVariantFromString(L"Refresh", newValue);  // ← Change this

    // ... etc
    }
}
```

### Add Button Icons

You'll need to:

1. Create icon resources (32x32 for large, 16x16 for small)
2. Add them to your resource file
3. Update the ribbon XML to reference them
4. Return icon handles in `UpdateProperty()`:

```cpp
if (key == UI_PKEY_LargeImage) {
    switch (commandId) {
    case cmdCustomButton1: {
        HBITMAP hBitmap = LoadBitmap(g_hInstance, MAKEINTRESOURCE(IDB_NEWTAB_LARGE));
        if (hBitmap) {
            // Convert HBITMAP to IUIImage
            Microsoft::WRL::ComPtr<IUIImage> image;
            if (SUCCEEDED(CreateUIImageFromBitmap(hBitmap, &image))) {
                return InitPropVariantFromInterface(image.Get(), newValue);
            }
        }
        break;
    }
    // ... other buttons
    }
}
```

---

## Testing the Integration

### Build and Test

1. **Build the project**:
   ```batch
   cmake --build build --config Debug
   ```

2. **Unregister old DLL** (if previously registered):
   ```batch
   regsvr32 /u "path\to\old\ShellTabs.dll"
   ```

3. **Register new DLL**:
   ```batch
   regsvr32 "build\bin\Debug\ShellTabs.dll"
   ```

4. **Restart Explorer**:
   - Task Manager → "Windows Explorer" → Restart
   - Or: `taskkill /f /im explorer.exe && start explorer.exe`

5. **Open File Explorer** and check for the "Custom" tab

### Enable Logging

To see debug output:

```cpp
// In Logging.h, set log level to Verbose
#define LOG_LEVEL LogLevel::Verbose

// Check log file (if file logging is enabled)
// or use DebugView from Sysinternals
```

### Debugging Tips

- Use **WinDbg** or **Visual Studio** to attach to `explorer.exe`
- Set breakpoints in `RibbonCommandHandler::Execute()`
- Check logs for initialization messages
- Verify `ExplorerRibbonHook::IsEnabled()` returns `true`

---

## Handling Errors Gracefully

### Fallback to Context Menus

If ribbon hooks fail to initialize, fall back to existing functionality:

```cpp
if (!ExplorerRibbonHook::Initialize()) {
    LogMessage(LogLevel::Warning, L"Ribbon hooks failed, using context menus only");
    // Your existing context menu extension still works
}
```

### Version Detection

Different Windows versions may require different approaches:

```cpp
DWORD version = GetWindowsVersion();

if (version >= WINDOWS_11) {
    // Windows 11 uses command bar, not ribbon
    LogMessage(LogLevel::Info, L"Windows 11 detected, ribbon hooks may not work");
    LogMessage(LogLevel::Info, L"Consider using context menus or ExplorerPatcher");
} else if (version >= WINDOWS_10) {
    // Windows 10 has ribbon
    ExplorerRibbonHook::Initialize();
} else {
    // Older Windows versions
    LogMessage(LogLevel::Warning, L"Windows version not supported for ribbon customization");
}
```

---

## Production Considerations

### ⚠️ Important Warnings

1. **Undocumented APIs**
   - This implementation uses undocumented techniques
   - May break with Windows updates
   - Test thoroughly on target OS versions

2. **Stability**
   - COM vtable hooking can crash Explorer if done incorrectly
   - Always validate pointers before use
   - Implement comprehensive error handling

3. **Performance**
   - Hooking adds overhead
   - Keep button callbacks fast
   - Don't block UI thread with long operations

4. **Maintenance**
   - Monitor Windows updates
   - Have fallback mechanisms
   - Consider context menus as primary interface

### Recommended Approach

For production code:

```cpp
// Primary: Context menus (official, supported)
RegisterContextMenuExtension();

// Secondary: Ribbon customization (if available)
if (IsRibbonCustomizationSupported()) {
    ExplorerRibbonHook::Initialize();
} else {
    LogMessage(LogLevel::Info, L"Ribbon customization not available, using context menus");
}

// Tertiary: Deskband (Windows 10 only)
if (IsDeskbandSupported()) {
    RegisterDeskband();
}
```

---

## Summary

You now have:

✅ COM vtable hooking framework (`ComVTableHook`)
✅ Explorer ribbon integration (`ExplorerRibbonHook`)
✅ Ribbon XML definition (`CustomRibbonTab.xml`)
✅ Integration with existing BHO
✅ Example button actions
✅ Error handling and fallbacks

**Next Steps:**

1. Customize button labels and icons
2. Implement your specific actions
3. Test on Windows 10 (and Windows 11 with ExplorerPatcher)
4. Add version detection and fallbacks
5. Consider context menus as primary interface

For more details, see:
- `docs/RIBBON_CUSTOMIZATION.md` - Full documentation
- `include/ExplorerRibbonHook.h` - API reference
- `src/ExplorerRibbonHook.cpp` - Implementation details
