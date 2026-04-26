# ShellTabs

A Windows Explorer deskband extension that adds tabbed browsing, tab groups, session persistence, visual customization, and more to the classic File Explorer window.

> **Note:** Deskband extensions require the legacy Explorer toolbar surface. On **Windows 10** this works out of the box. On **Windows 11**, use [ExplorerPatcher](https://github.com/valinet/ExplorerPatcher) or a similar tool to restore classic toolbars.

## Previews

<a href="https://github.com/user-attachments/assets/79bf6c67-6918-4f27-b9f6-bae69ee383b6" target="_blank">
    <img src="https://github.com/user-attachments/assets/79bf6c67-6918-4f27-b9f6-bae69ee383b6" width="200" alt="shelltabs_island_example">
</a>
<a href="https://github.com/user-attachments/assets/c117f3e5-ddf7-478d-8f98-11cc3d40b7ba" target="_blank">
    <img src="https://github.com/user-attachments/assets/c117f3e5-ddf7-478d-8f98-11cc3d40b7ba" width="200" alt="shelltabs_main">
</a>
<a href="https://github.com/user-attachments/assets/7f620497-967e-4fff-a7fe-076fcd5cbd09" target="_blank">
    <img src="https://github.com/user-attachments/assets/7f620497-967e-4fff-a7fe-076fcd5cbd09" width="200" alt="shelltabs_options_appearance">
</a>
<a href="https://github.com/user-attachments/assets/3035e227-4aa5-4b30-b5cb-0725a5666d3a" target="_blank">
    <img src="https://github.com/user-attachments/assets/3035e227-4aa5-4b30-b5cb-0725a5666d3a" width="200" alt="shelltabs_options_context_menus">
</a>
<a href="https://github.com/user-attachments/assets/da8b2866-6da6-4803-8a99-ef6c89805e0f" target="_blank">
    <img src="https://github.com/user-attachments/assets/da8b2866-6da6-4803-8a99-ef6c89805e0f" width="200" alt="shelltabs_options_general">
</a>
<a href="https://github.com/user-attachments/assets/ccee2abe-d6d8-4918-903a-5da6b08e0ad3" target="_blank">
    <img src="https://github.com/user-attachments/assets/ccee2abe-d6d8-4918-903a-5da6b08e0ad3" width="200" alt="shelltabs_options_groups">
</a>
<a href="https://github.com/user-attachments/assets/6dcb5df4-452e-4d3e-9427-84a331df2af4" target="_blank">
    <img width="243" height="32" alt="tab_progressbar" src="https://github.com/user-attachments/assets/6dcb5df4-452e-4d3e-9427-84a331df2af4" />
</a>

<a href="https://github.com/shefben/shelltabs/blob/main/images/shelltabs_2026_04_26.png" target="_blank">
    <img width="200" alt="tab_progressbar" src="https://github.com/shefben/shelltabs/blob/main/images/shelltabs_2026_04_26.png" />
</a>
<br>
## Features

### Tabs
- Create, close, clone, pin, hide/unhide, and detach tabs
- Close other tabs, close tabs to the left/right
- Reopen recently closed tabs with full undo history
- Per-tab back/forward navigation history
- Per-tab scroll position preservation across tab switches
- Tab thumbnail previews on hover
- Keyboard shortcuts: Ctrl+T (new tab), Ctrl+W (close), Ctrl+Tab / Ctrl+Shift+Tab (cycle), Ctrl+1-9 (jump to tab)

### Tab Groups (Islands)
- Organize tabs into named, collapsible groups
- Customizable group outline colors and styles (solid, dashed, dotted)
- Show/hide group headers
- Drag tabs between groups or reorder groups themselves
- Detach a group into a new Explorer window
- Close all tabs in a group at once

### Saved Groups
- Save a group configuration by name for reuse across windows and sessions
- Load saved groups from the tab strip context menu
- Edits to a saved group sync automatically to all windows using it
- Manage saved groups from the Options dialog

### Session Persistence
- All tabs, groups, and selection state are saved to disk continuously
- Sessions survive both clean Explorer exits and crashes
- Multi-window support: each Explorer window's tabs are restored independently
- Crash detection via active-session marker file

### Drag and Drop
- Reorder tabs within and across groups by dragging
- Drag tabs or groups between separate Explorer windows
- Drop files/folders onto a tab to trigger Explorer's copy/move workflow (hold Shift to force move)
- Drop directories onto empty tab strip area to open each as a new background tab (hold Ctrl to focus the first)
- Translucent drag preview under the cursor during reorder operations

### Visual Customization
- Full light and dark theme support with automatic detection
- Breadcrumb bar gradient overlay (customizable colors and opacity)
- Address bar edit field gradient
- Per-tab progress bar gradient (mirrors Explorer's file operation progress)
- Neon glow effects on Explorer panes via DWM/DirectComposition
- Custom folder background images per directory path
- Configurable active/inactive tab colors, hover highlights, and close button behavior

### Context Menu Integration
- Right-click a tab for quick actions: close, clone, pin, hide, detach, open terminal, open VS Code, copy path
- Access the full Explorer shell context menu for any tab's folder
- Add custom context menu entries with configurable commands and arguments from the Options dialog

### FTP Support
- Built-in FTP client with a custom `IShellFolder` implementation
- Browse FTP servers directly inside Explorer as native folder locations

### Options Dialog
- **General:** New-tab behavior, session restore settings, dock mode
- **Appearance:** Breadcrumb gradient, glow surfaces, folder backgrounds, progress bar colors
- **Groups:** Manage saved groups, edit names/colors/outlines
- **Context Menus:** Add, edit, and remove custom right-click menu entries

## Building

Requires **Visual Studio 2022** with the C++ desktop workload.

```bat
cmake -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Release
```

Output: `build\bin\Release\ShellTabs.dll`

To build with tests:

```bat
cmake -B build -DSHELLTABS_BUILD_TESTS=ON
cmake --build build --config Release
```

## Installation

1. Copy `ShellTabs.dll` to a permanent location.
2. Register from an elevated command prompt:
   ```bat
   regsvr32 "C:\Path\To\ShellTabs.dll"
   ```
3. Restart Explorer if the toolbar menu does not refresh.
4. Right-click the Explorer toolbar area and enable **Shell Tabs** from the **Toolbars** menu.

To unregister:
```bat
regsvr32 /u "C:\Path\To\ShellTabs.dll"
```

## Architecture

| Component | Description |
|---|---|
| **TabBand** | Root COM object (`IDeskBand2`). Orchestrates tab operations, session save/restore, and browser event handling. |
| **TabBandWindow** | Win32 window that renders the tab strip. Handles mouse/keyboard input, drag-and-drop, context menus, and incremental layout diffing. |
| **TabManager** | Pure data model for tabs and groups. No UI. Tracks selection, activation order, navigation history, scroll positions, and pinned/hidden state. |
| **CExplorerBHO** | Browser Helper Object. Subclasses Explorer's breadcrumb bar, address edit, progress bar, travel band, status bar, and frame window. |
| **SessionCoordinator** | Process-wide singleton managing a single session file for all Explorer windows. Handles crash detection, atomic writes, and multi-window coordination. |
| **OptionsStore** | Thread-safe singleton for persisting user settings to JSON in AppData. |
| **ThemeHooks** | Hooks into `uxtheme.dll` to detect light/dark theme transitions in real time. |
| **FtpShellFolder** | Custom `IShellFolder` exposing FTP servers as browsable Explorer locations. |

## Troubleshooting

- **DLL in use:** Kill `explorer.exe` from Task Manager, unregister the old DLL, then register the new one and restart Explorer.
- **Toolbar not visible:** Ensure classic toolbars are enabled. On Windows 10 with ribbon mode, enable "Show title bar" from Folder Options. On Windows 11, install ExplorerPatcher.
- **Logs:** Check `%LOCALAPPDATA%\ShellTabs\Logs\` for diagnostic output.
- **Settings:** Stored in `%APPDATA%\ShellTabs\`. Delete this folder to reset all options and session data.
