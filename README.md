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

<a href="https://github.com/shefben/shelltabs/blob/main/images/shelltabs_2026_04_26.png?raw=true" target="_blank">
    <img width="243" height="32" alt="tab_progressbar" src="https://github.com/shefben/shelltabs/blob/main/images/shelltabs_2026_04_26.png?raw=true" />
</a>
## Features

### Tabs
- Create, close, clone, pin, hide/unhide, and detach tabs
- Close other tabs, close tabs to the left, close tabs to the right
- Reopen recently closed tabs with full undo history (groups can be undone too)
- Per-tab back/forward navigation history with dropdown history menus on the toolbar travel buttons
- Per-tab scroll position preservation across tab switches and across session restore
- Tab thumbnail previews on hover, captured from the live shell view
- Keyboard shortcuts: Ctrl+T (new tab), Ctrl+W (close), Ctrl+Tab / Ctrl+Shift+Tab (cycle), Ctrl+1-9 (jump to tab)
- Configurable new-tab template: duplicate current folder, This PC, custom path, or load a saved group
- Configurable tab strip dock mode: top, bottom, left, right, or automatic
- Per-tab progress indicator that mirrors Explorer's active file operation on that tab's folder

### Tab Groups (Islands)
- Organize tabs into named, collapsible groups (islands)
- Customizable group outline colors and styles (solid, dashed, dotted)
- Show/hide group headers per group
- Drag tabs between groups, reorder tabs within a group, or reorder groups themselves
- Move a tab into a brand new island in one drop
- Detach a group into a new Explorer window
- Close all tabs in a group at once with a single undo entry

### Saved Groups
- Save a group configuration by name for reuse across windows and sessions
- Load saved groups from the tab strip context menu or via the new-tab template
- Edits to a saved group (rename, color, outline style, member folders) sync automatically to all windows using it
- Member folders are mirrored into `groups.db` whenever you add or remove a tab from a saved group, so the saved group always reflects its current contents
- Manage saved groups from the Options dialog: add, edit, remove, and edit the folder list directly

### Session Persistence
- All tabs, groups, selection state, and per-tab scroll positions are saved to disk continuously
- Sessions survive both clean Explorer exits and crashes
- Multi-window support: each Explorer window's tabs are restored independently
- Crash detection via active-session marker file with safe checkpoint-based recovery
- "Reuse existing window" option: when launching a new Explorer window, redirect the navigation to an already-open ShellTabs window as a new tab

### Drag and Drop
- Reorder tabs within and across groups by dragging
- Drag tabs or groups between separate Explorer windows
- Drop files/folders onto a tab to trigger Explorer's copy/move workflow (hold Shift to force move)
- Drop directories onto empty tab strip area to open each as a new background tab (hold Ctrl to focus the first)
- Translucent drag preview under the cursor during reorder operations

### Visual Customization
- Full light and dark theme support with automatic detection that follows the Windows app theme
- Breadcrumb bar gradient overlay with customizable colors, opacity, font color gradient, highlight intensity, and dropdown alpha multiplier
- Address bar edit field gradient that matches the breadcrumb theme
- Per-tab progress bar gradient that mirrors Explorer's file operation progress (customizable colors)
- Neon glow effects rendered on Explorer panes via DWM / DirectComposition, with per-surface toggles for list view, column header, rebar, toolbar, address bar, scrollbars, popup menus, and tooltips
- Per-surface glow color modes: Explorer accent, solid, or gradient
- Configurable active/inactive tab colors, hover highlights, and close button behavior
- Status bar dark theming for the bottom status strip plus the container directly above it
- Bitmap interception toggle for advanced glow rendering paths (can be disabled if it impacts performance)
- File / folder gradient font option for the listview labels

### Folder Background Images
- Render a custom background image inside any Explorer folder view (composited above the default chrome via a layered overlay over `DirectUIHWND`)
- **Universal background:** one image used as a fallback for every folder
- **Per-folder backgrounds:** assign a different image to each folder path; managed from the Backgrounds page in the Options dialog
- Adjustable opacity (0-100%) shared across both universal and per-folder backgrounds
- Image position modes: bottom right (default), bottom left, top left, top right, center, stretch to fill, zoom-fill, and tile
- Built-in cache manager and a "Clean Up" button to remove cached image copies for entries that have been deleted
- Light and dark theme aware so the overlay does not wash out the underlying file labels

### Web / HTTP Folders
- Browse remote HTTP/HTTPS directory listings as native Explorer folders via a custom `IShellFolder`
- HTML directory parser handles Apache `mod_autoindex`, Nginx `autoindex`, and table-based custom autoindex layouts (e.g. Myrient)
- Per-site display name, enable/disable toggle, optional parallel downloads (1-16 concurrent), and an optional per-site download speed limit (KB/s)
- Configure web folder sites from the Web Folders page in the Options dialog

### FTP Support
- Built-in FTP client with a custom `IShellFolder` implementation
- Browse FTP servers directly inside Explorer as native folder locations
- Configurable FTP site list with per-site display name, host, user, port, and enable/disable toggle
- Anonymous and credentialed connections (credentials managed via Windows Credential Manager)

### Context Menu Integration
- Right-click a tab for quick actions: close, clone, pin, hide, detach, open terminal, open VS Code, copy path
- Access the full Explorer shell context menu for any tab's folder, with `IContextMenu2` / `IContextMenu3` message forwarding for owner-drawn shell extensions
- Add custom context menu entries from the Options dialog with configurable executables, arguments, working directory, run-as-admin, window state, and confirmation prompts
- Submenus, separators, icons (file path or `shell32.dll,index` style), and selection-based visibility rules (file vs folder, count constraints, wildcard file patterns, folder path filters)
- Custom items can be anchored to specific positions in the shell menu (top, bottom, before/after shell items, default)

### Taskbar Integration
- Custom taskbar tab list popup that appears when hovering an Explorer taskbar button, showing each window's tabs with icons and titles instead of (or alongside) the default thumbnail
- Click a tab in the popup to switch to that window and activate the tab in one step
- Light/dark themed to match the system theme

### Options Dialog
- **General:** New-tab behavior, session restore settings, dock mode, reuse-existing-window toggle, persist-group-paths toggle
- **Appearance:** Breadcrumb gradient (font + background, custom colors, transparency, highlight/dropdown alpha), progress bar gradient colors, custom tab selected/unselected colors
- **Glow Effects:** Master enable, per-surface toggles, custom primary/secondary glow colors, gradient blend, Explorer accent override, bitmap intercept and gradient font toggles
- **Backgrounds:** Enable/disable, universal image picker, per-folder image list (add/edit/remove/clean up), opacity slider, position mode radio buttons
- **Context Menus:** Add, edit, remove, and reorder custom right-click menu entries with full visibility and command configuration
- **Groups:** Manage saved groups (add/edit/remove), edit each group's name, color, outline style, and folder list
- **Web Folders:** Add, edit, remove HTTP/HTTPS directory sites with per-site download settings

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
