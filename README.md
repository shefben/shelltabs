# ShellTabs

ShellTabs is a Windows Explorer deskband extension that adds a lightweight tabbed interface to the classic File Explorer window. The deskband hosts a custom tab control that tracks folder navigation and lets you jump between locations with a single click. A dedicated **+** button creates new tabs duplicating the current folder, and you can close tabs from the context menu.

> **Note**
> Deskband extensions are supported on Windows 10 and earlier. Microsoft removed the legacy toolbar surface from the Windows 11 File Explorer; on Windows 11 the extension must be hosted inside an alternative shell (e.g., [ExplorerPatcher](https://github.com/valinet/ExplorerPatcher)) to be visible.

## Features

- Custom Explorer toolbar (deskband) hosting a Win32 tab control.
- Automatic tab creation when you navigate to new folders.
- **+** button to open a new tab that duplicates the currently selected folder.
- Right-click a tab to close it.
- Tabs persist for the lifetime of the Explorer window and track the active folder.
- Session persistence reloads your tabs and islands after Explorer restarts, keeping group collapse state and ordering intact.
- Dragging tabs or islands shows a translucent preview under the cursor so you can place them precisely before dropping.

## Drag and Drop

- Dropping files or folders onto an existing tab continues to delegate to Explorer's copy/move workflow. Hold **Shift** to
  force a move, mirroring Explorer's default behavior; these drops are logged so you can confirm the fallback path was used.
- Drag directories onto the empty portions of the tab strip to open each folder in a background tab—even when other tabs are
  visible. Hold **Ctrl** to focus the first new tab in the foreground, or **Shift** to skip tab creation and reuse the
  Explorer copy/move routine instead.
- If the payload does not contain directories or no tab target is available, ShellTabs automatically falls back to Explorer's
  native handling so files are not lost.

## Building

The project is built as an in-process COM DLL using CMake and the Microsoft Visual C++ toolchain.

1. Open a **x64 Native Tools Command Prompt for VS 2022** (or the version that matches your compiler).
2. Configure and build the project:

   ```powershell
   cd path\to\shelltabs
   cmake -B build -S . -G "Visual Studio 17 2022" -A x64
   cmake --build build --config Release
   ```

   The compiled `ShellTabs.dll` is generated under `build\bin\Release`.

## Installation

1. Copy `ShellTabs.dll` to a permanent location (for example, `C:\Program Files\ShellTabs`).
2. Register the deskband from an elevated Developer Command Prompt or PowerShell session:

   ```powershell
   regsvr32 "C:\Program Files\ShellTabs\ShellTabs.dll"
   ```

   Registration writes entries to **HKCU**, so administrator rights are not required if you keep the DLL under your user profile. Use `/u` to unregister later:

   ```powershell
   regsvr32 /u "C:\Program Files\ShellTabs\ShellTabs.dll"
   ```

3. Restart Windows Explorer (e.g., from Task Manager) if the toolbar menu does not refresh automatically.
4. In File Explorer, right-click the toolbar area and enable **Shell Tabs** from the **Toolbars** menu. The tab strip appears docked at the top of the Explorer window.

An ATL-style registration script lives at `registration/ShellTabs.rgs`. You can feed it to `regsvr32 /c` or integrate it into installer tooling if you prefer declarative registration over the DLL’s self-registration entry points.

## Architecture Overview

- **Tab model (`TabManager`)** – Maintains the list of open tabs, their display names, and associated PIDLs. The manager owns the shell item identifiers, ensuring proper lifetime management.
- **Browser Helper Object (`CExplorerBHO`)** – A lightweight `IObjectWithSite` implementation that customizes Explorer’s UI, tracks navigation events, and surfaces Shell Tabs functionality inside each process.
- **Utilities** – Helper functions for cloning PIDLs, resolving display names, and querying the active folder via `IShellBrowser`/`IWebBrowser2`.

## Development Tips

- Building the project in **Debug** mode under Visual Studio loads the extension into the Explorer process. Use the “Restart Explorer” gesture cautiously; any unhandled exception will terminate Explorer.
- To simplify iterative development, keep a separate PowerShell window with the following aliases:

  ```powershell
  function Register-ShellTabs { regsvr32 /s "C:\Path\To\Build\bin\Debug\ShellTabs.dll" }
  function Unregister-ShellTabs { regsvr32 /s /u "C:\Path\To\Build\bin\Debug\ShellTabs.dll" }
  ```

  Run `Unregister-ShellTabs` before rebuilding if the module is loaded by Explorer.
- The deskband stores data only in memory. You can extend `TabManager` to persist sessions or implement advanced behaviors such as dragging tabs, reordering, or opening new folders in background tabs.
- When experimenting with list-view coloring, remember that `LVM_SETTEXTCOLOR`/`ListView_SetTextColor` only affects non-selected rows. Explorer’s custom draw handler is still required to recolor highlighted or hot items, so keep the `NM_CUSTOMDRAW` path in place even if you also call the global color APIs to set a baseline.

## Troubleshooting

- If `regsvr32` reports that the DLL is in use, ensure Explorer is not holding on to an older build. Kill and restart `explorer.exe`, unregister, then register the new version.
- When the toolbar does not appear in the Explorer toolbar menu, confirm that **Classic toolbars** are enabled (they are hidden by default on Windows 10 with ribbon mode). Enable “Show title bar” from Folder Options or use a shell such as [OldNewExplorer](https://www.msfn.org/board/topic/170375-oldnewexplorer-119/).
- Use the **Event Viewer** or tools like [DebugView](https://learn.microsoft.com/sysinternals/downloads/debugview) to trace diagnostics added to the project if you extend it with logging.

