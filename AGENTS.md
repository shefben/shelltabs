# ShellTabs Agent Guidelines

This repository contains **ShellTabs**, a Windows Explorer deskband implemented in modern C++20 with Win32/COM APIs. Follow the guidelines below when making changes anywhere in this repo.

## Platform & Compatibility
- **Important:** Target **only Windows 10 and Windows 11 File Explorer**. Do not spend time on earlier versions and avoid code paths that cater to unsupported operating systems.
- Treat Windows 11 as running the legacy Explorer surface via shells such as ExplorerPatcher; the code should remain compatible with the Windows 10 implementation.
- **Important:** Always validate behavior in both the **default (light)** and **dark** themes provided by Windows 10/11. Any visual change must work for both modes.

## UI & UX Expectations
- **Important:** When introducing or modifying GUI scrollbars, ensure they are operable via both the **mouse scroll wheel** and traditional interaction (dragging the thumb, clicking the track/buttons).
- Prefer replacing or extending existing Explorer UI elements (e.g., the breadcrumb bar) **in-place** rather than overlaying new widgets on top of the stock controls. This keeps accessibility, theming, and hit-testing consistent.

## Coding Practices
- Use C++20 with the Win32/COM stack already established in the project (`IDeskBand2`, `IObjectWithSite`, ATL-style registration scripts, etc.).
- Observe the compiler settings already configured in `CMakeLists.txt` (`/permissive-`, `/W4`, `_WIN32_WINNT=0x0A00`). Keep new code warning-free under these flags.
- Manage PIDL lifetimes and COM reference counting carefully—prefer RAII helpers where possible.
- Follow the project’s existing module organization (deskband, tab model, options dialog, FTP helpers, etc.). Place new code in the most appropriate component rather than creating ad-hoc files.
- Favor explicit Unicode-aware Win32 APIs and avoid ANSI variants.

## Building & Testing
- Build with the **x64 Native Tools Command Prompt for VS 2022** (or matching toolset) using the provided CMake configuration.
- The shared library target is `ShellTabs.dll`; registration occurs through `regsvr32` or the supplied `.rgs` script.
- If you add automated tests, gate them behind the `SHELLTABS_BUILD_TESTS` option and ensure they compile with the same Windows 10 SDK level.

## Research & Documentation
- The root document `Hooking Windows Explorer’s Tree and List View Panes for Custom Drawing.docx` describes the hooking strategy for Explorer panes. Review it before modifying hooks or the coloring/highlighting pipeline.
- Keep the existing `TODO.md` priorities in mind when implementing features (tab management, options dialog, pane customization).

## Miscellaneous
- Preserve the project’s logging and diagnostics patterns when adding instrumentation.
- Ensure session persistence features continue to function if you change tab or group storage.

