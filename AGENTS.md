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



<!-- headroom:rtk-instructions -->
# RTK (Rust Token Killer) - Token-Optimized Commands

When running shell commands, **always prefix with `rtk`**. This reduces context
usage by 60-90% with zero behavior change. If rtk has no filter for a command,
it passes through unchanged  so it is always safe to use.

## Key Commands
```bash
# Git (59-80% savings)
rtk git status          rtk git diff            rtk git log

# Files & Search (60-75% savings)
rtk ls <path>           rtk read <file>         rtk grep <pattern>
rtk find <pattern>      rtk diff <file>

# Test (90-99% savings)  shows failures only
rtk pytest tests/       rtk cargo test          rtk test <cmd>

# Build & Lint (80-90% savings)  shows errors only
rtk tsc                 rtk lint                rtk cargo build
rtk prettier --check    rtk mypy                rtk ruff check

# Analysis (70-90% savings)
rtk err <cmd>           rtk log <file>          rtk json <file>
rtk summary <cmd>       rtk deps                rtk env

# GitHub (26-87% savings)
rtk gh pr view <n>      rtk gh run list         rtk gh issue list

# Infrastructure (85% savings)
rtk docker ps           rtk kubectl get         rtk docker logs <c>

# Package managers (70-90% savings)
rtk pip list            rtk pnpm install        rtk npm run <script>
```

## Rules
- In command chains, prefix each segment: `rtk git add . && rtk git commit -m "msg"`
- For debugging, use raw command without rtk prefix
- `rtk proxy <cmd>` runs command without filtering but tracks usage
<!-- /headroom:rtk-instructions -->


<!-- headroom:memory-instructions -->
## Memory

Use the `headroom_memory` MCP server for persistent cross-session knowledge.

**Before** answering questions about prior decisions, conventions, project context,
architecture, user preferences, org info, codenames, debugging history, or anything
from past sessions - call `memory_search` first.

**After** making durable decisions, discovering conventions, or learning important
facts - call `memory_save` to persist them for future sessions.

Memory is your first source of truth for anything not visible in the current conversation.

@RTK.md
