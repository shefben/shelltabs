# Theme Transition Flow

ShellTabs listens for Windows theme changes without polling the registry. The deskband registers
three sources of truth when the `TabBandWindow` is created:

1. **UISettings color notifications** via `Windows::UI::ViewManagement::UISettings::ColorValuesChanged`.
   When the user flips between light and dark modes (or adjusts accent colors), the event fires on a
   background thread. The `ThemeNotifier` posts a `WM_SHELLTABS_THEME_CHANGED` message back to the
   band window so the palette refresh happens on the UI thread.
2. **Session change broadcasts** using `WTSRegisterSessionNotification`. Explorer does not restart
   when a remote session connects, disconnects, or unlocks, but those transitions often invalidate
   the system palette. Receiving `WM_WTSSESSION_CHANGE` prompts the band to rebuild colors.
3. **Classic Win32 broadcasts** (`WM_SETTINGCHANGE`, `WM_SYSCOLORCHANGE`, and
   `WM_THEMECHANGED`). These continue to call `RefreshTheme()` as a safety net for legacy
   transitions.

During `RefreshTheme()` the band:

- Snapshots the latest `UISettings` foreground/background colors and determines whether immersive
  dark mode should apply (skipping it entirely while Windows high-contrast mode is active).
- Samples Explorer's rebar chrome through the parent window DC to derive gradient start/end colors,
  aligning ShellTabs with the host frame.
- Regenerates the palette, respecting user options unless high-contrast mode forces system colors,
  and updates the new-tab button/close glyph rendering so contrast ratios stay compliant.

The new integration test `ShellTabsThemeTransitionTests` exercises the notifier by simulating
`UISettings` color changes and session events while keeping Explorer (and the band window) alive.
