#pragma once

#include <cstdint>

namespace shelltabs {

struct ShellTabsOptions;
enum class ExplorerSurfaceKind;

// Initializes the global theme hooks with the supplied options. Safe to call
// multiple times; subsequent calls refresh the cached palette without
// reinstalling the hooks.
bool InitializeThemeHooks(const ShellTabsOptions& options);

// Tears down every installed hook and releases any state associated with the
// hook layer.
void ShutdownThemeHooks();

// Refreshes the neon palette used by the hook layer.
void UpdateThemeHooks(const ShellTabsOptions& options);

// Forwards theme change notifications so the hook layer can rebuild any cached
// accent colors or accessibility state.
void NotifyThemeHooksThemeChanged();
void NotifyThemeHooksSettingChanged();

// Returns true when the specified Explorer surface is already repainted by the
// global hook layer and should not be overdrawn by overlay surfaces.
bool ThemeHooksOverrideSurface(ExplorerSurfaceKind kind) noexcept;

}  // namespace shelltabs

