#pragma once

namespace shelltabs {

// Installs process-wide hooks for UxTheme text drawing APIs so we can override
// the colors used when Explorer renders list view text.
bool InitializeThemeHooks();

// Restores any hooks installed by InitializeThemeHooks. Safe to call even if
// initialization never succeeded.
void ShutdownThemeHooks();

}  // namespace shelltabs
