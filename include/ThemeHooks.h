#pragma once

#include <windows.h>

#include "ExplorerGlowSurfaces.h"

namespace shelltabs {

bool InitializeThemeHooks();
void ShutdownThemeHooks();

void RegisterThemeSurface(HWND hwnd, ExplorerSurfaceKind kind, ExplorerGlowCoordinator* coordinator) noexcept;
void UnregisterThemeSurface(HWND hwnd) noexcept;

}  // namespace shelltabs
