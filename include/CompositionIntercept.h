#pragma once

#include <windows.h>

#include "ExplorerGlowSurfaces.h"

namespace shelltabs {

bool InitializeCompositionIntercept();
void ShutdownCompositionIntercept();

void RegisterCompositionSurface(HWND hwnd, ExplorerGlowCoordinator* coordinator) noexcept;
void UnregisterCompositionSurface(HWND hwnd) noexcept;
void NotifyCompositionColorChange(HWND hwnd) noexcept;

}  // namespace shelltabs

