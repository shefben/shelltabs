#pragma once

#include <windows.h>
#include <cstddef>

#include "ExplorerGlowSurfaces.h"

namespace shelltabs {

bool InitializeThemeHooks();
void ShutdownThemeHooks();

void RegisterThemeSurface(HWND hwnd, ExplorerSurfaceKind kind, ExplorerGlowCoordinator* coordinator) noexcept;
void UnregisterThemeSurface(HWND hwnd) noexcept;
void RegisterDirectUiHost(HWND hwnd) noexcept;
void UnregisterDirectUiHost(HWND hwnd) noexcept;
void RegisterDirectUiRenderInterface(void* element, size_t drawIndex, HWND host,
                                     ExplorerGlowCoordinator* coordinator) noexcept;
void InvalidateScrollbarMetrics(HWND hwnd) noexcept;

}  // namespace shelltabs
