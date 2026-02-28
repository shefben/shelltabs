#pragma once

#include <windows.h>
#include <cstddef>

#include "ExplorerGlowSurfaces.h"

namespace shelltabs {

bool InitializeThemeHooks();
void ShutdownThemeHooks();

bool AreThemeHooksActive() noexcept;

void RegisterThemeSurface(HWND hwnd, ExplorerSurfaceKind kind, ExplorerGlowCoordinator* coordinator) noexcept;
void UnregisterThemeSurface(HWND hwnd) noexcept;
void RegisterDirectUiHost(HWND hwnd) noexcept;
void UnregisterDirectUiHost(HWND hwnd) noexcept;
void RegisterDirectUiRenderInterface(void* element, size_t drawIndex, HWND host,
                                     ExplorerGlowCoordinator* coordinator) noexcept;
void InvalidateScrollbarMetrics(HWND hwnd) noexcept;

// Callback invoked by FillRectDetour after the fill completes (excluding re-entrant calls).
// Used by FolderBackgroundHooks to draw over the fill without needing its own FillRect hook.
using FillRectPostCallback = void (*)(HDC dc, const RECT* rect) noexcept;
void SetFillRectPostCallback(FillRectPostCallback cb) noexcept;

class ThemePaintOverrideGuard {
public:
    ThemePaintOverrideGuard(HWND window, ExplorerSurfaceKind kind, GlowColorSet colors,
                            bool suppressFallback) noexcept;
    ~ThemePaintOverrideGuard();

    ThemePaintOverrideGuard(const ThemePaintOverrideGuard&) = delete;
    ThemePaintOverrideGuard& operator=(const ThemePaintOverrideGuard&) = delete;

    ThemePaintOverrideGuard(ThemePaintOverrideGuard&& other) noexcept;
    ThemePaintOverrideGuard& operator=(ThemePaintOverrideGuard&& other) noexcept;

private:
    bool m_active = false;
};

}  // namespace shelltabs
