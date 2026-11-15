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
void RegisterDirectUiRenderInterface(void* element, size_t drawIndex, HWND host,
                                     ExplorerGlowCoordinator* coordinator) noexcept;

// Callback invoked by FillRectDetour after the fill completes (excluding re-entrant calls).
// Used by FolderBackgroundHooks to draw over the fill without needing its own FillRect hook.
using FillRectPostCallback = void (*)(HDC dc, const RECT* rect) noexcept;

// Optional filter for commands chosen from popup menus tracked with TPM_RETURNCMD. When it returns
// true the command was handled by ShellTabs and TrackPopupMenu(Ex) reports "no selection" (0).
using PopupMenuCommandFilter = bool (*)(UINT command);
void SetPopupMenuCommandFilter(PopupMenuCommandFilter filter) noexcept;

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
