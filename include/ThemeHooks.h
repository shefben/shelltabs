#pragma once

#include <windows.h>
#include <cstddef>

#include "ExplorerGlowSurfaces.h"

namespace shelltabs {

struct CreateWindowExInterceptorArgs {
    DWORD exStyle = 0;
    LPCWSTR className = nullptr;
    LPCWSTR windowName = nullptr;
    DWORD style = 0;
    int x = 0;
    int y = 0;
    int width = 0;
    int height = 0;
    HWND parent = nullptr;
    HMENU menu = nullptr;
    HINSTANCE instance = nullptr;
    LPVOID param = nullptr;
};

using CreateWindowExInterceptor = bool (*)(const CreateWindowExInterceptorArgs& args,
                                           HWND* result,
                                           void* context);

bool RegisterCreateWindowExInterceptor(CreateWindowExInterceptor callback, void* context) noexcept;
void UnregisterCreateWindowExInterceptor(CreateWindowExInterceptor callback, void* context) noexcept;

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
