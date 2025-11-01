#pragma once

#include <windows.h>
#include <uxtheme.h>

#include <mutex>
#include <vector>

namespace shelltabs {

enum class ExplorerSurfaceKind;
struct GlowColorSet;

class ExplorerGlowCoordinator;

class ThemeHooks {
public:
    static ThemeHooks& Instance();

    bool Initialize();
    void Shutdown();

    void AttachCoordinator(ExplorerGlowCoordinator& coordinator);
    void DetachCoordinator(const ExplorerGlowCoordinator& coordinator);
    void NotifyCoordinatorUpdated();

    bool IsActive() const noexcept;
    bool IsSurfaceHookActive(ExplorerSurfaceKind kind) const noexcept;

private:
    using DrawThemeBackgroundFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, LPCRECT, LPCRECT);
    using DrawThemeEdgeFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, LPCRECT, UINT, UINT, LPRECT);
    using FillRectFn = int(WINAPI*)(HDC, const RECT*, HBRUSH);
    using GdiGradientFillFn = BOOL(WINAPI*)(HDC, PTRIVERTEX, ULONG, PVOID, ULONG, ULONG);

    ThemeHooks() = default;
    ~ThemeHooks() = default;

    ThemeHooks(const ThemeHooks&) = delete;
    ThemeHooks& operator=(const ThemeHooks&) = delete;

    static HRESULT WINAPI HookedDrawThemeBackground(HTHEME theme, HDC dc, int partId, int stateId,
                                                    LPCRECT rect, LPCRECT clipRect);
    static HRESULT WINAPI HookedDrawThemeEdge(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect,
                                              UINT edge, UINT flags, LPRECT contentRect);
    static int WINAPI HookedFillRect(HDC dc, const RECT* rect, HBRUSH brush);
    static BOOL WINAPI HookedGdiGradientFill(HDC dc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh,
                                             ULONG meshCount, ULONG mode);

    bool OnDrawThemeBackground(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect, LPCRECT clipRect);
    bool OnDrawThemeEdge(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect, UINT edge, UINT flags,
                         LPRECT contentRect);
    bool OnFillRect(HDC dc, const RECT& rect, HBRUSH brush);
    bool OnGradientFill(HDC dc, PTRIVERTEX vertices, ULONG vertexCount, PVOID mesh, ULONG meshCount, ULONG mode);

    void UpdateActivationLocked();
    bool InstallLocked();
    void UninstallLocked();
    bool EnsureHooksInitializedLocked();
    void DisableHooksLocked();
    void DestroyHooksLocked();
    bool ResolveColorsForHook(ExplorerSurfaceKind kind, GlowColorSet& colors);
    bool ExpectHookFor(ExplorerSurfaceKind kind) const noexcept;

    DrawThemeBackgroundFn m_originalDrawThemeBackground = nullptr;
    DrawThemeEdgeFn m_originalDrawThemeEdge = nullptr;
    FillRectFn m_originalFillRect = nullptr;
    GdiGradientFillFn m_originalGradientFill = nullptr;

    void* m_drawThemeBackgroundTarget = nullptr;
    void* m_drawThemeEdgeTarget = nullptr;
    void* m_fillRectTarget = nullptr;
    void* m_gradientFillTarget = nullptr;

    ExplorerGlowCoordinator* m_coordinator = nullptr;
    bool m_active = false;

    bool m_hooksInitialized = false;
    bool m_hooksEnabled = false;

    bool m_expectScrollbar = false;
    bool m_expectToolbar = false;
    bool m_expectRebar = false;
    bool m_expectHeader = false;
    bool m_expectEdit = false;

    bool m_scrollbarHookEngaged = false;
    bool m_toolbarHookEngaged = false;
    bool m_rebarHookEngaged = false;
    bool m_headerHookEngaged = false;
    bool m_editHookEngaged = false;

    mutable std::mutex m_mutex;
};

}  // namespace shelltabs

