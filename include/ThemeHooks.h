#pragma once

#include <windows.h>

#include <mutex>
#include <vector>

namespace shelltabs {

enum class ExplorerSurfaceKind;
struct GlowColorSet;

class ExplorerGlowCoordinator;

class ThemeHooks {
public:
    static ThemeHooks& Instance();

    void AttachCoordinator(ExplorerGlowCoordinator& coordinator);
    void DetachCoordinator(const ExplorerGlowCoordinator& coordinator);
    void NotifyCoordinatorUpdated();

    bool IsActive() const noexcept;
    bool IsSurfaceHookActive(ExplorerSurfaceKind kind) const noexcept;

private:
    using DrawThemeBackgroundFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, LPCRECT, LPCRECT);
    using DrawThemeEdgeFn = HRESULT(WINAPI*)(HTHEME, HDC, int, int, LPCRECT, UINT, UINT, LPRECT);

    ThemeHooks() = default;
    ~ThemeHooks() = default;

    ThemeHooks(const ThemeHooks&) = delete;
    ThemeHooks& operator=(const ThemeHooks&) = delete;

    static HRESULT WINAPI HookedDrawThemeBackground(HTHEME theme, HDC dc, int partId, int stateId,
                                                    LPCRECT rect, LPCRECT clipRect);
    static HRESULT WINAPI HookedDrawThemeEdge(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect,
                                              UINT edge, UINT flags, LPRECT contentRect);

    bool OnDrawThemeBackground(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect, LPCRECT clipRect);
    bool OnDrawThemeEdge(HTHEME theme, HDC dc, int partId, int stateId, LPCRECT rect, UINT edge, UINT flags,
                         LPRECT contentRect);

    void UpdateActivationLocked();
    bool InstallLocked();
    void UninstallLocked();
    bool HookModuleImportsLocked(HMODULE module);
    bool ResolveColorsForHook(ExplorerSurfaceKind kind, GlowColorSet& colors);
    bool ExpectHookFor(ExplorerSurfaceKind kind) const noexcept;

    struct PatchedImport {
        HMODULE module = nullptr;
        FARPROC* slot = nullptr;
        FARPROC original = nullptr;
    };

    DrawThemeBackgroundFn m_originalDrawThemeBackground = nullptr;
    DrawThemeEdgeFn m_originalDrawThemeEdge = nullptr;

    ExplorerGlowCoordinator* m_coordinator = nullptr;
    bool m_active = false;

    bool m_expectScrollbar = false;
    bool m_expectToolbar = false;
    bool m_expectRebar = false;

    bool m_scrollbarHookEngaged = false;
    bool m_toolbarHookEngaged = false;
    bool m_rebarHookEngaged = false;

    std::vector<PatchedImport> m_backgroundPatches;
    std::vector<PatchedImport> m_edgePatches;

    mutable std::mutex m_mutex;
};

}  // namespace shelltabs

