#pragma once

#include <windows.h>

#include <mutex>
#include <vector>

namespace shelltabs {

class ExplorerGlowCoordinator;

class ThemeHooks {
public:
    static ThemeHooks& Instance();

    void AttachCoordinator(ExplorerGlowCoordinator& coordinator);
    void DetachCoordinator(const ExplorerGlowCoordinator& coordinator);
    void NotifyCoordinatorUpdated();

    bool IsActive() const noexcept;

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
    bool OnDrawThemeEdge(int partId) const;

    void UpdateActivationLocked();
    bool InstallLocked();
    void UninstallLocked();
    bool HookModuleImportsLocked(HMODULE module);

    struct PatchedImport {
        HMODULE module = nullptr;
        FARPROC* slot = nullptr;
        FARPROC original = nullptr;
    };

    DrawThemeBackgroundFn m_originalDrawThemeBackground = nullptr;
    DrawThemeEdgeFn m_originalDrawThemeEdge = nullptr;

    ExplorerGlowCoordinator* m_coordinator = nullptr;
    bool m_active = false;

    std::vector<PatchedImport> m_backgroundPatches;
    std::vector<PatchedImport> m_edgePatches;

    mutable std::mutex m_mutex;
};

}  // namespace shelltabs

