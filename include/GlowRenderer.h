#pragma once

#include <windows.h>
#include <CommCtrl.h>

#include <unordered_map>

#include "OptionsStore.h"

namespace shelltabs {

enum class ExplorerSurfaceKind {
    ListView,
    Header,
    Rebar,
    Toolbar,
    Edit,
    DirectUi,
};

class ExplorerGlowRenderer {
public:
    ExplorerGlowRenderer();

    void Configure(const ShellTabsOptions& options);
    void RegisterSurface(HWND hwnd, ExplorerSurfaceKind kind);
    void UnregisterSurface(HWND hwnd);
    void Reset();
    void InvalidateRegisteredSurfaces() const;

    void HandleThemeChanged(HWND hwnd);
    void HandleSettingChanged(HWND hwnd);
    void HandleDpiChanged(HWND hwnd, UINT dpiX, UINT dpiY);

    bool ShouldRender() const noexcept { return m_glowEnabled; }

    void PaintSurface(HWND hwnd, ExplorerSurfaceKind kind, HDC targetDc, const RECT& clipRect);

private:
    struct HandleHasher {
        size_t operator()(HWND hwnd) const noexcept { return reinterpret_cast<size_t>(hwnd); }
    };

    struct SurfaceState {
        ExplorerSurfaceKind kind = ExplorerSurfaceKind::ListView;
        UINT dpiX = 96;
        UINT dpiY = 96;
    };

    struct GlowColorSet {
        bool valid = false;
        bool gradient = false;
        COLORREF start = RGB(0, 0, 0);
        COLORREF end = RGB(0, 0, 0);
    };

    void EnsureSurfaceState(HWND hwnd, ExplorerSurfaceKind kind);
    GlowColorSet ResolveColors(ExplorerSurfaceKind kind) const;
    void UpdateAccentColor();

    void PaintListView(HWND hwnd, const SurfaceState& state, HDC targetDc, const RECT& clipRect,
                       const GlowColorSet& colors);
    void PaintHeader(HWND hwnd, const SurfaceState& state, HDC targetDc, const RECT& clipRect,
                     const GlowColorSet& colors);
    void PaintRebar(HWND hwnd, const SurfaceState& state, HDC targetDc, const RECT& clipRect,
                    const GlowColorSet& colors);
    void PaintToolbar(HWND hwnd, const SurfaceState& state, HDC targetDc, const RECT& clipRect,
                      const GlowColorSet& colors);
    void PaintEdit(HWND hwnd, const SurfaceState& state, HDC targetDc, const RECT& clipRect,
                   const GlowColorSet& colors);
    void PaintDirectUi(HWND hwnd, const SurfaceState& state, HDC targetDc, const RECT& clipRect,
                       const GlowColorSet& colors);

    std::unordered_map<HWND, SurfaceState, HandleHasher> m_surfaces;
    GlowSurfacePalette m_palette{};
    bool m_glowEnabled = false;
    COLORREF m_accentColor = RGB(0, 120, 215);
};

}  // namespace shelltabs

