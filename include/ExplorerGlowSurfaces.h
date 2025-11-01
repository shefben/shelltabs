#pragma once

#include <windows.h>
#include <CommCtrl.h>

#include <memory>
#include <optional>

#include "OptionsStore.h"

namespace shelltabs {

enum class ExplorerSurfaceKind {
    ListView,
    Header,
    Rebar,
    Toolbar,
    Edit,
    Scrollbar,
    DirectUi,
};

class ExplorerGlowSurface;

struct GlowColorSet {
    bool valid = false;
    bool gradient = false;
    COLORREF start = RGB(0, 0, 0);
    COLORREF end = RGB(0, 0, 0);
};

class ExplorerGlowCoordinator {
public:
    ExplorerGlowCoordinator();

    void Configure(const ShellTabsOptions& options);
    bool HandleThemeChanged();
    bool HandleSettingChanged();

    bool ShouldRender() const noexcept { return m_glowEnabled && !m_highContrastActive; }
    bool ShouldRenderSurface(ExplorerSurfaceKind kind) const noexcept;
    GlowColorSet ResolveColors(ExplorerSurfaceKind kind) const;

private:
    const GlowSurfaceOptions* ResolveSurfaceOptions(ExplorerSurfaceKind kind) const noexcept;
    void UpdateAccentColor();
    bool RefreshAccessibilityState();

    GlowSurfacePalette m_palette{};
    bool m_glowEnabled = false;
    bool m_highContrastActive = false;
    COLORREF m_accentColor = RGB(0, 120, 215);
};

class ExplorerGlowSurface {
public:
    ExplorerGlowSurface(ExplorerSurfaceKind kind, ExplorerGlowCoordinator& coordinator);
    virtual ~ExplorerGlowSurface();

    ExplorerGlowSurface(const ExplorerGlowSurface&) = delete;
    ExplorerGlowSurface& operator=(const ExplorerGlowSurface&) = delete;

    ExplorerGlowSurface(ExplorerGlowSurface&&) = delete;
    ExplorerGlowSurface& operator=(ExplorerGlowSurface&&) = delete;

    ExplorerSurfaceKind Kind() const noexcept { return m_kind; }
    HWND Handle() const noexcept { return m_hwnd; }

    bool Attach(HWND hwnd);
    void Detach();
    bool IsAttached() const noexcept;

    void RequestRepaint() const;

    bool SupportsImmediatePainting() const noexcept;
    bool PaintImmediately(HDC targetDc, const RECT& clipRect);

    virtual bool HandleNotify(const NMHDR& header, LRESULT* result);

protected:
    ExplorerGlowCoordinator& Coordinator() noexcept { return m_coordinator; }
    const ExplorerGlowCoordinator& Coordinator() const noexcept { return m_coordinator; }

    UINT DpiX() const noexcept { return m_dpiX; }
    UINT DpiY() const noexcept { return m_dpiY; }

    virtual void OnAttached();
    virtual void OnDetached();
    virtual void OnDpiChanged(UINT dpiX, UINT dpiY);
    virtual void OnThemeChanged();
    virtual void OnSettingsChanged();
    virtual bool UsesCustomDraw() const noexcept;
    virtual void OnPaint(HDC targetDc, const RECT& clipRect, const GlowColorSet& colors) = 0;
    void PaintInternal(HDC targetDc, const RECT& clipRect);

private:
    static LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                         UINT_PTR subclassId, DWORD_PTR refData);

    std::optional<LRESULT> HandleMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    LRESULT HandlePaintMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    LRESULT HandlePrintClient(HWND hwnd, WPARAM wParam, LPARAM lParam);

    ExplorerSurfaceKind m_kind;
    ExplorerGlowCoordinator& m_coordinator;
    HWND m_hwnd = nullptr;
    bool m_subclassInstalled = false;
    UINT m_dpiX = 96;
    UINT m_dpiY = 96;
};

std::unique_ptr<ExplorerGlowSurface> CreateGlowSurfaceWrapper(ExplorerSurfaceKind kind,
                                                              ExplorerGlowCoordinator& coordinator);

}  // namespace shelltabs

