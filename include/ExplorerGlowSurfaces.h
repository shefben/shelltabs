#pragma once

#include <windows.h>
#include <CommCtrl.h>

#include <memory>
#include <mutex>
#include <optional>
#include <unordered_map>

#include "BreadcrumbGradient.h"
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
    PopupMenu,
    Tooltip,
};

class ExplorerGlowSurface;

struct GlowColorSet {
    bool valid = false;
    bool gradient = false;
    COLORREF start = RGB(0, 0, 0);
    COLORREF end = RGB(0, 0, 0);
};

enum class SurfacePaintRole {
    Generic,
    ListViewRows,
    StatusPane,
};

struct SurfaceColorDescriptor {
    ExplorerSurfaceKind kind = ExplorerSurfaceKind::ListView;
    SurfacePaintRole role = SurfacePaintRole::Generic;
    GlowColorSet fillColors{};
    bool fillOverride = false;
    COLORREF textColor = CLR_INVALID;
    bool textOverride = false;
    COLORREF backgroundColor = CLR_INVALID;
    bool backgroundOverride = false;
    bool forceOpaqueBackground = false;
    using BackgroundPaintCallback = bool (*)(HDC dc, HWND window, const RECT& rect, void* context);
    BackgroundPaintCallback backgroundPaintCallback = nullptr;
    void* backgroundPaintContext = nullptr;
    bool forcedHooks = false;
    bool userAccessibilityOptOut = false;
    bool accessibilityOptOut = false;
};

struct ScrollbarGlowDefinition {
    GlowColorSet colors{};
    BYTE trackLineAlpha = 0;
    BYTE trackHaloAlpha = 0;
    BYTE thumbFillAlpha = 0;
    BYTE thumbHaloAlpha = 0;
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
    std::optional<ScrollbarGlowDefinition> ResolveScrollbarDefinition() const;
    const BreadcrumbGradientConfig& BreadcrumbFontGradient() const noexcept { return m_breadcrumbFontGradient; }
    bool BitmapInterceptEnabled() const noexcept { return m_bitmapInterceptEnabled; }

    SurfaceColorDescriptor* AcquireSurfaceDescriptor(HWND hwnd, ExplorerSurfaceKind kind);
    SurfaceColorDescriptor* LookupSurfaceDescriptor(HWND hwnd) const;
    void ReleaseSurfaceDescriptor(HWND hwnd);
    void UpdateSurfaceDescriptor(HWND hwnd, const SurfaceColorDescriptor& descriptor);
    void SetSurfaceForcedHooks(HWND hwnd, bool forced);
    void SetSurfaceRole(HWND hwnd, SurfacePaintRole role);
    void SetSurfaceAccessibilityOptOut(HWND hwnd, bool optOut);

private:
    const GlowSurfaceOptions* ResolveSurfaceOptions(ExplorerSurfaceKind kind) const noexcept;
    void UpdateAccentColor();
    bool RefreshAccessibilityState();
    void RefreshDescriptorAccessibility();

    struct HandleHasher {
        size_t operator()(HWND hwnd) const noexcept { return reinterpret_cast<size_t>(hwnd); }
    };

    GlowSurfacePalette m_palette{};
    BreadcrumbGradientConfig m_breadcrumbFontGradient{};
    bool m_glowEnabled = false;
    bool m_highContrastActive = false;
    COLORREF m_accentColor = RGB(0, 120, 215);
    mutable std::mutex m_descriptorMutex;
    std::unordered_map<HWND, std::unique_ptr<SurfaceColorDescriptor>, HandleHasher> m_surfaceDescriptors;
    bool m_bitmapInterceptEnabled = true;
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

