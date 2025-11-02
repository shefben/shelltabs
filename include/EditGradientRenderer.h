#pragma once

#include <optional>

#include <windows.h>

#include "BreadcrumbGradient.h"

namespace shelltabs {

struct GradientEditRenderOptions {
    bool hideCaret = true;
    bool requestEraseBackground = true;
    std::optional<RECT> clipRect;
};

// Renders edit control text using the configured breadcrumb gradient.
// Returns true if custom rendering occurred.
bool RenderGradientEditContent(HWND hwnd, HDC dc, const BreadcrumbGradientConfig& gradientConfig,
                               const GradientEditRenderOptions& options = {});

}  // namespace shelltabs

