#pragma once

#include <windows.h>

namespace shelltabs {

struct BreadcrumbGradientConfig {
    bool enabled = false;
    int brightness = 85;
    bool useCustomFontColors = false;
    bool useCustomGradientColors = false;
    COLORREF fontGradientStartColor = RGB(255, 255, 255);
    COLORREF fontGradientEndColor = RGB(255, 255, 255);
    COLORREF gradientStartColor = RGB(255, 59, 48);
    COLORREF gradientEndColor = RGB(175, 82, 222);
};

struct BreadcrumbGradientPalette {
    COLORREF start = RGB(255, 255, 255);
    COLORREF end = RGB(255, 255, 255);
    int brightness = 85;
};

BreadcrumbGradientPalette ResolveBreadcrumbGradientPalette(const BreadcrumbGradientConfig& config) noexcept;
COLORREF EvaluateBreadcrumbGradientColor(const BreadcrumbGradientPalette& palette, double position) noexcept;

}  // namespace shelltabs

