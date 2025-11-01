#include "BreadcrumbGradient.h"

#include <algorithm>
#include <array>
#include <cmath>

namespace shelltabs {
namespace {

constexpr std::array<COLORREF, 7> kRainbowColors = {
    RGB(255, 59, 48),   // red
    RGB(255, 149, 0),   // orange
    RGB(255, 204, 0),   // yellow
    RGB(52, 199, 89),   // green
    RGB(0, 122, 255),   // blue
    RGB(88, 86, 214),   // indigo
    RGB(175, 82, 222)   // violet
};

BYTE BrightenChannel(BYTE channel, int delta) noexcept {
    const int value = static_cast<int>(channel) + delta;
    return static_cast<BYTE>(std::clamp(value, 0, 255));
}

}  // namespace

BreadcrumbGradientPalette ResolveBreadcrumbGradientPalette(const BreadcrumbGradientConfig& config) noexcept {
    BreadcrumbGradientPalette palette{};
    palette.brightness = std::clamp(config.brightness, 0, 100);

    COLORREF start = config.fontGradientStartColor;
    COLORREF end = config.fontGradientEndColor;

    if (!config.useCustomFontColors) {
        if (config.useCustomGradientColors) {
            start = config.gradientStartColor;
            end = config.gradientEndColor;
        } else {
            start = kRainbowColors.front();
            end = kRainbowColors[1];
        }
    }

    if (start == end) {
        COLORREF adjusted = RGB(BrightenChannel(GetRValue(end), 20), BrightenChannel(GetGValue(end), 20),
                                BrightenChannel(GetBValue(end), 20));
        if (adjusted == end) {
            adjusted = RGB(BrightenChannel(GetRValue(end), 10), BrightenChannel(GetGValue(end), 10),
                           BrightenChannel(GetBValue(end), 10));
        }
        end = adjusted;
    }

    palette.start = start;
    palette.end = end;
    return palette;
}

COLORREF EvaluateBreadcrumbGradientColor(const BreadcrumbGradientPalette& palette, double position) noexcept {
    const double clamped = std::clamp(position, 0.0, 1.0);
    const int brightness = std::clamp(palette.brightness, 0, 100);

    auto interpolate = [&](BYTE start, BYTE end) noexcept -> BYTE {
        const double interpolated = static_cast<double>(start) +
                                    (static_cast<double>(end) - static_cast<double>(start)) * clamped;
        int value = static_cast<int>(std::lround(interpolated));
        value = std::clamp(value, 0, 255);
        value = value + ((255 - value) * brightness) / 100;
        return static_cast<BYTE>(std::clamp(value, 0, 255));
    };

    const BYTE red = interpolate(GetRValue(palette.start), GetRValue(palette.end));
    const BYTE green = interpolate(GetGValue(palette.start), GetGValue(palette.end));
    const BYTE blue = interpolate(GetBValue(palette.start), GetBValue(palette.end));
    return RGB(red, green, blue);
}

}  // namespace shelltabs

