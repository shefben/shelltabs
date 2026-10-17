#include "ColorUtils.h"

#include <algorithm>
#include <cmath>

namespace shelltabs {
namespace {

double SrgbChannelToLinear(BYTE channel) noexcept {
    const double srgb = static_cast<double>(channel) / 255.0;
    if (srgb <= 0.04045) {
        return srgb / 12.92;
    }
    return std::pow((srgb + 0.055) / 1.055, 2.4);
}

}  // namespace

double ComputeColorLuminance(COLORREF color) noexcept {
    const double r = SrgbChannelToLinear(GetRValue(color));
    const double g = SrgbChannelToLinear(GetGValue(color));
    const double b = SrgbChannelToLinear(GetBValue(color));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

double ComputeContrastRatio(double a, double b) noexcept {
    const double lighter = std::max(a, b);
    const double darker = std::min(a, b);
    return (lighter + 0.05) / (darker + 0.05);
}

COLORREF ResolveTabRecencyColor(COLORREF background, COLORREF accent, double opacity) noexcept {
    if (opacity <= 0.0) {
        return background;
    }
    const auto blend = [](COLORREF base, COLORREF color, double amount) {
        const auto channel = [=](BYTE a, BYTE b) {
            return static_cast<BYTE>(std::lround(a * (1.0 - amount) + b * amount));
        };
        return RGB(channel(GetRValue(base), GetRValue(color)),
                   channel(GetGValue(base), GetGValue(color)),
                   channel(GetBValue(base), GetBValue(color)));
    };
    const double backgroundLuminance = ComputeColorLuminance(background);
    const auto contrast = [=](COLORREF color) {
        return ComputeContrastRatio(backgroundLuminance, ComputeColorLuminance(color));
    };
    const auto reachContrast = [&](COLORREF from, COLORREF to, double target) {
        double low = 0.0;
        double high = 1.0;
        for (int iteration = 0; iteration < 16; ++iteration) {
            const double middle = (low + high) / 2.0;
            if (contrast(blend(from, to, middle)) >= target) {
                high = middle;
            } else {
                low = middle;
            }
        }
        return blend(from, to, high);
    };
    // Give the full-strength dot enough contrast for all five opacity steps.
    // Mid-tone backgrounds may require pure black or white to achieve this.
    const double whiteContrast = contrast(RGB(255, 255, 255));
    const double blackContrast = contrast(RGB(0, 0, 0));
    const double fullContrast = std::min(7.0, std::max(whiteContrast, blackContrast));
    if (contrast(accent) < fullContrast) {
        const COLORREF endpoint = whiteContrast > blackContrast ? RGB(255, 255, 255) : RGB(0, 0, 0);
        accent = reachContrast(accent, endpoint, fullContrast);
    }
    // Preserve actual opacity differences instead of squeezing every rank into
    // a narrow high-contrast range, which makes small dots look alike.
    return blend(background, accent, std::clamp(opacity, 0.0, 1.0));
}

}  // namespace shelltabs

