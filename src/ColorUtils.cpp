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

}  // namespace shelltabs

