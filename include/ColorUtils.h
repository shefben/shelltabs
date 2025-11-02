#pragma once

#include <windows.h>

namespace shelltabs {

double ComputeColorLuminance(COLORREF color) noexcept;
double ComputeContrastRatio(double a, double b) noexcept;

}  // namespace shelltabs

