#include "ExplorerThemeUtils.h"

#include <algorithm>
#include <array>
#include <optional>

namespace shelltabs {

std::optional<COLORREF> SampleAverageColor(HDC dc, const RECT& rect) {
    if (!dc || rect.left >= rect.right || rect.top >= rect.bottom) {
        return std::nullopt;
    }

    const LONG left = std::max(rect.left, static_cast<LONG>(0));
    const LONG top = std::max(rect.top, static_cast<LONG>(0));
    const LONG right = std::max(rect.right - 1, left);
    const LONG bottom = std::max(rect.bottom - 1, top);

    const std::array<POINT, 4> samplePoints = {{{left, top},
                                               {right, top},
                                               {left, bottom},
                                               {right, bottom}}};

    int totalRed = 0;
    int totalGreen = 0;
    int totalBlue = 0;
    int count = 0;

    for (const auto& point : samplePoints) {
        const COLORREF pixel = GetPixel(dc, point.x, point.y);
        if (pixel == CLR_INVALID) {
            continue;
        }
        totalRed += GetRValue(pixel);
        totalGreen += GetGValue(pixel);
        totalBlue += GetBValue(pixel);
        ++count;
    }

    if (count == 0) {
        return std::nullopt;
    }

    return RGB(totalRed / count, totalGreen / count, totalBlue / count);
}

}  // namespace shelltabs

