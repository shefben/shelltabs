#include "ExplorerThemeUtils.h"

#include <algorithm>
#include <array>
#include <optional>
#include <utility>
#include <winreg.h>

namespace shelltabs {

namespace {

constexpr const wchar_t kThemePreferenceKey[] =
    L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
constexpr const wchar_t kThemePreferenceValue[] = L"AppsUseLightTheme";

RECT ShrinkRectIfPossible(const RECT& rect, LONG insetX, LONG insetY) {
    RECT result = rect;
    const LONG width = rect.right - rect.left;
    const LONG height = rect.bottom - rect.top;
    if (width > insetX * 2) {
        result.left += insetX;
        result.right -= insetX;
    }
    if (height > insetY * 2) {
        result.top += insetY;
        result.bottom -= insetY;
    }
    if (result.right <= result.left || result.bottom <= result.top) {
        return rect;
    }
    return result;
}

std::optional<COLORREF> SampleWindowRectAverage(HWND window, const RECT& rect) {
    if (!window || !IsWindow(window)) {
        return std::nullopt;
    }
    HDC dc = GetDC(window);
    if (!dc) {
        return std::nullopt;
    }
    std::optional<COLORREF> color = SampleAverageColor(dc, rect);
    ReleaseDC(window, dc);
    return color;
}

}  // namespace

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

bool IsSystemHighContrastActive() {
    HIGHCONTRASTW info{sizeof(info)};
    if (!SystemParametersInfoW(SPI_GETHIGHCONTRAST, sizeof(info), &info, FALSE)) {
        return false;
    }
    return (info.dwFlags & HCF_HIGHCONTRASTON) != 0;
}

std::optional<bool> QueryAppsUseLightThemePreference() {
    DWORD value = 1;
    DWORD size = sizeof(value);
    const LSTATUS status =
        RegGetValueW(HKEY_CURRENT_USER, kThemePreferenceKey, kThemePreferenceValue, RRF_RT_DWORD, nullptr, &value, &size);
    if (status != ERROR_SUCCESS) {
        return std::nullopt;
    }
    return value != 0;
}

bool IsAppDarkModePreferred() {
    const auto lightPreference = QueryAppsUseLightThemePreference();
    if (!lightPreference.has_value()) {
        return false;
    }
    return !lightPreference.value();
}

std::optional<ToolbarChromeSample> SampleToolbarChrome(HWND window) {
    if (!window || !IsWindow(window)) {
        return std::nullopt;
    }

    RECT client{};
    if (!GetClientRect(window, &client) || client.right <= client.left || client.bottom <= client.top) {
        return std::nullopt;
    }

    const LONG height = client.bottom - client.top;
    const LONG sampleHeight = std::max<LONG>(1, height / 4);

    RECT topRect = client;
    topRect.bottom = topRect.top + sampleHeight;
    topRect = ShrinkRectIfPossible(topRect, 4, 2);

    RECT bottomRect = client;
    bottomRect.top = std::max(bottomRect.bottom - sampleHeight, bottomRect.top);
    bottomRect = ShrinkRectIfPossible(bottomRect, 4, 2);

    const auto topSample = SampleWindowRectAverage(window, topRect);
    const auto bottomSample = SampleWindowRectAverage(window, bottomRect);

    if (!topSample.has_value() && !bottomSample.has_value()) {
        return std::nullopt;
    }

    ToolbarChromeSample sample{};
    sample.topColor = topSample.value_or(bottomSample.value_or(0));
    sample.bottomColor = bottomSample.value_or(sample.topColor);
    return sample;
}

}  // namespace shelltabs

