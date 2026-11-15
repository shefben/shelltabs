#include "ExplorerThemeUtils.h"

#include <algorithm>
#include <array>
#include <optional>
#include <utility>
#include <winreg.h>
#include <uxtheme.h>
#include <vssym32.h>

namespace shelltabs {

namespace {

constexpr const wchar_t kThemePreferenceKey[] =
    L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";
constexpr const wchar_t kThemePreferenceValue[] = L"AppsUseLightTheme";



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


std::optional<COLORREF> QueryEditThemeTextColor(HWND editWindow) {
    if (!editWindow || !IsWindow(editWindow)) {
        return std::nullopt;
    }

    auto openTheme = [&](HWND target) -> HTHEME {
        if (!target || !IsWindow(target)) {
            return nullptr;
        }
        return OpenThemeData(target, L"Edit");
    };

    HTHEME theme = openTheme(editWindow);
    if (!theme) {
        HWND parent = GetParent(editWindow);
        while (parent && IsWindow(parent) && !theme) {
            theme = openTheme(parent);
            if (!theme) {
                parent = GetParent(parent);
            }
        }
    }

    if (!theme) {
        return std::nullopt;
    }

    const int candidateStates[] = {ETS_NORMAL, ETS_HOT, ETS_SELECTED, ETS_FOCUSED, ETS_READONLY};
    COLORREF textColor = 0;
    HRESULT hr = E_FAIL;
    for (int state : candidateStates) {
        hr = GetThemeColor(theme, EP_EDITTEXT, state, TMT_TEXTCOLOR, &textColor);
        if (SUCCEEDED(hr)) {
            break;
        }
    }

    CloseThemeData(theme);
    if (FAILED(hr)) {
        return std::nullopt;
    }

    return textColor;
}

std::optional<COLORREF> QueryStatusBarThemeTextColor(HWND statusBar) {
    if (!statusBar || !IsWindow(statusBar)) {
        return std::nullopt;
    }

    HTHEME theme = OpenThemeData(statusBar, L"Status");
    if (!theme) {
        return std::nullopt;
    }

    COLORREF textColor = 0;
    HRESULT hr = GetThemeColor(theme, SP_PANE, 0, TMT_TEXTCOLOR, &textColor);
    if (FAILED(hr)) {
        hr = GetThemeColor(theme, SP_PANE, 1, TMT_TEXTCOLOR, &textColor);
    }
    if (FAILED(hr)) {
        hr = GetThemeColor(theme, SP_GRIPPERPANE, 0, TMT_TEXTCOLOR, &textColor);
    }

    CloseThemeData(theme);
    if (FAILED(hr)) {
        return std::nullopt;
    }

    return textColor;
}

}  // namespace shelltabs

