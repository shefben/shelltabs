#pragma once

#include <windows.h>

#include <optional>

namespace shelltabs {

struct ToolbarChromeSample {
    COLORREF topColor = 0;
    COLORREF bottomColor = 0;
};

std::optional<COLORREF> SampleAverageColor(HDC dc, const RECT& rect);

bool IsSystemHighContrastActive();

std::optional<bool> QueryAppsUseLightThemePreference();

bool IsAppDarkModePreferred();

std::optional<ToolbarChromeSample> SampleToolbarChrome(HWND window);

std::optional<COLORREF> QueryEditThemeTextColor(HWND editWindow);

std::optional<COLORREF> QueryStatusBarThemeTextColor(HWND statusBar);

}

