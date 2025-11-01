#pragma once

#include <windows.h>

#include <optional>

namespace shelltabs {

std::optional<COLORREF> SampleAverageColor(HDC dc, const RECT& rect);

bool IsSystemHighContrastActive();

}

