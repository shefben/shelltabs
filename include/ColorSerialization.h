#pragma once

#include <windows.h>

#include <string>

#include "TabManager.h"

namespace shelltabs {

COLORREF ParseColor(const std::wstring& token, COLORREF fallback);
std::wstring ColorToString(COLORREF color);
TabGroupOutlineStyle ParseOutlineStyle(const std::wstring& token, TabGroupOutlineStyle fallback);
std::wstring OutlineStyleToString(TabGroupOutlineStyle style);

}  // namespace shelltabs

