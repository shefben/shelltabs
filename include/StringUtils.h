#pragma once

#include <string>
#include <vector>

namespace shelltabs {

enum class TabBandDockMode;

std::wstring Trim(const std::wstring& value);
std::vector<std::wstring> Split(const std::wstring& value, wchar_t delimiter);
bool ParseBool(const std::wstring& token);
TabBandDockMode ParseDockMode(const std::wstring& token);
std::wstring DockModeToString(TabBandDockMode mode);

}  // namespace shelltabs
