#pragma once

#include <string>
#include <vector>

namespace shelltabs {

std::wstring Trim(const std::wstring& value);
std::vector<std::wstring> Split(const std::wstring& value, wchar_t delimiter);

}  // namespace shelltabs
