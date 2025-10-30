#include "StringUtils.h"

#include "OptionsStore.h"

#include <cwchar>

namespace shelltabs {

std::wstring Trim(const std::wstring& value) {
    const size_t begin = value.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos) {
        return {};
    }
    const size_t end = value.find_last_not_of(L" \t\r\n");
    return value.substr(begin, end - begin + 1);
}

std::vector<std::wstring> Split(const std::wstring& value, wchar_t delimiter) {
    std::vector<std::wstring> parts;
    size_t start = 0;
    while (start <= value.size()) {
        const size_t pos = value.find(delimiter, start);
        if (pos == std::wstring::npos) {
            parts.emplace_back(value.substr(start));
            break;
        }
        parts.emplace_back(value.substr(start, pos - start));
        start = pos + 1;
    }
    return parts;
}

bool ParseBool(const std::wstring& token) {
    if (token.empty()) {
        return false;
    }

    if (token == L"1") {
        return true;
    }
    if (token == L"0") {
        return false;
    }

    if (_wcsicmp(token.c_str(), L"true") == 0 || _wcsicmp(token.c_str(), L"yes") == 0 ||
        _wcsicmp(token.c_str(), L"on") == 0) {
        return true;
    }

    if (_wcsicmp(token.c_str(), L"false") == 0 || _wcsicmp(token.c_str(), L"no") == 0 ||
        _wcsicmp(token.c_str(), L"off") == 0) {
        return false;
    }

    return false;
}

TabBandDockMode ParseDockMode(const std::wstring& token) {
    if (token.empty()) {
        return TabBandDockMode::kAutomatic;
    }

    if (_wcsicmp(token.c_str(), L"top") == 0) {
        return TabBandDockMode::kTop;
    }
    if (_wcsicmp(token.c_str(), L"bottom") == 0) {
        return TabBandDockMode::kBottom;
    }
    if (_wcsicmp(token.c_str(), L"left") == 0) {
        return TabBandDockMode::kLeft;
    }
    if (_wcsicmp(token.c_str(), L"right") == 0) {
        return TabBandDockMode::kRight;
    }

    return TabBandDockMode::kAutomatic;
}

std::wstring DockModeToString(TabBandDockMode mode) {
    switch (mode) {
        case TabBandDockMode::kTop:
            return L"top";
        case TabBandDockMode::kBottom:
            return L"bottom";
        case TabBandDockMode::kLeft:
            return L"left";
        case TabBandDockMode::kRight:
            return L"right";
        default:
            return L"auto";
    }
}

}  // namespace shelltabs
