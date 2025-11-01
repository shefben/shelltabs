#include "StringUtils.h"

#include "OptionsStore.h"

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <limits>

namespace shelltabs {

namespace {

constexpr std::wstring_view kWhitespace = L" \t\r\n";

}  // namespace

bool EqualsIgnoreCase(std::wstring_view lhs, std::wstring_view rhs) {
    if (lhs.size() != rhs.size()) {
        return false;
    }
    for (size_t i = 0; i < lhs.size(); ++i) {
        if (std::towlower(static_cast<wint_t>(lhs[i])) != std::towlower(static_cast<wint_t>(rhs[i]))) {
            return false;
        }
    }
    return true;
}

std::wstring_view TrimView(std::wstring_view value) {
    const size_t begin = value.find_first_not_of(kWhitespace);
    if (begin == std::wstring_view::npos) {
        return {};
    }
    const size_t end = value.find_last_not_of(kWhitespace);
    return value.substr(begin, end - begin + 1);
}

std::wstring Trim(std::wstring_view value) {
    return std::wstring(TrimView(value));
}

std::vector<std::wstring_view> Split(std::wstring_view value, wchar_t delimiter) {
    std::vector<std::wstring_view> parts;
    size_t start = 0;
    while (start <= value.size()) {
        const size_t pos = value.find(delimiter, start);
        if (pos == std::wstring_view::npos) {
            parts.emplace_back(value.substr(start));
            break;
        }
        parts.emplace_back(value.substr(start, pos - start));
        start = pos + 1;
    }
    return parts;
}

bool ParseBool(std::wstring_view token) {
    if (token.empty()) {
        return false;
    }

    if (token == L"1") {
        return true;
    }
    if (token == L"0") {
        return false;
    }

    if (EqualsIgnoreCase(token, L"true") || EqualsIgnoreCase(token, L"yes") ||
        EqualsIgnoreCase(token, L"on")) {
        return true;
    }

    if (EqualsIgnoreCase(token, L"false") || EqualsIgnoreCase(token, L"no") ||
        EqualsIgnoreCase(token, L"off")) {
        return false;
    }

    return false;
}

int ParseInt(std::wstring_view token) {
    if (token.empty()) {
        return 0;
    }

    bool negative = false;
    size_t index = 0;
    if (token[index] == L'+') {
        ++index;
    } else if (token[index] == L'-') {
        negative = true;
        ++index;
    }

    long long value = 0;
    constexpr long long kIntMax = static_cast<long long>(std::numeric_limits<int>::max());
    constexpr long long kIntMin = static_cast<long long>(std::numeric_limits<int>::min());
    for (; index < token.size(); ++index) {
        const wchar_t ch = token[index];
        if (ch < L'0' || ch > L'9') {
            break;
        }
        value = value * 10 + (ch - L'0');
        if (!negative && value > kIntMax) {
            return std::numeric_limits<int>::max();
        }
        if (negative && -value < kIntMin) {
            return std::numeric_limits<int>::min();
        }
    }

    if (negative) {
        value = -value;
    }
    if (value > kIntMax) {
        return std::numeric_limits<int>::max();
    }
    if (value < kIntMin) {
        return std::numeric_limits<int>::min();
    }
    return static_cast<int>(value);
}

bool TryParseUint64(std::wstring_view token, uint64_t* valueOut) {
    if (!valueOut || token.empty()) {
        return false;
    }

    uint64_t value = 0;
    for (wchar_t ch : token) {
        if (ch < L'0' || ch > L'9') {
            return false;
        }
        const uint64_t digit = static_cast<uint64_t>(ch - L'0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10) {
            return false;
        }
        value = value * 10 + digit;
    }

    *valueOut = value;
    return true;
}

TabBandDockMode ParseDockMode(std::wstring_view token) {
    if (token.empty()) {
        return TabBandDockMode::kAutomatic;
    }

    if (EqualsIgnoreCase(token, L"top")) {
        return TabBandDockMode::kTop;
    }
    if (EqualsIgnoreCase(token, L"bottom")) {
        return TabBandDockMode::kBottom;
    }
    if (EqualsIgnoreCase(token, L"left")) {
        return TabBandDockMode::kLeft;
    }
    if (EqualsIgnoreCase(token, L"right")) {
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

NewTabTemplate ParseNewTabTemplate(std::wstring_view token) {
    if (token.empty()) {
        return NewTabTemplate::kDuplicateCurrent;
    }

    if (EqualsIgnoreCase(token, L"this_pc") || EqualsIgnoreCase(token, L"thispc")) {
        return NewTabTemplate::kThisPc;
    }
    if (EqualsIgnoreCase(token, L"custom_path") || EqualsIgnoreCase(token, L"custom")) {
        return NewTabTemplate::kCustomPath;
    }
    if (EqualsIgnoreCase(token, L"saved_group") || EqualsIgnoreCase(token, L"group")) {
        return NewTabTemplate::kSavedGroup;
    }

    return NewTabTemplate::kDuplicateCurrent;
}

std::wstring NewTabTemplateToString(NewTabTemplate value) {
    switch (value) {
        case NewTabTemplate::kThisPc:
            return L"this_pc";
        case NewTabTemplate::kCustomPath:
            return L"custom_path";
        case NewTabTemplate::kSavedGroup:
            return L"saved_group";
        default:
            return L"duplicate_current";
    }
}

}  // namespace shelltabs
