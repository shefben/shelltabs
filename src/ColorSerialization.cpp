#include "ColorSerialization.h"

#include <cwchar>
#include <sstream>

namespace shelltabs {

COLORREF ParseColor(const std::wstring& token, COLORREF fallback) {
    if (token.empty()) {
        return fallback;
    }

    unsigned int value = 0;
    std::wistringstream stream(token);
    if (token.size() > 2 && token[0] == L'0' && (token[1] == L'x' || token[1] == L'X')) {
        stream.ignore(2);
        stream >> std::hex >> value;
    } else if (!token.empty() && token[0] == L'#') {
        stream.ignore(1);
        stream >> std::hex >> value;
    } else {
        stream >> std::hex >> value;
    }

    return RGB((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
}

std::wstring ColorToString(COLORREF color) {
    wchar_t buffer[16];
    swprintf_s(buffer, L"%06X", (GetRValue(color) << 16) | (GetGValue(color) << 8) | GetBValue(color));
    return buffer;
}

TabGroupOutlineStyle ParseOutlineStyle(const std::wstring& token, TabGroupOutlineStyle fallback) {
    if (token.empty()) {
        return fallback;
    }

    if (_wcsicmp(token.c_str(), L"solid") == 0) {
        return TabGroupOutlineStyle::kSolid;
    }
    if (_wcsicmp(token.c_str(), L"dashed") == 0) {
        return TabGroupOutlineStyle::kDashed;
    }
    if (_wcsicmp(token.c_str(), L"dotted") == 0) {
        return TabGroupOutlineStyle::kDotted;
    }

    const int value = _wtoi(token.c_str());
    switch (value) {
        case 0:
            return TabGroupOutlineStyle::kSolid;
        case 1:
            return TabGroupOutlineStyle::kDashed;
        case 2:
            return TabGroupOutlineStyle::kDotted;
        default:
            break;
    }

    return fallback;
}

std::wstring OutlineStyleToString(TabGroupOutlineStyle style) {
    switch (style) {
        case TabGroupOutlineStyle::kDashed:
            return L"dashed";
        case TabGroupOutlineStyle::kDotted:
            return L"dotted";
        case TabGroupOutlineStyle::kSolid:
        default:
            return L"solid";
    }
}

}  // namespace shelltabs

