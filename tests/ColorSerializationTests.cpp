#include "ColorSerialization.h"

#include <windows.h>

#include <array>
#include <iostream>
#include <string>
#include <vector>

namespace {

struct TestDefinition {
    const wchar_t* name;
    bool (*fn)();
};

void PrintFailure(const wchar_t* testName, const std::wstring& message) {
    std::wcerr << L"[" << testName << L"] " << message << std::endl;
}

bool TestParseHexColors() {
    const COLORREF fallback = RGB(0x12, 0x34, 0x56);
    struct Case {
        std::wstring token;
        COLORREF expected;
    };

    const std::vector<Case> cases = {
        {L"0xFFAABB", RGB(0xFF, 0xAA, 0xBB)},
        {L"0X112233", RGB(0x11, 0x22, 0x33)},
        {L"FFAABB", RGB(0xFF, 0xAA, 0xBB)},
        {L"ffaabb", RGB(0xFF, 0xAA, 0xBB)},
        {L"#010203", RGB(0x01, 0x02, 0x03)},
    };

    bool success = true;
    for (const auto& testCase : cases) {
        const COLORREF parsed = shelltabs::ParseColor(testCase.token, fallback);
        if (parsed != testCase.expected) {
            PrintFailure(L"TestParseHexColors", L"Failed to parse token: " + testCase.token);
            success = false;
        }
    }

    if (shelltabs::ParseColor(L"", fallback) != fallback) {
        PrintFailure(L"TestParseHexColors", L"Empty token did not return fallback");
        success = false;
    }

    return success;
}

bool TestOutlineStyleParsing() {
    struct Case {
        std::wstring token;
        shelltabs::TabGroupOutlineStyle expected;
    };

    const std::vector<Case> cases = {
        {L"solid", shelltabs::TabGroupOutlineStyle::kSolid},
        {L"SOLID", shelltabs::TabGroupOutlineStyle::kSolid},
        {L"dashed", shelltabs::TabGroupOutlineStyle::kDashed},
        {L"DOTTED", shelltabs::TabGroupOutlineStyle::kDotted},
        {L"0", shelltabs::TabGroupOutlineStyle::kSolid},
        {L"1", shelltabs::TabGroupOutlineStyle::kDashed},
        {L"2", shelltabs::TabGroupOutlineStyle::kDotted},
    };

    bool success = true;
    for (const auto& testCase : cases) {
        const auto parsed = shelltabs::ParseOutlineStyle(testCase.token, shelltabs::TabGroupOutlineStyle::kSolid);
        if (parsed != testCase.expected) {
            PrintFailure(L"TestOutlineStyleParsing", L"Failed to parse token: " + testCase.token);
            success = false;
        }
    }

    if (shelltabs::ParseOutlineStyle(L"", shelltabs::TabGroupOutlineStyle::kDotted) !=
        shelltabs::TabGroupOutlineStyle::kDotted) {
        PrintFailure(L"TestOutlineStyleParsing", L"Empty token did not return fallback");
        success = false;
    }

    return success;
}

bool TestRoundTripSerialization() {
    const std::array<COLORREF, 3> colors = {
        RGB(0x00, 0x00, 0x00),
        RGB(0x12, 0x34, 0x56),
        RGB(0xFF, 0xEE, 0xDD),
    };

    bool success = true;
    for (const COLORREF color : colors) {
        const std::wstring serialized = shelltabs::ColorToString(color);
        const COLORREF parsed = shelltabs::ParseColor(serialized, RGB(0, 0, 0));
        if (parsed != color) {
            PrintFailure(L"TestRoundTripSerialization", L"Color round-trip failed for " + serialized);
            success = false;
        }
    }

    const std::array<shelltabs::TabGroupOutlineStyle, 3> styles = {
        shelltabs::TabGroupOutlineStyle::kSolid,
        shelltabs::TabGroupOutlineStyle::kDashed,
        shelltabs::TabGroupOutlineStyle::kDotted,
    };

    for (const auto style : styles) {
        const std::wstring serialized = shelltabs::OutlineStyleToString(style);
        const auto parsed = shelltabs::ParseOutlineStyle(serialized, shelltabs::TabGroupOutlineStyle::kSolid);
        if (parsed != style) {
            PrintFailure(L"TestRoundTripSerialization", L"Outline style round-trip failed for " + serialized);
            success = false;
        }
    }

    return success;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestParseHexColors", &TestParseHexColors},
        {L"TestOutlineStyleParsing", &TestOutlineStyleParsing},
        {L"TestRoundTripSerialization", &TestRoundTripSerialization},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}

