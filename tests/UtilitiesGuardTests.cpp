#include "Utilities.h"

#include <windows.h>

#include <functional>
#include <iostream>
#include <string>
#include <stdexcept>

namespace {

int g_wideExceptionCount = 0;
int g_narrowExceptionCount = 0;
std::wstring g_lastWideContext;
std::wstring g_lastWideDetails;
std::wstring g_lastNarrowContext;
std::string g_lastNarrowDetails;

void ResetExceptionState() {
    g_wideExceptionCount = 0;
    g_narrowExceptionCount = 0;
    g_lastWideContext.clear();
    g_lastWideDetails.clear();
    g_lastNarrowContext.clear();
    g_lastNarrowDetails.clear();
}

void PrintFailure(const wchar_t* testName, const std::wstring& details) {
    std::wcerr << L"[" << testName << L"] " << details << std::endl;
}

bool TestGuardExplorerCallReturnsValue() {
    ResetExceptionState();
    const auto result = shelltabs::GuardExplorerCall(L"GuardSuccess", []() -> int { return 42; }, []() -> int {
        PrintFailure(L"TestGuardExplorerCallReturnsValue", L"Fallback should not execute");
        return -1;
    });

    if (result != 42) {
        PrintFailure(L"TestGuardExplorerCallReturnsValue", L"Unexpected result returned");
        return false;
    }
    if (g_wideExceptionCount != 0 || g_narrowExceptionCount != 0) {
        PrintFailure(L"TestGuardExplorerCallReturnsValue", L"Exception handlers were triggered unexpectedly");
        return false;
    }
    return true;
}

bool TestGuardExplorerCallHandlesStdException() {
    ResetExceptionState();
    bool fallbackInvoked = false;

    const auto result = shelltabs::GuardExplorerCall(
        L"GuardStdException",
        []() -> int {
            throw std::runtime_error("simulated failure");
        },
        [&]() -> int {
            fallbackInvoked = true;
            return 99;
        });

    if (!fallbackInvoked || result != 99) {
        PrintFailure(L"TestGuardExplorerCallHandlesStdException", L"Fallback path not executed correctly");
        return false;
    }
    if (g_narrowExceptionCount != 1 || g_lastNarrowContext != L"GuardStdException") {
        PrintFailure(L"TestGuardExplorerCallHandlesStdException", L"Narrow exception not logged");
        return false;
    }
    return true;
}

bool TestGuardExplorerCallHandlesUnknownException() {
    ResetExceptionState();
    bool fallbackInvoked = false;

    const auto result = shelltabs::GuardExplorerCall(
        L"GuardUnknownException",
        []() -> int {
            throw 7;
            return 0;
        },
        [&]() -> int {
            fallbackInvoked = true;
            return -7;
        });

    if (!fallbackInvoked || result != -7) {
        PrintFailure(L"TestGuardExplorerCallHandlesUnknownException", L"Fallback path not executed correctly");
        return false;
    }
    if (g_wideExceptionCount != 1 || g_lastWideContext != L"GuardUnknownException") {
        PrintFailure(L"TestGuardExplorerCallHandlesUnknownException", L"Wide exception not logged");
        return false;
    }
    return true;
}

}  // namespace

namespace shelltabs {

void LogUnhandledException(const wchar_t* context, const wchar_t* details) {
    ++g_wideExceptionCount;
    g_lastWideContext = context ? context : L"";
    g_lastWideDetails = details ? details : L"";
}

void LogUnhandledExceptionNarrow(const wchar_t* context, const char* details) {
    ++g_narrowExceptionCount;
    g_lastNarrowContext = context ? context : L"";
    g_lastNarrowDetails = details ? details : "";
}

}  // namespace shelltabs

int wmain() {
    const struct TestCase {
        const wchar_t* name;
        bool (*fn)();
    } tests[] = {
        {L"TestGuardExplorerCallReturnsValue", &TestGuardExplorerCallReturnsValue},
        {L"TestGuardExplorerCallHandlesStdException", &TestGuardExplorerCallHandlesStdException},
        {L"TestGuardExplorerCallHandlesUnknownException", &TestGuardExplorerCallHandlesUnknownException},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
