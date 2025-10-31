#include "Utilities.h"

#include <windows.h>

#include <iostream>
#include <stdexcept>
#include <string>
#include <vector>

namespace {

struct GuardLogCapture {
    void Reset() {
        wideCount = 0;
        narrowCount = 0;
        context.clear();
        details.clear();
        narrowContext.clear();
        narrowDetails.clear();
    }

    int wideCount = 0;
    int narrowCount = 0;
    std::wstring context;
    std::wstring details;
    std::wstring narrowContext;
    std::string narrowDetails;
} g_logCapture;

void ResetLogs() {
    g_logCapture.Reset();
}

struct TestDefinition {
    const wchar_t* name;
    bool (*fn)();
};

void PrintFailure(const wchar_t* testName, const std::wstring& message) {
    std::wcerr << L"[" << testName << L"] " << message << std::endl;
}

}  // namespace

namespace shelltabs {

void LogUnhandledException(const wchar_t* context, const wchar_t* details) {
    ++g_logCapture.wideCount;
    g_logCapture.context = context ? context : L"";
    g_logCapture.details = details ? details : L"";
}

void LogUnhandledExceptionNarrow(const wchar_t* context, const char* details) {
    ++g_logCapture.narrowCount;
    g_logCapture.narrowContext = context ? context : L"";
    g_logCapture.narrowDetails = details ? details : "";
}

}  // namespace shelltabs

namespace {

bool TestGuardExplorerCallSuccess() {
    ResetLogs();
    bool executed = false;
    bool fallbackExecuted = false;

    const HRESULT hr = shelltabs::GuardExplorerCall(
        L"Success",
        [&]() -> HRESULT {
            executed = true;
            return S_OK;
        },
        [&]() -> HRESULT {
            fallbackExecuted = true;
            return E_FAIL;
        });

    if (!executed) {
        PrintFailure(L"TestGuardExplorerCallSuccess", L"Primary callable was not executed");
        return false;
    }

    if (fallbackExecuted) {
        PrintFailure(L"TestGuardExplorerCallSuccess", L"Fallback executed unexpectedly");
        return false;
    }

    if (hr != S_OK) {
        PrintFailure(L"TestGuardExplorerCallSuccess", L"Guard returned unexpected HRESULT");
        return false;
    }

    if (g_logCapture.wideCount != 0 || g_logCapture.narrowCount != 0) {
        PrintFailure(L"TestGuardExplorerCallSuccess", L"Exceptions were logged unexpectedly");
        return false;
    }

    return true;
}

bool TestGuardExplorerCallStdException() {
    ResetLogs();
    bool fallbackExecuted = false;

    const HRESULT hr = shelltabs::GuardExplorerCall(
        L"StdException",
        [&]() -> HRESULT {
            throw std::runtime_error("boom");
        },
        [&]() -> HRESULT {
            fallbackExecuted = true;
            return E_FAIL;
        });

    if (!fallbackExecuted) {
        PrintFailure(L"TestGuardExplorerCallStdException", L"Fallback was not executed after exception");
        return false;
    }

    if (hr != E_FAIL) {
        PrintFailure(L"TestGuardExplorerCallStdException", L"Unexpected HRESULT returned after fallback");
        return false;
    }

    if (g_logCapture.narrowCount != 1 || g_logCapture.wideCount != 0) {
        PrintFailure(L"TestGuardExplorerCallStdException", L"Expected narrow log entry was not recorded");
        return false;
    }

    if (g_logCapture.narrowContext != L"StdException") {
        PrintFailure(L"TestGuardExplorerCallStdException", L"Context was not propagated to logger");
        return false;
    }

    if (g_logCapture.narrowDetails.find("boom") == std::string::npos) {
        PrintFailure(L"TestGuardExplorerCallStdException", L"Exception details were not captured");
        return false;
    }

    return true;
}

bool TestGuardExplorerCallUnknownException() {
    ResetLogs();
    bool fallbackExecuted = false;

    const HRESULT fallbackResult = HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
    const HRESULT hr = shelltabs::GuardExplorerCall(
        L"UnknownException",
        [&]() -> HRESULT {
            throw 42;
        },
        [&]() -> HRESULT {
            fallbackExecuted = true;
            return fallbackResult;
        });

    if (!fallbackExecuted) {
        PrintFailure(L"TestGuardExplorerCallUnknownException", L"Fallback was not executed for unknown exception");
        return false;
    }

    if (hr != fallbackResult) {
        PrintFailure(L"TestGuardExplorerCallUnknownException", L"Unexpected HRESULT returned for unknown exception");
        return false;
    }

    if (g_logCapture.wideCount != 1 || g_logCapture.narrowCount != 0) {
        PrintFailure(L"TestGuardExplorerCallUnknownException", L"Expected wide log entry was not recorded");
        return false;
    }

    if (g_logCapture.context != L"UnknownException") {
        PrintFailure(L"TestGuardExplorerCallUnknownException", L"Context was not forwarded to wide logger");
        return false;
    }

    return true;
}

bool TestGuardExplorerCallVoidException() {
    ResetLogs();
    bool executed = false;

    shelltabs::GuardExplorerCall(
        L"VoidException",
        [&]() {
            executed = true;
            throw std::runtime_error("void failure");
        });

    if (!executed) {
        PrintFailure(L"TestGuardExplorerCallVoidException", L"Primary callable was not executed");
        return false;
    }

    if (g_logCapture.narrowCount != 1 || g_logCapture.wideCount != 0) {
        PrintFailure(L"TestGuardExplorerCallVoidException", L"Expected narrow log entry missing after void exception");
        return false;
    }

    if (g_logCapture.narrowContext != L"VoidException") {
        PrintFailure(L"TestGuardExplorerCallVoidException", L"Context was not passed to logger");
        return false;
    }

    if (g_logCapture.narrowDetails.find("void failure") == std::string::npos) {
        PrintFailure(L"TestGuardExplorerCallVoidException", L"Exception details not captured for void handler");
        return false;
    }

    return true;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestGuardExplorerCallSuccess", &TestGuardExplorerCallSuccess},
        {L"TestGuardExplorerCallStdException", &TestGuardExplorerCallStdException},
        {L"TestGuardExplorerCallUnknownException", &TestGuardExplorerCallUnknownException},
        {L"TestGuardExplorerCallVoidException", &TestGuardExplorerCallVoidException},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
