#include "TabManager.h"

#include <windows.h>

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

shelltabs::TabManager::ExplorerWindowId MakeId(uintptr_t hwndValue, uintptr_t cookie) {
    shelltabs::TabManager::ExplorerWindowId id;
    id.hwnd = reinterpret_cast<HWND>(hwndValue);
    id.frameCookie = cookie;
    return id;
}

bool TestRegistrationLifecycle() {
    shelltabs::TabManager manager;
    const auto id = MakeId(0x1001, 0xABCDEF01);
    manager.SetWindowId(id);
    if (shelltabs::TabManager::ActiveWindowCount() != 1) {
        PrintFailure(L"TestRegistrationLifecycle", L"Active window count mismatch after SetWindowId");
        return false;
    }
    manager.ClearWindowId();
    if (shelltabs::TabManager::ActiveWindowCount() != 0) {
        PrintFailure(L"TestRegistrationLifecycle", L"Active window count mismatch after ClearWindowId");
        return false;
    }
    return true;
}

bool TestDestructorClearsRegistration() {
    const auto id = MakeId(0x2002, 0x12345678);
    {
        shelltabs::TabManager manager;
        manager.SetWindowId(id);
        if (shelltabs::TabManager::ActiveWindowCount() != 1) {
            PrintFailure(L"TestDestructorClearsRegistration", L"Registration did not increase count");
            return false;
        }
    }
    if (shelltabs::TabManager::ActiveWindowCount() != 0) {
        PrintFailure(L"TestDestructorClearsRegistration", L"Destructor left registration behind");
        return false;
    }
    return true;
}

bool TestStressOpenCloseWindows() {
    constexpr int kIterations = 64;
    for (int i = 0; i < kIterations; ++i) {
        shelltabs::TabManager manager;
        const auto id = MakeId(0x3000 + static_cast<uintptr_t>(i), 0xCAFEB000 + static_cast<uintptr_t>(i));
        manager.SetWindowId(id);
        if (shelltabs::TabManager::ActiveWindowCount() != 1) {
            PrintFailure(L"TestStressOpenCloseWindows",
                         L"Registration count mismatch during iteration " + std::to_wstring(i));
            return false;
        }
        manager.Clear();
        manager.ClearWindowId();
        if (shelltabs::TabManager::ActiveWindowCount() != 0) {
            PrintFailure(L"TestStressOpenCloseWindows",
                         L"Registration persisted after ClearWindowId during iteration " + std::to_wstring(i));
            return false;
        }
    }
    return true;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestRegistrationLifecycle", &TestRegistrationLifecycle},
        {L"TestDestructorClearsRegistration", &TestDestructorClearsRegistration},
        {L"TestStressOpenCloseWindows", &TestStressOpenCloseWindows},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
