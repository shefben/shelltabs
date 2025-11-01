#include "IconCache.h"

#include <windows.h>
#include <shellapi.h>

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

HICON DuplicateStockIcon() {
    HICON base = LoadIconW(nullptr, IDI_APPLICATION);
    if (!base) {
        return nullptr;
    }
    return CopyIcon(base);
}

bool TestInvalidateFamilyScalesWithIndex() {
    auto& cache = shelltabs::IconCache::Instance();
    cache.DebugSetCapacity(4096);
    cache.DebugResetLastFamilyInvalidationCount();

    const std::wstring targetFamily = L"StressTargetFamily";
    auto loader = []() -> HICON { return DuplicateStockIcon(); };

    shelltabs::IconCache::Reference inUseSmall = cache.Acquire(targetFamily, SHGFI_SMALLICON, loader);
    if (!inUseSmall) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex", L"Failed to load initial small icon");
        return false;
    }
    const HICON initialHandle = inUseSmall.Get();

    shelltabs::IconCache::Reference targetLarge = cache.Acquire(targetFamily, SHGFI_LARGEICON, loader);
    if (!targetLarge) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex", L"Failed to load initial large icon");
        return false;
    }
    targetLarge.Reset();

    for (size_t i = 0; i < 3000; ++i) {
        const std::wstring family = L"StressFamily" + std::to_wstring(i);
        shelltabs::IconCache::Reference icon = cache.Acquire(family, SHGFI_SMALLICON, loader);
        if (!icon) {
            PrintFailure(L"TestInvalidateFamilyScalesWithIndex", L"Failed to populate cache with " + family);
            return false;
        }
        icon.Reset();
    }

    cache.DebugResetLastFamilyInvalidationCount();
    cache.InvalidateFamily(targetFamily);

    const size_t firstPassTouched = cache.DebugGetLastFamilyInvalidationCount();
    if (firstPassTouched == 0 || firstPassTouched > 3) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex", L"Unexpected entry count touched during first invalidation");
        return false;
    }
    if (firstPassTouched >= 100) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex",
                     L"Invalidation touched too many entries, secondary index may be ineffective");
        return false;
    }

    shelltabs::IconCache::Reference refreshedSmall = cache.Acquire(targetFamily, SHGFI_SMALLICON, loader);
    if (!refreshedSmall) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex", L"Failed to load refreshed small icon");
        return false;
    }
    if (refreshedSmall.Get() == initialHandle) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex",
                     L"Refreshed icon reused stale handle, stale token handling broken");
        return false;
    }

    inUseSmall.Reset();

    cache.DebugResetLastFamilyInvalidationCount();
    cache.InvalidateFamily(targetFamily);
    const size_t secondPassTouched = cache.DebugGetLastFamilyInvalidationCount();
    if (secondPassTouched > 2) {
        PrintFailure(L"TestInvalidateFamilyScalesWithIndex",
                     L"Second invalidation touched unexpected number of entries");
        return false;
    }

    refreshedSmall.Reset();
    cache.DebugSetCapacity(128);

    return true;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestInvalidateFamilyScalesWithIndex", &TestInvalidateFamilyScalesWithIndex},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            std::wcerr << L"[FAILED] " << test.name << std::endl;
            success = false;
        } else {
            std::wcout << L"[PASSED] " << test.name << std::endl;
        }
    }

    if (!success) {
        std::wcerr << L"Icon cache stress tests failed." << std::endl;
        return 1;
    }

    std::wcout << L"Icon cache stress tests passed." << std::endl;
    return 0;
}
