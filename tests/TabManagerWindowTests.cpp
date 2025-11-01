#define private public
#define protected public
#include "TabManager.h"
#undef private
#undef protected

#include <windows.h>

#include <cmath>
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

bool TestCollectProgressSnapshot() {
    shelltabs::TabManager manager;
    manager.Clear();

    auto first = manager.Add({}, L"First", L"First", false);
    auto second = manager.Add({}, L"Second", L"Second", false);

    auto* firstTab = manager.Get(first);
    auto* secondTab = manager.Get(second);
    if (!firstTab || !secondTab) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Failed to retrieve inserted tabs");
        return false;
    }

    firstTab->progress.active = true;
    firstTab->progress.indeterminate = false;
    firstTab->progress.fraction = 0.5;
    firstTab->lastActivatedTick = 1234;
    firstTab->activationOrdinal = 42;

    secondTab->hidden = true;
    secondTab->lastActivatedTick = 1000;
    secondTab->activationOrdinal = 21;

    const auto snapshot = manager.CollectProgressStates();
    if (snapshot.size() != 2) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Unexpected snapshot size");
        return false;
    }

    const auto& header = snapshot[0];
    if (header.type != shelltabs::TabViewItemType::kGroupHeader || header.location.groupIndex != 0 ||
        header.location.tabIndex != -1) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Header entry mismatch");
        return false;
    }
    if (header.lastActivatedTick != firstTab->lastActivatedTick ||
        header.activationOrdinal != firstTab->activationOrdinal) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Header activation data mismatch");
        return false;
    }

    const auto& tabEntry = snapshot[1];
    if (tabEntry.type != shelltabs::TabViewItemType::kTab || tabEntry.location.groupIndex != first.groupIndex ||
        tabEntry.location.tabIndex != first.tabIndex) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Tab entry location mismatch");
        return false;
    }
    if (!tabEntry.progress.visible || tabEntry.progress.indeterminate ||
        std::abs(tabEntry.progress.fraction - 0.5) > 1e-4) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Tab progress mismatch");
        return false;
    }
    if (tabEntry.lastActivatedTick != firstTab->lastActivatedTick ||
        tabEntry.activationOrdinal != firstTab->activationOrdinal) {
        PrintFailure(L"TestCollectProgressSnapshot", L"Tab activation data mismatch");
        return false;
    }

    return true;
}

bool TestLookupAfterTabMovesAndRemovals() {
    shelltabs::TabManager manager;
    manager.Clear();

    const std::wstring pathA = L"C:\\Lookup\\A";
    const std::wstring pathB = L"C:\\Lookup\\B";
    const std::wstring pathC = L"C:\\Lookup\\C";

    shelltabs::TabInfo tabA;
    tabA.name = L"A";
    tabA.path = pathA;
    auto locA = manager.InsertTab(std::move(tabA), 0, 0, false);

    shelltabs::TabInfo tabB;
    tabB.name = L"B";
    tabB.path = pathB;
    auto locB = manager.InsertTab(std::move(tabB), 0, 1, false);

    shelltabs::TabInfo tabC;
    tabC.name = L"C";
    tabC.path = pathC;
    manager.InsertTab(std::move(tabC), 0, 2, false);

    auto foundA = manager.FindByPath(pathA);
    if (foundA.groupIndex != locA.groupIndex || foundA.tabIndex != locA.tabIndex) {
        PrintFailure(L"TestLookupAfterTabMovesAndRemovals", L"Initial lookup for tab A failed");
        return false;
    }
    auto foundB = manager.FindByPath(pathB);
    if (foundB.groupIndex != locB.groupIndex || foundB.tabIndex != locB.tabIndex) {
        PrintFailure(L"TestLookupAfterTabMovesAndRemovals", L"Initial lookup for tab B failed");
        return false;
    }

    manager.MoveTab(foundA, {foundA.groupIndex, 3});
    auto movedA = manager.FindByPath(pathA);
    if (movedA.groupIndex != foundA.groupIndex || movedA.tabIndex != 2) {
        PrintFailure(L"TestLookupAfterTabMovesAndRemovals", L"Lookup after intra-group move returned wrong location");
        return false;
    }

    const int newGroupIndex = manager.CreateGroupAfter(0, L"Moved", true);
    manager.MoveTab(movedA, {newGroupIndex, 0});
    auto movedToGroup = manager.FindByPath(pathA);
    if (movedToGroup.groupIndex != newGroupIndex || movedToGroup.tabIndex != 0) {
        PrintFailure(L"TestLookupAfterTabMovesAndRemovals", L"Lookup after moving to new group failed");
        return false;
    }

    manager.Remove(movedToGroup);
    if (manager.FindByPath(pathA).IsValid()) {
        PrintFailure(L"TestLookupAfterTabMovesAndRemovals", L"Lookup returned removed tab");
        return false;
    }

    auto updatedB = manager.FindByPath(pathB);
    if (!updatedB.IsValid() || updatedB.groupIndex != 0 || updatedB.tabIndex != 0) {
        PrintFailure(L"TestLookupAfterTabMovesAndRemovals", L"Lookup for remaining tab B returned unexpected location");
        return false;
    }

    return true;
}

bool TestLookupAfterGroupMove() {
    shelltabs::TabManager manager;
    manager.Clear();

    const std::wstring pathA = L"C:\\Groups\\A";
    const std::wstring pathB = L"C:\\Groups\\B";

    shelltabs::TabInfo tabA;
    tabA.name = L"A";
    tabA.path = pathA;
    manager.InsertTab(std::move(tabA), 0, 0, false);

    const int secondGroupIndex = manager.CreateGroupAfter(0, L"Second", true);

    shelltabs::TabInfo tabB;
    tabB.name = L"B";
    tabB.path = pathB;
    manager.InsertTab(std::move(tabB), secondGroupIndex, 0, false);

    auto beforeMoveA = manager.FindByPath(pathA);
    auto beforeMoveB = manager.FindByPath(pathB);
    if (beforeMoveA.groupIndex != 0 || beforeMoveB.groupIndex != secondGroupIndex) {
        PrintFailure(L"TestLookupAfterGroupMove", L"Initial group lookups incorrect");
        return false;
    }

    manager.MoveGroup(0, 2);
    auto afterMoveA = manager.FindByPath(pathA);
    auto afterMoveB = manager.FindByPath(pathB);
    if (afterMoveA.groupIndex != 1 || afterMoveA.tabIndex != 0) {
        PrintFailure(L"TestLookupAfterGroupMove", L"Tab A lookup incorrect after group move to end");
        return false;
    }
    if (afterMoveB.groupIndex != 0 || afterMoveB.tabIndex != 0) {
        PrintFailure(L"TestLookupAfterGroupMove", L"Tab B lookup incorrect after group move to end");
        return false;
    }

    manager.MoveGroup(1, 0);
    auto finalA = manager.FindByPath(pathA);
    auto finalB = manager.FindByPath(pathB);
    if (finalA.groupIndex != 0 || finalB.groupIndex != 1) {
        PrintFailure(L"TestLookupAfterGroupMove", L"Lookups incorrect after moving group back to front");
        return false;
    }

    return true;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestRegistrationLifecycle", &TestRegistrationLifecycle},
        {L"TestDestructorClearsRegistration", &TestDestructorClearsRegistration},
        {L"TestStressOpenCloseWindows", &TestStressOpenCloseWindows},
        {L"TestCollectProgressSnapshot", &TestCollectProgressSnapshot},
        {L"TestLookupAfterTabMovesAndRemovals", &TestLookupAfterTabMovesAndRemovals},
        {L"TestLookupAfterGroupMove", &TestLookupAfterGroupMove},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
