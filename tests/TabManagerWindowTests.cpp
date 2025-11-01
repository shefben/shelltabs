#define private public
#define protected public
#include "TabManager.h"
#undef private
#undef protected

#include <windows.h>

#include <cmath>
#include <iostream>
#include <limits>
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

bool TestGroupAggregateMaintenance() {
    shelltabs::TabManager manager;
    manager.Clear();

    auto makeTab = [](const wchar_t* name, bool hidden = false) {
        shelltabs::TabInfo tab;
        tab.name = name;
        tab.tooltip = name;
        tab.hidden = hidden;
        return tab;
    };

    auto first = manager.InsertTab(makeTab(L"One"), 0, 0, true);
    auto second = manager.InsertTab(makeTab(L"Two"), 0, 1, false);
    auto third = manager.InsertTab(makeTab(L"Three"), 0, 2, false);

    auto verifyCounts = [&](size_t expectedVisible, size_t expectedHidden, const wchar_t* stage) {
        const auto* group = manager.GetGroup(0);
        if (!group) {
            PrintFailure(L"TestGroupAggregateMaintenance", std::wstring(L"Missing group during ") + stage);
            return false;
        }
        if (group->visibleCount != expectedVisible || group->hiddenCount != expectedHidden) {
            PrintFailure(L"TestGroupAggregateMaintenance",
                         std::wstring(stage) + L" visible/hidden mismatch: expected " +
                             std::to_wstring(expectedVisible) + L"/" + std::to_wstring(expectedHidden) + L", got " +
                             std::to_wstring(group->visibleCount) + L"/" +
                             std::to_wstring(group->hiddenCount));
            return false;
        }
        const auto view = manager.BuildView();
        if (view.empty() || view[0].type != shelltabs::TabViewItemType::kGroupHeader) {
            PrintFailure(L"TestGroupAggregateMaintenance", std::wstring(L"Missing header during ") + stage);
            return false;
        }
        if (view[0].visibleTabs != expectedVisible || view[0].hiddenTabs != expectedHidden) {
            PrintFailure(L"TestGroupAggregateMaintenance",
                         std::wstring(stage) + L" header aggregate mismatch: expected " +
                             std::to_wstring(expectedVisible) + L"/" + std::to_wstring(expectedHidden) + L", got " +
                             std::to_wstring(view[0].visibleTabs) + L"/" + std::to_wstring(view[0].hiddenTabs));
            return false;
        }
        if (manager.HiddenCount(0) != expectedHidden) {
            PrintFailure(L"TestGroupAggregateMaintenance",
                         std::wstring(stage) + L" HiddenCount mismatch: expected " +
                             std::to_wstring(expectedHidden) + L", got " +
                             std::to_wstring(manager.HiddenCount(0)));
            return false;
        }
        return true;
    };

    if (!verifyCounts(3, 0, L"initial")) {
        return false;
    }

    manager.HideTab(second);
    if (!verifyCounts(2, 1, L"after hide")) {
        return false;
    }

    manager.Remove(second);
    third.tabIndex = 1;
    if (!verifyCounts(2, 0, L"after remove")) {
        return false;
    }

    auto fourth = manager.InsertTab(makeTab(L"Four", true), 0, 2, false);
    if (!verifyCounts(2, 1, L"after insert hidden")) {
        return false;
    }

    manager.UnhideTab(fourth);
    if (!verifyCounts(3, 0, L"after unhide inserted")) {
        return false;
    }

    manager.HideTab(first);
    if (!verifyCounts(2, 1, L"after hide first")) {
        return false;
    }

    manager.UnhideAllInGroup(0);
    if (!verifyCounts(3, 0, L"after unhide all")) {
        return false;
    }

    return true;
}

bool TestLookupAfterMovesAndRemovals() {
    shelltabs::TabManager manager;
    manager.Clear();

    auto makeTab = [](const wchar_t* name, const wchar_t* path) {
        shelltabs::TabInfo tab;
        tab.name = name;
        tab.tooltip = name;
        tab.path = path;
        return tab;
    };

    const auto first = manager.InsertTab(makeTab(L"Alpha", L"C:\\Test\\Shared"), 0, 0, true);
    auto second = manager.InsertTab(makeTab(L"Beta", L"C:\\Test\\Second"), 0, 1, false);

    auto lookup = manager.FindByPath(L"c:\\TEST\\shared");
    if (!lookup.IsValid() || lookup.groupIndex != first.groupIndex || lookup.tabIndex != first.tabIndex) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Initial lookup did not resolve to inserted tab");
        return false;
    }

    manager.MoveTab(first, {first.groupIndex, 1});
    auto moved = manager.FindByPath(L"C:\\Test\\Shared");
    if (!moved.IsValid() || moved.groupIndex != first.groupIndex || moved.tabIndex != 1) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Lookup failed after moving tab within group");
        return false;
    }

    second = manager.FindByPath(L"C:\\Test\\Second");
    if (!second.IsValid()) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Failed to locate secondary tab prior to removal");
        return false;
    }

    manager.Remove(second);

    auto afterRemoval = manager.FindByPath(L"C:\\Test\\Shared");
    if (!afterRemoval.IsValid() || afterRemoval.groupIndex != 0 || afterRemoval.tabIndex != 0) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Lookup returned unexpected location after neighbor removal");
        return false;
    }

    if (manager.FindByPath(L"C:\\Test\\Second").IsValid()) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Removed tab was still discoverable by path");
        return false;
    }

    const int newGroup = manager.CreateGroupAfter(0, L"Later", true);
    manager.MoveTab(afterRemoval, {newGroup, 0});
    auto movedGroup = manager.FindByPath(L"C:\\Test\\Shared");
    if (!movedGroup.IsValid() || movedGroup.groupIndex != newGroup || movedGroup.tabIndex != 0) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Lookup failed after moving tab to new group");
        return false;
    }

    manager.Remove(movedGroup);
    if (manager.FindByPath(L"C:\\Test\\Shared").IsValid()) {
        PrintFailure(L"TestLookupAfterMovesAndRemovals", L"Lookup succeeded after tab deletion");
        return false;
    }

    return true;
}

bool TestActivationOrderSnapshot() {
    shelltabs::TabManager manager;
    manager.Clear();

    const auto first = manager.Add({}, L"One", L"One", false);
    const auto second = manager.Add({}, L"Two", L"Two", false);
    const auto third = manager.Add({}, L"Three", L"Three", false);

    auto* firstTab = manager.Get(first);
    auto* secondTab = manager.Get(second);
    auto* thirdTab = manager.Get(third);
    if (!firstTab || !secondTab || !thirdTab) {
        PrintFailure(L"TestActivationOrderSnapshot", L"Failed to retrieve inserted tabs");
        return false;
    }

    firstTab->activationOrdinal = 10;
    firstTab->lastActivatedTick = 100;
    firstTab->activationEpoch = 0;
    secondTab->activationOrdinal = 20;
    secondTab->lastActivatedTick = 300;
    secondTab->activationEpoch = 0;
    thirdTab->activationOrdinal = 5;
    thirdTab->lastActivatedTick = 200;
    thirdTab->activationEpoch = 0;

    manager.RebuildIndices();

    const auto fullOrder = manager.GetTabsByActivationOrder(true);
    if (fullOrder.size() != 3 || fullOrder[0].groupIndex != second.groupIndex ||
        fullOrder[0].tabIndex != second.tabIndex || fullOrder[1].groupIndex != first.groupIndex ||
        fullOrder[1].tabIndex != first.tabIndex || fullOrder[2].groupIndex != third.groupIndex ||
        fullOrder[2].tabIndex != third.tabIndex) {
        PrintFailure(L"TestActivationOrderSnapshot", L"Unexpected MRU order for all tabs");
        return false;
    }

    manager.HideTab(second);
    const auto visibleOrder = manager.GetTabsByActivationOrder(false);
    if (visibleOrder.size() != 2 || visibleOrder[0].groupIndex != first.groupIndex ||
        visibleOrder[0].tabIndex != first.tabIndex || visibleOrder[1].groupIndex != third.groupIndex ||
        visibleOrder[1].tabIndex != third.tabIndex) {
        PrintFailure(L"TestActivationOrderSnapshot", L"Unexpected MRU order after filtering hidden tab");
        return false;
    }

    return true;
}

bool TestActivationOrderWrapAndTickRegression() {
    shelltabs::TabManager manager;
    manager.Clear();

    const auto first = manager.Add({}, L"First", L"First", false);
    const auto second = manager.Add({}, L"Second", L"Second", false);

    auto* firstTab = manager.Get(first);
    auto* secondTab = manager.Get(second);
    if (!firstTab || !secondTab) {
        PrintFailure(L"TestActivationOrderWrapAndTickRegression", L"Failed to retrieve inserted tabs");
        return false;
    }

    const uint64_t maxOrdinal = std::numeric_limits<uint64_t>::max();
    firstTab->activationOrdinal = maxOrdinal;
    firstTab->lastActivatedTick = 9000;
    firstTab->activationEpoch = 0;
    secondTab->activationOrdinal = maxOrdinal;
    secondTab->lastActivatedTick = 8000;
    secondTab->activationEpoch = 0;

    manager.RebuildIndices();

    manager.m_lastActivationOrdinalSeen = maxOrdinal;
    manager.m_lastActivationTickSeen = 9000;
    manager.m_activationEpoch = 0;

    secondTab->activationOrdinal = 1;
    secondTab->lastActivatedTick = 5;
    manager.ActivationUpdateTab(second);

    auto order = manager.GetTabsByActivationOrder(true);
    if (order.empty() || order[0].groupIndex != second.groupIndex || order[0].tabIndex != second.tabIndex) {
        PrintFailure(L"TestActivationOrderWrapAndTickRegression", L"Ordinal wrap did not promote updated tab");
        return false;
    }

    manager.m_lastActivationOrdinalSeen = maxOrdinal;
    manager.m_lastActivationTickSeen = 5000;

    firstTab->activationOrdinal = maxOrdinal;
    firstTab->lastActivatedTick = 1;
    manager.ActivationUpdateTab(first);

    order = manager.GetTabsByActivationOrder(true);
    if (order.empty() || order[0].groupIndex != first.groupIndex || order[0].tabIndex != first.tabIndex) {
        PrintFailure(L"TestActivationOrderWrapAndTickRegression", L"Tick regression did not promote updated tab");
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
        {L"TestGroupAggregateMaintenance", &TestGroupAggregateMaintenance},
        {L"TestLookupAfterMovesAndRemovals", &TestLookupAfterMovesAndRemovals},
        {L"TestActivationOrderSnapshot", &TestActivationOrderSnapshot},
        {L"TestActivationOrderWrapAndTickRegression", &TestActivationOrderWrapAndTickRegression},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
