#include <windows.h>
#include <CommCtrl.h>

#include <iostream>
#include <string>
#include <vector>

#include "PaneHooks.h"

namespace {

struct TestDefinition {
    const wchar_t* name;
    bool (*fn)();
};

void PrintFailure(const wchar_t* testName, const std::wstring& message) {
    std::wcerr << L"[" << testName << L"] " << message << std::endl;
}

struct InvalidationEvent {
    HWND hwnd = nullptr;
    shelltabs::HighlightPaneType pane = shelltabs::HighlightPaneType::ListView;
    shelltabs::PaneHighlightInvalidationTargets targets;
};

std::vector<InvalidationEvent> g_invalidationEvents;

void ResetInvalidationTracking() { g_invalidationEvents.clear(); }

void TestInvalidationCallback(HWND hwnd, shelltabs::HighlightPaneType pane,
                              const shelltabs::PaneHighlightInvalidationTargets& targets) {
    g_invalidationEvents.push_back({hwnd, pane, targets});
}

bool TestHighlightRegistryNotifiesSubscribers() {
    ResetInvalidationTracking();
    shelltabs::SetPaneHighlightInvalidationCallback(&TestInvalidationCallback);

    HWND listView = reinterpret_cast<HWND>(0x4567);
    HWND treeView = reinterpret_cast<HWND>(0x5678);
    shelltabs::SubscribeListViewForHighlights(listView);
    shelltabs::SubscribeTreeViewForHighlights(treeView);

    shelltabs::PaneHighlight highlight{};
    highlight.hasTextColor = true;
    highlight.textColor = RGB(1, 2, 3);
    shelltabs::RegisterPaneHighlight(L"C:\\Temp\\file.txt", highlight);

    if (g_invalidationEvents.size() != 2) {
        PrintFailure(L"TestHighlightRegistryNotifiesSubscribers", L"Expected two invalidation events");
        return false;
    }

    if (g_invalidationEvents[0].hwnd != listView ||
        g_invalidationEvents[0].pane != shelltabs::HighlightPaneType::ListView) {
        PrintFailure(L"TestHighlightRegistryNotifiesSubscribers", L"List view invalidation missing");
        return false;
    }

    if (g_invalidationEvents[1].hwnd != treeView ||
        g_invalidationEvents[1].pane != shelltabs::HighlightPaneType::TreeView) {
        PrintFailure(L"TestHighlightRegistryNotifiesSubscribers", L"Tree view invalidation missing");
        return false;
    }

    ResetInvalidationTracking();
    shelltabs::UnsubscribeTreeViewForHighlights(treeView);
    shelltabs::RegisterPaneHighlight(L"C:\\Temp\\file.txt", highlight);

    if (g_invalidationEvents.size() != 1) {
        PrintFailure(L"TestHighlightRegistryNotifiesSubscribers",
                     L"Expected only list view to be invalidated after tree unsubscribe");
        return false;
    }

    if (g_invalidationEvents[0].hwnd != listView ||
        g_invalidationEvents[0].pane != shelltabs::HighlightPaneType::ListView) {
        PrintFailure(L"TestHighlightRegistryNotifiesSubscribers", L"Unexpected invalidation target");
        return false;
    }

    shelltabs::UnsubscribeListViewForHighlights(listView);
    shelltabs::SetPaneHighlightInvalidationCallback(nullptr);
    shelltabs::ClearPaneHighlights();
    ResetInvalidationTracking();
    return true;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestHighlightRegistryNotifiesSubscribers", &TestHighlightRegistryNotifiesSubscribers},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
