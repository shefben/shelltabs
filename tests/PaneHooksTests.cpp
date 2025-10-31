#include "PaneHooks.h"

#include <windows.h>
#include <CommCtrl.h>

#include <iostream>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace {

struct TestDefinition {
    const wchar_t* name;
    bool (*fn)();
};

void PrintFailure(const wchar_t* testName, const std::wstring& message) {
    std::wcerr << L"[" << testName << L"] " << message << std::endl;
}

class MockHighlightProvider : public shelltabs::PaneHighlightProvider {
public:
    bool TryGetListViewHighlight(HWND, int itemIndex, shelltabs::PaneHighlight* highlight) override {
        auto it = m_listHighlights.find(itemIndex);
        if (it == m_listHighlights.end()) {
            return false;
        }
        if (highlight) {
            *highlight = it->second;
        }
        return true;
    }

    bool TryGetTreeViewHighlight(HWND, HTREEITEM item, shelltabs::PaneHighlight* highlight) override {
        auto it = m_treeHighlights.find(item);
        if (it == m_treeHighlights.end()) {
            return false;
        }
        if (highlight) {
            *highlight = it->second;
        }
        return true;
    }

    void SetListHighlight(int index, const shelltabs::PaneHighlight& highlight) {
        m_listHighlights[index] = highlight;
    }

    void SetTreeHighlight(HTREEITEM item, const shelltabs::PaneHighlight& highlight) {
        m_treeHighlights[item] = highlight;
    }

private:
    std::unordered_map<int, shelltabs::PaneHighlight> m_listHighlights;
    std::unordered_map<HTREEITEM, shelltabs::PaneHighlight> m_treeHighlights;
};

std::vector<std::pair<HWND, shelltabs::HighlightPaneType>> g_invalidationEvents;

void ResetInvalidationTracking() {
    g_invalidationEvents.clear();
}

void TestInvalidationCallback(HWND hwnd, shelltabs::HighlightPaneType pane) {
    g_invalidationEvents.emplace_back(hwnd, pane);
}

bool TestListViewPrepaintRequestsCallbacks() {
    MockHighlightProvider provider;
    shelltabs::PaneHookRouter router(&provider);
    HWND listView = reinterpret_cast<HWND>(0x1234);
    router.SetListView(listView);

    NMLVCUSTOMDRAW customDraw{};
    customDraw.nmcd.hdr.hwndFrom = listView;
    customDraw.nmcd.dwDrawStage = CDDS_PREPAINT;

    LRESULT result = 0;
    if (!router.HandleNotify(&customDraw.nmcd.hdr, &result)) {
        PrintFailure(L"TestListViewPrepaintRequestsCallbacks", L"HandleNotify returned false");
        return false;
    }

    if (result != (CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW)) {
        PrintFailure(L"TestListViewPrepaintRequestsCallbacks", L"Unexpected result flags");
        return false;
    }

    return true;
}

bool TestListViewHighlightApplied() {
    MockHighlightProvider provider;
    shelltabs::PaneHookRouter router(&provider);
    HWND listView = reinterpret_cast<HWND>(0x2345);
    router.SetListView(listView);

    shelltabs::PaneHighlight highlight{};
    highlight.hasTextColor = true;
    highlight.textColor = RGB(10, 20, 30);
    highlight.hasBackgroundColor = true;
    highlight.backgroundColor = RGB(200, 210, 220);
    provider.SetListHighlight(0, highlight);

    NMLVCUSTOMDRAW customDraw{};
    customDraw.nmcd.hdr.hwndFrom = listView;
    customDraw.nmcd.dwDrawStage = CDDS_ITEMPREPAINT;
    customDraw.nmcd.dwItemSpec = 0;

    LRESULT result = 0;
    if (!router.HandleNotify(&customDraw.nmcd.hdr, &result)) {
        PrintFailure(L"TestListViewHighlightApplied", L"HandleNotify returned false");
        return false;
    }

    if (result != CDRF_NEWFONT) {
        PrintFailure(L"TestListViewHighlightApplied", L"Expected CDRF_NEWFONT");
        return false;
    }

    if (customDraw.clrText != highlight.textColor || customDraw.clrTextBk != highlight.backgroundColor) {
        PrintFailure(L"TestListViewHighlightApplied", L"Colors were not propagated");
        return false;
    }

    return true;
}

bool TestTreeViewHighlightApplied() {
    MockHighlightProvider provider;
    shelltabs::PaneHookRouter router(&provider);
    HWND treeView = reinterpret_cast<HWND>(0x3456);
    router.SetTreeView(treeView);

    HTREEITEM item = reinterpret_cast<HTREEITEM>(0x1);
    shelltabs::PaneHighlight highlight{};
    highlight.hasTextColor = true;
    highlight.textColor = RGB(100, 110, 120);
    provider.SetTreeHighlight(item, highlight);

    NMTVCUSTOMDRAW customDraw{};
    customDraw.nmcd.hdr.hwndFrom = treeView;
    customDraw.nmcd.dwDrawStage = CDDS_ITEMPREPAINT;
    customDraw.nmcd.dwItemSpec = reinterpret_cast<DWORD_PTR>(item);

    LRESULT result = 0;
    if (!router.HandleNotify(&customDraw.nmcd.hdr, &result)) {
        PrintFailure(L"TestTreeViewHighlightApplied", L"HandleNotify returned false");
        return false;
    }

    if (result != CDRF_NEWFONT) {
        PrintFailure(L"TestTreeViewHighlightApplied", L"Expected CDRF_NEWFONT");
        return false;
    }

    if (customDraw.clrText != highlight.textColor) {
        PrintFailure(L"TestTreeViewHighlightApplied", L"Tree text color mismatch");
        return false;
    }

    return true;
}

bool TestHighlightRegistryInvalidatesSubscribers() {
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
        PrintFailure(L"TestHighlightRegistryInvalidatesSubscribers", L"Expected two invalidation events");
        return false;
    }

    if (g_invalidationEvents[0].first != listView ||
        g_invalidationEvents[0].second != shelltabs::HighlightPaneType::ListView) {
        PrintFailure(L"TestHighlightRegistryInvalidatesSubscribers", L"List view invalidation missing");
        return false;
    }

    if (g_invalidationEvents[1].first != treeView ||
        g_invalidationEvents[1].second != shelltabs::HighlightPaneType::TreeView) {
        PrintFailure(L"TestHighlightRegistryInvalidatesSubscribers", L"Tree view invalidation missing");
        return false;
    }

    ResetInvalidationTracking();
    shelltabs::UnsubscribeTreeViewForHighlights(treeView);
    shelltabs::RegisterPaneHighlight(L"C:\\Temp\\file.txt", highlight);

    if (g_invalidationEvents.size() != 1) {
        PrintFailure(L"TestHighlightRegistryInvalidatesSubscribers",
                     L"Expected only list view to be invalidated after tree unsubscribe");
        return false;
    }

    if (g_invalidationEvents[0].first != listView ||
        g_invalidationEvents[0].second != shelltabs::HighlightPaneType::ListView) {
        PrintFailure(L"TestHighlightRegistryInvalidatesSubscribers", L"Unexpected invalidation target");
        return false;
    }

    shelltabs::UnsubscribeListViewForHighlights(listView);
    shelltabs::SetPaneHighlightInvalidationCallback(nullptr);
    shelltabs::ClearPaneHighlights();
    ResetInvalidationTracking();
    return true;
}

bool TestHighlightRegistryLookupIsCaseInsensitive() {
    shelltabs::ClearPaneHighlights();

    shelltabs::PaneHighlight highlight{};
    highlight.hasTextColor = true;
    highlight.textColor = RGB(11, 22, 33);
    shelltabs::RegisterPaneHighlight(L"C:\\Temp\\Folder", highlight);

    shelltabs::PaneHighlight fetched{};
    if (!shelltabs::TryGetPaneHighlight(L"c:/temp/folder", &fetched)) {
        PrintFailure(L"TestHighlightRegistryLookupIsCaseInsensitive",
                     L"Lower-case lookup did not resolve registered highlight");
        shelltabs::ClearPaneHighlights();
        return false;
    }

    if (!fetched.hasTextColor || fetched.textColor != highlight.textColor) {
        PrintFailure(L"TestHighlightRegistryLookupIsCaseInsensitive",
                     L"Retrieved highlight did not match registered value");
        shelltabs::ClearPaneHighlights();
        return false;
    }

    shelltabs::ClearPaneHighlights();
    return true;
}

bool TestHighlightRegistryLookupIgnoresTrailingSlash() {
    shelltabs::ClearPaneHighlights();

    shelltabs::PaneHighlight highlight{};
    highlight.hasBackgroundColor = true;
    highlight.backgroundColor = RGB(44, 55, 66);
    shelltabs::RegisterPaneHighlight(L"D:/Projects/Sample", highlight);

    shelltabs::PaneHighlight fetched{};
    if (!shelltabs::TryGetPaneHighlight(L"D:\\Projects\\Sample\\", &fetched)) {
        PrintFailure(L"TestHighlightRegistryLookupIgnoresTrailingSlash",
                     L"Lookup with trailing separator failed");
        shelltabs::ClearPaneHighlights();
        return false;
    }

    if (!fetched.hasBackgroundColor || fetched.backgroundColor != highlight.backgroundColor) {
        PrintFailure(L"TestHighlightRegistryLookupIgnoresTrailingSlash",
                     L"Trailing-slash lookup returned incorrect highlight");
        shelltabs::ClearPaneHighlights();
        return false;
    }

    shelltabs::ClearPaneHighlights();
    return true;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestListViewPrepaintRequestsCallbacks", &TestListViewPrepaintRequestsCallbacks},
        {L"TestListViewHighlightApplied", &TestListViewHighlightApplied},
        {L"TestTreeViewHighlightApplied", &TestTreeViewHighlightApplied},
        {L"TestHighlightRegistryInvalidatesSubscribers", &TestHighlightRegistryInvalidatesSubscribers},
        {L"TestHighlightRegistryLookupIsCaseInsensitive", &TestHighlightRegistryLookupIsCaseInsensitive},
        {L"TestHighlightRegistryLookupIgnoresTrailingSlash", &TestHighlightRegistryLookupIgnoresTrailingSlash},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
