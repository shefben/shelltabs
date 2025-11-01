#include <windows.h>
#include <CommCtrl.h>

#include <iostream>
#include <string>
#include <vector>

#define private public
#define protected public
#include "ShellTabsListView.h"
#undef private
#undef protected

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

bool TestBackgroundResolverResetsSurface() {
    shelltabs::ShellTabsListView view;
    view.m_backgroundSurface.cacheKey = L"old";
    view.m_backgroundSurface.size = {123, 456};

    bool resolverCalled = false;
    auto resolver = [&]() -> shelltabs::ShellTabsListView::BackgroundSource {
        resolverCalled = true;
        shelltabs::ShellTabsListView::BackgroundSource source{};
        source.cacheKey = L"new";
        return source;
    };

    view.SetBackgroundResolver(resolver);

    if (!view.m_backgroundSurface.cacheKey.empty() || view.m_backgroundSurface.size.cx != 0 ||
        view.m_backgroundSurface.size.cy != 0) {
        PrintFailure(L"TestBackgroundResolverResetsSurface",
                     L"Background surface was not cleared after resolver update");
        return false;
    }

    if (!view.m_backgroundResolver) {
        PrintFailure(L"TestBackgroundResolverResetsSurface", L"Resolver was not stored");
        return false;
    }

    const auto source = view.m_backgroundResolver();
    if (!resolverCalled || source.cacheKey != L"new") {
        PrintFailure(L"TestBackgroundResolverResetsSurface",
                     L"Resolver was not invoked or returned unexpected data");
        return false;
    }

    return true;
}

bool TestAccentResolverResetsState() {
    shelltabs::ShellTabsListView view;
    view.m_accentResources.accentColor = RGB(10, 20, 30);
    view.m_accentResources.textColor = RGB(200, 210, 220);

    bool resolverCalled = false;
    auto resolver = [&](COLORREF* accent, COLORREF* text) {
        resolverCalled = true;
        if (accent) {
            *accent = RGB(1, 2, 3);
        }
        if (text) {
            *text = RGB(4, 5, 6);
        }
        return true;
    };

    view.SetAccentColorResolver(resolver);

    if (view.m_accentResources.accentColor != 0 || view.m_accentResources.textColor != 0) {
        PrintFailure(L"TestAccentResolverResetsState", L"Accent resources were not reset");
        return false;
    }

    if (!view.m_accentResolver) {
        PrintFailure(L"TestAccentResolverResetsState", L"Resolver was not stored");
        return false;
    }

    COLORREF accent = 0;
    COLORREF text = 0;
    if (!view.m_accentResolver(&accent, &text) || !resolverCalled) {
        PrintFailure(L"TestAccentResolverResetsState", L"Resolver did not run as expected");
        return false;
    }

    if (accent != RGB(1, 2, 3) || text != RGB(4, 5, 6)) {
        PrintFailure(L"TestAccentResolverResetsState", L"Resolver returned unexpected colors");
        return false;
    }

    return true;
}

bool TestUseAccentColorsToggleResets() {
    shelltabs::ShellTabsListView view;
    view.m_useAccentColors = true;
    view.m_accentResources.accentColor = RGB(50, 60, 70);

    view.SetUseAccentColors(false);

    if (view.m_useAccentColors) {
        PrintFailure(L"TestUseAccentColorsToggleResets", L"Accent usage flag was not updated");
        return false;
    }

    if (view.m_accentResources.accentColor != 0) {
        PrintFailure(L"TestUseAccentColorsToggleResets", L"Accent resources were not cleared");
        return false;
    }

    return true;
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
        {L"TestBackgroundResolverResetsSurface", &TestBackgroundResolverResetsSurface},
        {L"TestAccentResolverResetsState", &TestAccentResolverResetsState},
        {L"TestUseAccentColorsToggleResets", &TestUseAccentColorsToggleResets},
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
