#include "TabBandWindow.h"

#include <windows.h>
#include <shellapi.h>

#include <iostream>
#include <string>
#include <vector>

namespace shelltabs {

struct TabBandWindowDiffTestHarness {
    using VisualItem = TabBandWindow::VisualItem;

    static void InitializeWindow(TabBandWindow& window, const RECT& rect) {
        window.m_hwnd = reinterpret_cast<HWND>(1);
        window.m_clientRect = rect;
    }

    static VisualItem MakeVisualItem(const TabViewItem& data, const RECT& bounds) {
        VisualItem item{};
        item.data = data;
        item.bounds = bounds;
        return item;
    }

    static void AssignIcon(VisualItem& item, IconCache::Reference icon, int width, int height) {
        item.icon = std::move(icon);
        item.iconWidth = width;
        item.iconHeight = height;
    }

    static TabBandWindow::LayoutDiffStats Diff(TabBandWindow& window,
                                               std::vector<VisualItem>& oldItems,
                                               std::vector<VisualItem>& newItems) {
        return window.ComputeLayoutDiff(oldItems, newItems);
    }

    static void Destroy(TabBandWindow& window, std::vector<VisualItem>& items) {
        window.DestroyVisualItemResources(items);
    }
};

}  // namespace shelltabs

namespace {

using namespace shelltabs;

struct TestDefinition {
    const wchar_t* name;
    bool (*fn)();
};

void PrintFailure(const wchar_t* testName, const std::wstring& message) {
    std::wcerr << L"[" << testName << L"] " << message << std::endl;
}

IconCache::Reference AcquireTestIcon(const wchar_t* family) {
    return IconCache::Instance().Acquire(family, SHGFI_SMALLICON, []() -> HICON {
        HICON base = LoadIconW(nullptr, IDI_APPLICATION);
        if (!base) {
            return nullptr;
        }
        return CopyIcon(base);
    });
}

bool TestIconPreservedWhenMovingGroups() {
    TabBandWindow window(nullptr);
    TabBandWindowDiffTestHarness::InitializeWindow(window, {0, 0, 400, 40});

    TabViewItem oldData;
    oldData.type = TabViewItemType::kTab;
    oldData.location = {0, 0};
    oldData.name = L"Tab";
    oldData.path = L"C:/Tabs/One";
    oldData.activationOrdinal = 42;

    auto icon = AcquireTestIcon(L"DiffMove");
    if (!icon) {
        PrintFailure(L"TestIconPreservedWhenMovingGroups", L"Failed to create icon reference");
        return false;
    }
    const HICON oldHandle = icon.Get();

    TabBandWindowDiffTestHarness::VisualItem oldItem =
        TabBandWindowDiffTestHarness::MakeVisualItem(oldData, {0, 0, 120, 24});
    TabBandWindowDiffTestHarness::AssignIcon(oldItem, std::move(icon), 16, 16);

    TabViewItem newData = oldData;
    newData.location.groupIndex = 1;
    newData.location.tabIndex = 1;

    TabBandWindowDiffTestHarness::VisualItem newItem =
        TabBandWindowDiffTestHarness::MakeVisualItem(newData, {150, 0, 280, 24});
    TabBandWindowDiffTestHarness::AssignIcon(newItem, AcquireTestIcon(L"DiffMove"), 16, 16);

    std::vector<TabBandWindowDiffTestHarness::VisualItem> oldItems;
    oldItems.emplace_back(std::move(oldItem));
    std::vector<TabBandWindowDiffTestHarness::VisualItem> newItems;
    newItems.emplace_back(std::move(newItem));

    auto stats = TabBandWindowDiffTestHarness::Diff(window, oldItems, newItems);

    bool success = true;
    if (!stats.removedIndices.empty()) {
        PrintFailure(L"TestIconPreservedWhenMovingGroups", L"Unexpected removed indices");
        success = false;
    }
    if (stats.inserted != 0 || stats.removed != 0) {
        PrintFailure(L"TestIconPreservedWhenMovingGroups", L"Unexpected insertion/removal counts");
        success = false;
    }
    if (newItems[0].icon.Get() != oldHandle) {
        PrintFailure(L"TestIconPreservedWhenMovingGroups", L"Icon handle was not transferred");
        success = false;
    }
    if (oldItems[0].icon.Get() != nullptr) {
        PrintFailure(L"TestIconPreservedWhenMovingGroups", L"Old item retained icon reference");
        success = false;
    }
    if (newItems[0].iconWidth != 16 || newItems[0].iconHeight != 16) {
        PrintFailure(L"TestIconPreservedWhenMovingGroups", L"Icon metrics not preserved");
        success = false;
    }

    TabBandWindowDiffTestHarness::Destroy(window, newItems);
    TabBandWindowDiffTestHarness::Destroy(window, oldItems);
    IconCache::Instance().InvalidateFamily(L"DiffMove");

    return success;
}

bool TestIconPreservedWhenResized() {
    TabBandWindow window(nullptr);
    TabBandWindowDiffTestHarness::InitializeWindow(window, {0, 0, 400, 40});

    TabViewItem data;
    data.type = TabViewItemType::kTab;
    data.location = {0, 0};
    data.name = L"Sized";
    data.path = L"C:/Tabs/Sized";
    data.activationOrdinal = 99;

    auto icon = AcquireTestIcon(L"DiffResize");
    if (!icon) {
        PrintFailure(L"TestIconPreservedWhenResized", L"Failed to create icon reference");
        return false;
    }
    const HICON oldHandle = icon.Get();

    TabBandWindowDiffTestHarness::VisualItem oldItem =
        TabBandWindowDiffTestHarness::MakeVisualItem(data, {0, 0, 140, 24});
    TabBandWindowDiffTestHarness::AssignIcon(oldItem, std::move(icon), 20, 20);

    TabBandWindowDiffTestHarness::VisualItem newItem =
        TabBandWindowDiffTestHarness::MakeVisualItem(data, {0, 0, 200, 24});
    TabBandWindowDiffTestHarness::AssignIcon(newItem, AcquireTestIcon(L"DiffResize"), 20, 20);

    std::vector<TabBandWindowDiffTestHarness::VisualItem> oldItems;
    oldItems.emplace_back(std::move(oldItem));
    std::vector<TabBandWindowDiffTestHarness::VisualItem> newItems;
    newItems.emplace_back(std::move(newItem));

    auto stats = TabBandWindowDiffTestHarness::Diff(window, oldItems, newItems);

    bool success = true;
    if (stats.removed != 0 || stats.inserted != 0) {
        PrintFailure(L"TestIconPreservedWhenResized", L"Unexpected removal/insertion counts");
        success = false;
    }
    if (newItems[0].icon.Get() != oldHandle) {
        PrintFailure(L"TestIconPreservedWhenResized", L"Icon handle mismatch after resize");
        success = false;
    }
    if (oldItems[0].icon.Get() != nullptr) {
        PrintFailure(L"TestIconPreservedWhenResized", L"Old item retained icon");
        success = false;
    }
    if (newItems[0].iconWidth != 20 || newItems[0].iconHeight != 20) {
        PrintFailure(L"TestIconPreservedWhenResized", L"Icon metrics changed during resize");
        success = false;
    }

    TabBandWindowDiffTestHarness::Destroy(window, newItems);
    TabBandWindowDiffTestHarness::Destroy(window, oldItems);
    IconCache::Instance().InvalidateFamily(L"DiffResize");

    return success;
}

}  // namespace

int wmain() {
    const std::vector<TestDefinition> tests = {
        {L"TestIconPreservedWhenMovingGroups", &TestIconPreservedWhenMovingGroups},
        {L"TestIconPreservedWhenResized", &TestIconPreservedWhenResized},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            success = false;
        }
    }

    return success ? 0 : 1;
}
