#include "TaskbarTabProvider.h"
#include "TabManager.h"
#include "ShellTabsMessages.h"
#include "Utilities.h"

namespace shelltabs {

class ShellTabsTabProvider : public ITaskbarTabProvider {
public:
    bool IsMultiTabWindow(HWND hwnd) const override {
        auto* mgr = TabManager::FindByHwnd(hwnd);
        return mgr && mgr->TotalTabCount() > 1;
    }

    TaskbarTabSnapshot QueryTabs(HWND hwnd) const override {
        TaskbarTabSnapshot snap;
        snap.explorerHwnd = hwnd;

        auto* mgr = TabManager::FindByHwnd(hwnd);
        if (!mgr) return snap;

        wchar_t title[256] = {};
        GetWindowTextW(hwnd, title, 256);
        snap.windowTitle = title;

        auto selected = mgr->SelectedLocation();
        int flatIndex = 0;

        for (int g = 0; g < mgr->GroupCount(); g++) {
            const auto* group = mgr->GetGroup(g);
            if (!group) continue;
            for (int t = 0; t < static_cast<int>(group->tabs.size()); t++) {
                const auto& tab = group->tabs[t];
                if (tab.hidden) continue;

                TaskbarTabEntry entry;
                entry.groupIndex = g;
                entry.tabIndex = t;
                entry.name = tab.name;
                entry.path = tab.path;
                entry.pidl = tab.pidl.get();
                entry.isPinned = tab.pinned;
                entry.isSelected = (g == selected.groupIndex && t == selected.tabIndex);
                if (entry.isSelected) snap.selectedIndex = flatIndex;

                snap.tabs.push_back(std::move(entry));
                flatIndex++;
            }
        }
        return snap;
    }

    bool ActivateTab(HWND hwnd, int groupIndex, int tabIndex) override {
        HWND bandWnd = FindDescendantByClassEnum(hwnd, L"ShellTabsBandWindow");
        if (!bandWnd) return false;

        SendMessageW(bandWnd, WM_SHELLTABS_SELECT_TAB,
                     static_cast<WPARAM>(groupIndex),
                     static_cast<LPARAM>(tabIndex));
        return true;
    }

    void ClearCachedResources() override {
        // No persistent caches in this implementation
    }
};

std::unique_ptr<ITaskbarTabProvider> CreateShellTabsTabProvider() {
    return std::make_unique<ShellTabsTabProvider>();
}

}  // namespace shelltabs
