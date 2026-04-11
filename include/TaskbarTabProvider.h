#pragma once

#include <memory>
#include <string>
#include <vector>
#include <windows.h>
#include <shtypes.h>

namespace shelltabs {

struct TaskbarTabEntry {
    int groupIndex = -1;
    int tabIndex = -1;
    std::wstring name;
    std::wstring path;
    PCIDLIST_ABSOLUTE pidl = nullptr;  // non-owning, valid only during snapshot lifetime
    bool isSelected = false;
    bool isPinned = false;
};

struct TaskbarTabSnapshot {
    HWND explorerHwnd = nullptr;
    std::wstring windowTitle;
    std::vector<TaskbarTabEntry> tabs;
    int selectedIndex = -1;  // index into tabs[] of the active tab
};

class ITaskbarTabProvider {
public:
    virtual ~ITaskbarTabProvider() = default;
    virtual bool IsMultiTabWindow(HWND hwnd) const = 0;
    virtual TaskbarTabSnapshot QueryTabs(HWND hwnd) const = 0;
    virtual bool ActivateTab(HWND hwnd, int groupIndex, int tabIndex) = 0;
    virtual void ClearCachedResources() = 0;
};

// Factory function (implemented in ShellTabsTabProvider.cpp)
std::unique_ptr<ITaskbarTabProvider> CreateShellTabsTabProvider();

}  // namespace shelltabs
