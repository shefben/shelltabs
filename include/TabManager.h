#pragma once

#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "Utilities.h"
#include "GitStatus.h"

namespace shelltabs {

enum class TabViewItemType {
    kGroupHeader,
    kTab,
};

struct TabLocation {
    int groupIndex = -1;
    int tabIndex = -1;

    bool IsValid() const noexcept { return groupIndex >= 0 && tabIndex >= 0; }
};

struct TabInfo {
    UniquePidl pidl;
    std::wstring name;
    std::wstring tooltip;
    bool hidden = false;
    std::wstring path;
};

struct TabGroup {
    std::wstring name;
    bool collapsed = false;
    std::vector<TabInfo> tabs;
    bool splitView = false;
    int splitPrimary = -1;
    int splitSecondary = -1;
    bool headerVisible = true;
    std::wstring savedGroupId;
    bool hasCustomOutline = false;
    COLORREF outlineColor = RGB(0, 120, 215);
};

struct TabViewItem {
    TabViewItemType type = TabViewItemType::kGroupHeader;
    TabLocation location;
    std::wstring name;
    std::wstring tooltip;
    PCIDLIST_ABSOLUTE pidl = nullptr;
    bool selected = false;
    bool collapsed = false;
    size_t totalTabs = 0;
    size_t visibleTabs = 0;
    size_t hiddenTabs = 0;
    bool hasTagColor = false;
    COLORREF tagColor = 0;
    std::vector<std::wstring> tags;
    bool hasGitStatus = false;
    GitStatusInfo gitStatus;
    bool splitActive = false;
    bool splitPrimary = false;
    bool splitSecondary = false;
    bool splitAvailable = false;
    bool splitEnabled = false;
    std::wstring path;
    bool hasCustomOutline = false;
    COLORREF outlineColor = 0;
    std::wstring savedGroupId;
    bool isSavedGroup = false;
    bool headerVisible = true;
};

class TabManager {
public:
    TabManager();

    int TotalTabCount() const noexcept;

    TabLocation SelectedLocation() const noexcept { return {m_selectedGroup, m_selectedTab}; }
    void SetSelectedLocation(TabLocation location);

    int GroupCount() const noexcept { return static_cast<int>(m_groups.size()); }
    const TabGroup* GetGroup(int index) const noexcept;
    TabGroup* GetGroup(int index) noexcept;

    const TabInfo* Get(TabLocation location) const noexcept;
    TabInfo* Get(TabLocation location) noexcept;
    TabLocation Find(PCIDLIST_ABSOLUTE pidl) const;

    TabLocation Add(UniquePidl pidl, std::wstring name, std::wstring tooltip, bool select, int groupIndex = -1);
    void Remove(TabLocation location);
    std::optional<TabInfo> TakeTab(TabLocation location);
    TabLocation InsertTab(TabInfo tab, int groupIndex, int tabIndex, bool select);
    std::optional<TabGroup> TakeGroup(int groupIndex);
    int InsertGroup(TabGroup group, int insertIndex);
    void Clear();
    void Restore(std::vector<TabGroup> groups, int selectedGroup, int selectedTab, int groupSequence);

    std::vector<TabViewItem> BuildView() const;

    void ToggleGroupCollapsed(int groupIndex);
    void SetGroupCollapsed(int groupIndex, bool collapsed);
    void HideTab(TabLocation location);
    void UnhideTab(TabLocation location);
    void UnhideAllInGroup(int groupIndex);
    std::vector<std::pair<TabLocation, std::wstring>> GetHiddenTabs(int groupIndex) const;
    size_t HiddenCount(int groupIndex) const;

    int CreateGroupAfter(int groupIndex, std::wstring name = {}, bool headerVisible = true);
    void MoveTab(TabLocation from, TabLocation to);
    void MoveGroup(int fromGroup, int toGroup);
    TabLocation MoveTabToNewGroup(TabLocation from, int insertIndex, bool headerVisible);
    void SetGroupHeaderVisible(int groupIndex, bool visible);
    bool IsGroupHeaderVisible(int groupIndex) const;

    void ToggleSplitView(int groupIndex);
    void SetSplitSecondary(TabLocation location);
    void ClearSplitSecondary(int groupIndex);
    TabLocation GetSplitSecondary(int groupIndex) const;
    bool IsSplitViewEnabled(int groupIndex) const;
    void SwapSplitSelection(int groupIndex);

    int NextGroupSequence() const noexcept { return m_groupSequence; }

private:
    void EnsureDefaultGroup();
    void EnsureVisibleSelection();
    void EnsureSplitIntegrity(int groupIndex);
    void EnsureSplitIntegrity(TabGroup& group);
    bool NextVisibleTabIndex(const TabGroup& group, int* index) const;

    std::vector<TabGroup> m_groups;
    int m_selectedGroup = -1;
    int m_selectedTab = -1;
    int m_groupSequence = 1;
};

}  // namespace shelltabs

