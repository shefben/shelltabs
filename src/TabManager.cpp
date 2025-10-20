#include "TabManager.h"

#include <algorithm>
#include <numeric>
#include <optional>


namespace shelltabs {

namespace {
constexpr wchar_t kDefaultGroupNamePrefix[] = L"Island ";
}  // namespace

TabManager::TabManager() = default;

TabManager& TabManager::Get() {
    static TabManager instance;
    return instance;
}

TabLocation TabManager::SelectedLocation() const noexcept {
    if (m_selectedFloating && m_floatingTab) {
        return FloatingLocation();
    }
    return {m_selectedGroup, m_selectedTab};
}

int TabManager::TotalTabCount() const noexcept {
    int total = HasFloatingTab() ? 1 : 0;
    for (const auto& group : m_groups) {
        total += static_cast<int>(group.tabs.size());
    }
    return total;
}

void TabManager::SetSelectedLocation(TabLocation location) {
    if (!location.IsValid()) {
        m_selectedFloating = false;
        m_selectedGroup = -1;
        m_selectedTab = -1;
        return;
    }
    if (location.floating) {
        if (!HasFloatingTab()) {
            return;
        }
        m_selectedFloating = true;
        m_selectedGroup = -1;
        m_selectedTab = location.tabIndex;
        EnsureVisibleSelection();
        return;
    }
    m_selectedFloating = false;
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }
    m_selectedGroup = location.groupIndex;
    m_selectedTab = location.tabIndex;
    EnsureVisibleSelection();
}

const TabGroup* TabManager::GetGroup(int index) const noexcept {
    if (index < 0 || index >= static_cast<int>(m_groups.size())) {
        return nullptr;
    }
    return &m_groups[static_cast<size_t>(index)];
}

TabGroup* TabManager::GetGroup(int index) noexcept {
    if (index < 0 || index >= static_cast<int>(m_groups.size())) {
        return nullptr;
    }
    return &m_groups[static_cast<size_t>(index)];
}

const TabInfo* TabManager::Get(TabLocation location) const noexcept {
    if (!location.IsValid()) {
        return nullptr;
    }
    if (location.floating) {
        if (!HasFloatingTab() || location.tabIndex != 0) {
            return nullptr;
        }
        return &m_floatingTab.value();
    }
    const auto* group = GetGroup(location.groupIndex);
    if (!group) {
        return nullptr;
    }
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group->tabs.size())) {
        return nullptr;
    }
    return &group->tabs[static_cast<size_t>(location.tabIndex)];
}

TabInfo* TabManager::Get(TabLocation location) noexcept {
    if (!location.IsValid()) {
        return nullptr;
    }
    if (location.floating) {
        if (!HasFloatingTab() || location.tabIndex != 0) {
            return nullptr;
        }
        return &m_floatingTab.value();
    }
    auto* group = GetGroup(location.groupIndex);
    if (!group) {
        return nullptr;
    }
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group->tabs.size())) {
        return nullptr;
    }
    return &group->tabs[static_cast<size_t>(location.tabIndex)];
}

TabLocation TabManager::Find(PCIDLIST_ABSOLUTE pidl) const {
    if (!pidl) {
        return {};
    }
    if (HasFloatingTab() && ArePidlsEqual(m_floatingTab->pidl.get(), pidl)) {
        return FloatingLocation();
    }
    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            if (ArePidlsEqual(group.tabs[t].pidl.get(), pidl)) {
                return {static_cast<int>(g), static_cast<int>(t)};
            }
        }
    }
    return {};
}

TabLocation TabManager::Add(UniquePidl pidl, std::wstring name, std::wstring tooltip, bool select, int groupIndex) {
    if (!pidl) {
        return {};
    }

    TabInfo info{
        .pidl = std::move(pidl),
        .name = std::move(name),
        .tooltip = std::move(tooltip),
        .hidden = false,
    };
    info.path = GetParsingName(info.pidl.get());

    if (!HasFloatingTab() && m_groups.empty() && groupIndex < 0) {
        m_floatingTab.emplace(std::move(info));
        if (select) {
            m_selectedFloating = true;
            m_selectedGroup = -1;
            m_selectedTab = 0;
        }
        EnsureVisibleSelection();
        return FloatingLocation();
    }

    int resolvedGroup = groupIndex;
    if (HasFloatingTab()) {
        resolvedGroup = PromoteFloatingTabToGroup(true);
    }

    if (m_groups.empty()) {
        TabGroup group;
        group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
        group.headerVisible = true;
        m_groups.emplace_back(std::move(group));
        if (m_selectedGroup < 0) {
            m_selectedGroup = 0;
            m_selectedTab = -1;
        }
    }

    if (resolvedGroup < 0 || resolvedGroup >= static_cast<int>(m_groups.size())) {
        resolvedGroup = (m_selectedGroup >= 0) ? m_selectedGroup : 0;
    }

    auto& group = m_groups[static_cast<size_t>(resolvedGroup)];
    group.tabs.emplace_back(std::move(info));

    const TabLocation location{resolvedGroup, static_cast<int>(group.tabs.size() - 1)};
    if (select) {
        m_selectedFloating = false;
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
        group.collapsed = false;
    }

    EnsureVisibleSelection();
    return location;
}

void TabManager::Remove(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }
    if (location.floating) {
        if (!HasFloatingTab() || location.tabIndex != 0) {
            return;
        }
        const bool wasSelected = m_selectedFloating;
        m_floatingTab.reset();
        m_selectedFloating = false;
        if (wasSelected) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        }
        EnsureVisibleSelection();
        return;
    }
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }

    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }

    const bool wasSelected = (m_selectedGroup == location.groupIndex && m_selectedTab == location.tabIndex);

    group.tabs.erase(group.tabs.begin() + location.tabIndex);

    if (group.tabs.empty()) {
        m_groups.erase(m_groups.begin() + location.groupIndex);
        if (m_selectedGroup == location.groupIndex) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        } else if (m_selectedGroup > location.groupIndex) {
            --m_selectedGroup;
        }
    } else if (m_selectedGroup == location.groupIndex && m_selectedTab > location.tabIndex) {
        --m_selectedTab;
    }

    if (wasSelected) {
        if (!m_groups.empty()) {
            m_selectedGroup = std::min(std::max(location.groupIndex, 0), static_cast<int>(m_groups.size()) - 1);
            auto& newGroup = m_groups[static_cast<size_t>(m_selectedGroup)];
            if (!newGroup.tabs.empty()) {
                const int newIndex = std::min(location.tabIndex, static_cast<int>(newGroup.tabs.size()) - 1);
                m_selectedTab = std::max(newIndex, 0);
            } else {
                m_selectedTab = -1;
            }
        } else {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        }
    }

    EnsureVisibleSelection();
}

std::optional<TabInfo> TabManager::TakeTab(TabLocation location) {
    if (!location.IsValid()) {
        return std::nullopt;
    }
    if (location.floating) {
        if (!HasFloatingTab() || location.tabIndex != 0) {
            return std::nullopt;
        }
        TabInfo removed = std::move(m_floatingTab.value());
        const bool wasSelected = m_selectedFloating;
        m_floatingTab.reset();
        m_selectedFloating = false;
        if (wasSelected) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        }
        EnsureVisibleSelection();
        return removed;
    }
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return std::nullopt;
    }

    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return std::nullopt;
    }

    TabInfo removed = std::move(group.tabs[static_cast<size_t>(location.tabIndex)]);

    const bool wasSelected = (m_selectedGroup == location.groupIndex && m_selectedTab == location.tabIndex);

    group.tabs.erase(group.tabs.begin() + location.tabIndex);

    bool removedGroup = false;
    if (group.tabs.empty()) {
        m_groups.erase(m_groups.begin() + location.groupIndex);
        removedGroup = true;
        if (m_selectedGroup == location.groupIndex) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        } else if (m_selectedGroup > location.groupIndex) {
            --m_selectedGroup;
        }
    } else if (m_selectedGroup == location.groupIndex && m_selectedTab > location.tabIndex) {
        --m_selectedTab;
    }

    if (wasSelected) {
        if (!m_groups.empty()) {
            m_selectedGroup = std::min(std::max(location.groupIndex, 0), static_cast<int>(m_groups.size()) - 1);
            auto& newGroup = m_groups[static_cast<size_t>(m_selectedGroup)];
            if (!newGroup.tabs.empty()) {
                const int newIndex = std::min(location.tabIndex, static_cast<int>(newGroup.tabs.size()) - 1);
                m_selectedTab = std::max(newIndex, 0);
            } else {
                m_selectedTab = -1;
            }
        } else {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        }
    }

    EnsureVisibleSelection();

    return removed;
}

TabLocation TabManager::InsertTab(TabInfo tab, int groupIndex, int tabIndex, bool select) {
    if (HasFloatingTab()) {
        PromoteFloatingTabToGroup(true);
    }

    if (m_groups.empty()) {
        TabGroup group;
        group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
        group.headerVisible = true;
        m_groups.emplace_back(std::move(group));
        m_selectedGroup = 0;
        m_selectedTab = -1;
    }

    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        groupIndex = std::clamp(groupIndex, 0, static_cast<int>(m_groups.size()) - 1);
    }

    auto& group = m_groups[static_cast<size_t>(groupIndex)];
    const int insertIndex = std::clamp(tabIndex, 0, static_cast<int>(group.tabs.size()));

    group.tabs.insert(group.tabs.begin() + insertIndex, std::move(tab));

    if (select) {
        m_selectedGroup = groupIndex;
        m_selectedTab = insertIndex;
        group.collapsed = false;
    } else if (m_selectedGroup == groupIndex && m_selectedTab >= insertIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();

    return {groupIndex, insertIndex};
}

std::optional<TabGroup> TabManager::TakeGroup(int groupIndex) {
    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        return std::nullopt;
    }

    TabGroup removed = std::move(m_groups[static_cast<size_t>(groupIndex)]);
    const bool wasSelected = (m_selectedGroup == groupIndex);

    m_groups.erase(m_groups.begin() + groupIndex);

    if (wasSelected) {
        m_selectedGroup = -1;
        m_selectedTab = -1;
    } else if (m_selectedGroup > groupIndex) {
        --m_selectedGroup;
    }

    EnsureVisibleSelection();

    return removed;
}

int TabManager::InsertGroup(TabGroup group, int insertIndex) {
    if (HasFloatingTab()) {
        PromoteFloatingTabToGroup(true);
    }

    if (insertIndex < 0) {
        insertIndex = 0;
    }
    if (insertIndex > static_cast<int>(m_groups.size())) {
        insertIndex = static_cast<int>(m_groups.size());
    }

    const auto position = m_groups.begin() + insertIndex;
    m_groups.insert(position, std::move(group));

    if (m_selectedGroup >= insertIndex) {
        ++m_selectedGroup;
    }

    EnsureVisibleSelection();
    return insertIndex;
}

void TabManager::Clear() {
    m_groups.clear();
    m_selectedGroup = -1;
    m_selectedTab = -1;
    m_groupSequence = 1;
    m_floatingTab.reset();
    m_selectedFloating = false;
}

void TabManager::Restore(std::vector<TabGroup> groups, int selectedGroup, int selectedTab, int groupSequence) {
    m_groups = std::move(groups);
    m_selectedGroup = selectedGroup;
    m_selectedTab = selectedTab;
    m_groupSequence = std::max(groupSequence, 1);
    m_floatingTab.reset();
    m_selectedFloating = false;
    EnsureVisibleSelection();
}

std::vector<TabViewItem> TabManager::BuildView() const {
    std::vector<TabViewItem> items;
    items.reserve(TotalTabCount() + static_cast<int>(m_groups.size()));

    if (HasFloatingTab()) {
        const auto& tab = *m_floatingTab;
        if (!tab.hidden) {
            TabViewItem item;
            item.type = TabViewItemType::kTab;
            item.location = FloatingLocation();
            item.name = tab.name;
            item.tooltip = tab.tooltip.empty() ? tab.name : tab.tooltip;
            item.pidl = tab.pidl.get();
            item.selected = m_selectedFloating;
            item.path = tab.path;
            item.hasCustomOutline = false;
            item.outlineColor = 0;
            item.headerVisible = false;
            item.floating = true;
            items.emplace_back(std::move(item));
        }
    }

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        const size_t total = group.tabs.size();
        size_t visible = 0;
        size_t hidden = 0;

        for (const auto& tab : group.tabs) {
            if (tab.hidden) {
                ++hidden;
                continue;
            }
            ++visible;
        }

        if (group.headerVisible) {
            TabViewItem header;
            header.type = TabViewItemType::kGroupHeader;
            header.location = {static_cast<int>(g), -1};
            header.name = group.name;
            header.tooltip = group.name;
            header.pidl = nullptr;
            if (total > 0) {
                header.tooltip += L" (" + std::to_wstring(visible) + L" visible of " + std::to_wstring(total) + L")";
            }
            if (hidden > 0) {
                header.tooltip += L" - " + std::to_wstring(hidden) + L" hidden";
            }
            header.selected = (m_selectedGroup == static_cast<int>(g));
            header.collapsed = group.collapsed;
            header.totalTabs = total;
            header.visibleTabs = visible;
            header.hiddenTabs = hidden;
            header.hasCustomOutline = group.hasCustomOutline;
            header.outlineColor = group.outlineColor;
            header.savedGroupId = group.savedGroupId;
            header.isSavedGroup = !group.savedGroupId.empty();
            header.headerVisible = group.headerVisible;
            header.floating = false;
            items.emplace_back(std::move(header));
        }

        if (group.collapsed) {
            continue;
        }

        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            if (tab.hidden) {
                continue;
            }

            TabViewItem item;
            item.type = TabViewItemType::kTab;
            item.location = {static_cast<int>(g), static_cast<int>(t)};
            item.name = tab.name;
            item.tooltip = tab.tooltip.empty() ? tab.name : tab.tooltip;
            item.pidl = tab.pidl.get();
            item.selected = (m_selectedGroup == static_cast<int>(g) && m_selectedTab == static_cast<int>(t));
            item.path = tab.path;
            item.hasCustomOutline = group.hasCustomOutline;
            item.outlineColor = group.outlineColor;
            item.savedGroupId = group.savedGroupId;
            item.isSavedGroup = !group.savedGroupId.empty();
            item.headerVisible = group.headerVisible;
            item.floating = false;

            items.emplace_back(std::move(item));
        }
    }

    return items;
}

void TabManager::ToggleGroupCollapsed(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    group->collapsed = !group->collapsed;
    if (!group->collapsed) {
        EnsureVisibleSelection();
    }
}

void TabManager::SetGroupCollapsed(int groupIndex, bool collapsed) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    group->collapsed = collapsed;
    if (!collapsed) {
        EnsureVisibleSelection();
    }
}

void TabManager::HideTab(TabLocation location) {
    auto* tab = Get(location);
    if (!tab) {
        return;
    }
    const bool wasSelected = location.floating ? m_selectedFloating
                                               : (m_selectedGroup == location.groupIndex &&
                                                  m_selectedTab == location.tabIndex);
    tab->hidden = true;
    if (location.floating) {
        if (wasSelected) {
            m_selectedFloating = false;
        }
    }
    if (wasSelected) {
        EnsureVisibleSelection();
    }
}

void TabManager::UnhideTab(TabLocation location) {
    auto* tab = Get(location);
    if (!tab) {
        return;
    }
    tab->hidden = false;
    if (location.floating) {
        if (!m_selectedFloating && (m_selectedGroup < 0 || m_groups.empty())) {
            m_selectedFloating = true;
            m_selectedGroup = -1;
            m_selectedTab = 0;
        }
    } else if (m_selectedGroup < 0 || m_selectedGroup >= static_cast<int>(m_groups.size())) {
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
    }
    EnsureVisibleSelection();
}

void TabManager::UnhideAllInGroup(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    for (auto& tab : group->tabs) {
        tab.hidden = false;
    }
    if (m_selectedGroup < 0 || m_selectedGroup >= static_cast<int>(m_groups.size())) {
        m_selectedGroup = groupIndex;
        m_selectedTab = group->tabs.empty() ? -1 : 0;
    }
    EnsureVisibleSelection();
}

std::vector<std::pair<TabLocation, std::wstring>> TabManager::GetHiddenTabs(int groupIndex) const {
    std::vector<std::pair<TabLocation, std::wstring>> result;
    const auto* group = GetGroup(groupIndex);
    if (!group) {
        return result;
    }
    for (size_t i = 0; i < group->tabs.size(); ++i) {
        const auto& tab = group->tabs[i];
        if (!tab.hidden) {
            continue;
        }
        result.emplace_back(TabLocation{groupIndex, static_cast<int>(i)}, tab.name);
    }
    return result;
}

size_t TabManager::HiddenCount(int groupIndex) const {
    const auto* group = GetGroup(groupIndex);
    if (!group) {
        return 0;
    }
    return static_cast<size_t>(std::count_if(group->tabs.begin(), group->tabs.end(), [](const TabInfo& tab) {
        return tab.hidden;
    }));
}

int TabManager::CreateGroupAfter(int groupIndex, std::wstring name, bool headerVisible) {
    if (HasFloatingTab()) {
        PromoteFloatingTabToGroup(true);
    }
    if (groupIndex < -1 || groupIndex >= static_cast<int>(m_groups.size())) {
        groupIndex = static_cast<int>(m_groups.size()) - 1;
    }

    TabGroup group;
    if (name.empty()) {
        group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(++m_groupSequence);
    } else {
        group.name = std::move(name);
    }
    group.headerVisible = headerVisible;

    const int insertIndex = groupIndex + 1;
    const auto position = m_groups.begin() + std::clamp(insertIndex, 0, static_cast<int>(m_groups.size()));
    m_groups.insert(position, std::move(group));

    if (m_selectedGroup >= insertIndex) {
        ++m_selectedGroup;
    }

    EnsureVisibleSelection();

    return std::clamp(insertIndex, 0, static_cast<int>(m_groups.size()) - 1);
}

void TabManager::MoveTab(TabLocation from, TabLocation to) {
    if (!from.IsValid()) {
        return;
    }
    if (from.floating) {
        if (!HasFloatingTab() || from.tabIndex != 0) {
            return;
        }
        TabInfo movingTab = std::move(m_floatingTab.value());
        const bool wasSelected = m_selectedFloating;
        m_floatingTab.reset();
        m_selectedFloating = false;

        if (m_groups.empty()) {
            TabGroup group;
            group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
            group.headerVisible = true;
            m_groups.emplace_back(std::move(group));
            m_selectedGroup = 0;
            m_selectedTab = -1;
        }

        if (to.groupIndex < 0 || to.groupIndex >= static_cast<int>(m_groups.size())) {
            to.groupIndex = std::clamp(to.groupIndex, 0, static_cast<int>(m_groups.size()) - 1);
        }

        auto& destinationGroup = m_groups[static_cast<size_t>(to.groupIndex)];
        if (to.tabIndex < 0 || to.tabIndex > static_cast<int>(destinationGroup.tabs.size())) {
            to.tabIndex = static_cast<int>(destinationGroup.tabs.size());
        }

        destinationGroup.tabs.insert(destinationGroup.tabs.begin() + to.tabIndex, std::move(movingTab));

        if (wasSelected) {
            m_selectedGroup = to.groupIndex;
            m_selectedTab = to.tabIndex;
        } else if (m_selectedGroup == to.groupIndex && m_selectedTab >= to.tabIndex) {
            ++m_selectedTab;
        }

        EnsureVisibleSelection();
        return;
    }

    if (from.groupIndex < 0 || from.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }

    if (to.floating) {
        // Moving into a floating slot is not supported once groups exist.
        return;
    }

    if (to.groupIndex < 0 || to.groupIndex >= static_cast<int>(m_groups.size())) {
        to.groupIndex = from.groupIndex;
    }

    auto& sourceGroup = m_groups[static_cast<size_t>(from.groupIndex)];
    if (from.tabIndex < 0 || from.tabIndex >= static_cast<int>(sourceGroup.tabs.size())) {
        return;
    }

    TabInfo movingTab = std::move(sourceGroup.tabs[static_cast<size_t>(from.tabIndex)]);
    const bool wasSelected = (m_selectedGroup == from.groupIndex && m_selectedTab == from.tabIndex);

    sourceGroup.tabs.erase(sourceGroup.tabs.begin() + from.tabIndex);

    if (m_selectedGroup == from.groupIndex && m_selectedTab > from.tabIndex) {
        --m_selectedTab;
    }

    bool removedSourceGroup = false;
    if (sourceGroup.tabs.empty()) {
        m_groups.erase(m_groups.begin() + from.groupIndex);
        removedSourceGroup = true;
        if (m_selectedGroup == from.groupIndex) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        } else if (m_selectedGroup > from.groupIndex) {
            --m_selectedGroup;
        }
    }

    if (removedSourceGroup && to.groupIndex > from.groupIndex) {
        --to.groupIndex;
    }

    if (m_groups.empty()) {
        TabGroup group;
        group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
        group.headerVisible = true;
        m_groups.emplace_back(std::move(group));
        to.groupIndex = 0;
        m_selectedGroup = 0;
        m_selectedTab = -1;
    }

    to.groupIndex = std::clamp(to.groupIndex, 0, static_cast<int>(m_groups.size()) - 1);
    auto& destinationGroup = m_groups[static_cast<size_t>(to.groupIndex)];

    if (to.tabIndex < 0 || to.tabIndex > static_cast<int>(destinationGroup.tabs.size())) {
        to.tabIndex = static_cast<int>(destinationGroup.tabs.size());
    }

    destinationGroup.tabs.insert(destinationGroup.tabs.begin() + to.tabIndex, std::move(movingTab));

    if (wasSelected) {
        m_selectedGroup = to.groupIndex;
        m_selectedTab = to.tabIndex;
    } else if (m_selectedGroup == to.groupIndex && m_selectedTab >= to.tabIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();
}

void TabManager::MoveGroup(int fromGroup, int toGroup) {
    const int groupCount = static_cast<int>(m_groups.size());
    if (fromGroup < 0 || fromGroup >= groupCount) {
        return;
    }
    if (toGroup < 0) {
        toGroup = 0;
    }
    if (toGroup > groupCount) {
        toGroup = groupCount;
    }
    if (fromGroup == toGroup || fromGroup + 1 == toGroup) {
        return;
    }

    auto moving = std::move(m_groups[static_cast<size_t>(fromGroup)]);
    m_groups.erase(m_groups.begin() + fromGroup);

    int selectedGroup = m_selectedGroup;
    if (selectedGroup == fromGroup) {
        selectedGroup = -1;
    } else if (selectedGroup > fromGroup) {
        --selectedGroup;
    }

    if (toGroup > fromGroup) {
        --toGroup;
    }
    if (toGroup < 0) {
        toGroup = 0;
    }
    if (toGroup > static_cast<int>(m_groups.size())) {
        toGroup = static_cast<int>(m_groups.size());
    }

    m_groups.insert(m_groups.begin() + toGroup, std::move(moving));

    if (selectedGroup == -1) {
        m_selectedGroup = toGroup;
    } else {
        if (selectedGroup >= toGroup) {
            ++selectedGroup;
        }
        m_selectedGroup = selectedGroup;
    }

    EnsureVisibleSelection();
}

TabLocation TabManager::MoveTabToNewGroup(TabLocation from, int insertIndex, bool headerVisible) {
    if (!from.IsValid()) {
        return {};
    }
    const int targetIndex = CreateGroupAfter(insertIndex - 1, {}, headerVisible);
    MoveTab(from, {targetIndex, 0});
    return {targetIndex, 0};
}

void TabManager::SetGroupHeaderVisible(int groupIndex, bool visible) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    if (group->headerVisible == visible) {
        return;
    }
    group->headerVisible = visible;
    if (!visible && group->collapsed) {
        group->collapsed = false;
    }
}

bool TabManager::IsGroupHeaderVisible(int groupIndex) const {
    const auto* group = GetGroup(groupIndex);
    if (!group) {
        return false;
    }
    return group->headerVisible;
}

void TabManager::EnsureVisibleSelection() {
    if (m_selectedFloating) {
        if (!HasFloatingTab() || m_floatingTab->hidden) {
            m_selectedFloating = false;
        } else {
            m_selectedGroup = -1;
            m_selectedTab = 0;
            return;
        }
    }

    if (HasFloatingTab() && !m_floatingTab->hidden) {
        if (m_groups.empty() || m_selectedGroup < 0) {
            m_selectedFloating = true;
            m_selectedGroup = -1;
            m_selectedTab = 0;
            return;
        }
    }

    m_selectedFloating = false;

    if (m_groups.empty()) {
        m_selectedGroup = -1;
        m_selectedTab = -1;
        return;
    }

    if (m_selectedGroup < 0 || m_selectedGroup >= static_cast<int>(m_groups.size())) {
        m_selectedGroup = 0;
    }

    auto& group = m_groups[static_cast<size_t>(m_selectedGroup)];
    if (group.tabs.empty()) {
        m_selectedTab = -1;
        return;
    }

    if (m_selectedTab < 0 || m_selectedTab >= static_cast<int>(group.tabs.size())) {
        m_selectedTab = 0;
    }

    if (group.tabs[static_cast<size_t>(m_selectedTab)].hidden) {
        int newIndex = -1;
        for (size_t i = 0; i < group.tabs.size(); ++i) {
            if (!group.tabs[i].hidden) {
                newIndex = static_cast<int>(i);
                break;
            }
        }
        m_selectedTab = newIndex;
    }

    if (m_selectedTab < 0 && !group.tabs.empty()) {
        return;
    }

}

TabLocation TabManager::FloatingLocation() const noexcept {
    return {-1, 0, true};
}

int TabManager::PromoteFloatingTabToGroup(bool headerVisible) {
    if (!HasFloatingTab()) {
        return -1;
    }

    TabGroup group;
    group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
    group.headerVisible = headerVisible;
    group.tabs.emplace_back(std::move(m_floatingTab.value()));

    const bool wasSelected = m_selectedFloating;
    m_floatingTab.reset();
    m_selectedFloating = false;

    const int insertIndex = 0;
    if (m_selectedGroup >= insertIndex) {
        ++m_selectedGroup;
    }
    m_groups.insert(m_groups.begin() + insertIndex, std::move(group));

    if (wasSelected) {
        m_selectedGroup = insertIndex;
        m_selectedTab = 0;
    }

    return insertIndex;
}

}  // namespace shelltabs
