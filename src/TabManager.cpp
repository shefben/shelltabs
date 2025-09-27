#include "TabManager.h"

#include <algorithm>
#include <numeric>
#include <optional>

#include "GitStatus.h"
#include "Logging.h"
#include "Tagging.h"

namespace shelltabs {

namespace {
constexpr wchar_t kDefaultGroupNamePrefix[] = L"Island ";

COLORREF BlendColors(COLORREF a, COLORREF b) {
    const int r = (GetRValue(a) + GetRValue(b)) / 2;
    const int g = (GetGValue(a) + GetGValue(b)) / 2;
    const int bValue = (GetBValue(a) + GetBValue(b)) / 2;
    return RGB(r, g, bValue);
}
}  // namespace

TabManager::TabManager() { EnsureDefaultGroup(); }

int TabManager::TotalTabCount() const noexcept {
    int total = 0;
    for (const auto& group : m_groups) {
        total += static_cast<int>(group.tabs.size());
    }
    return total;
}

void TabManager::SetSelectedLocation(TabLocation location) {
    if (!location.IsValid()) {
        m_selectedGroup = -1;
        m_selectedTab = -1;
        return;
    }
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }
    m_selectedGroup = location.groupIndex;
    m_selectedTab = location.tabIndex;
    if (group.splitView) {
        group.splitPrimary = location.tabIndex;
        if (group.splitSecondary == group.splitPrimary) {
            group.splitSecondary = -1;
        }
        EnsureSplitIntegrity(location.groupIndex);
    }
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

    EnsureDefaultGroup();

    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        groupIndex = (m_selectedGroup >= 0) ? m_selectedGroup : 0;
    }

    TabInfo info{
        .pidl = std::move(pidl),
        .name = std::move(name),
        .tooltip = std::move(tooltip),
        .hidden = false,
    };
    info.path = GetParsingName(info.pidl.get());

    auto& group = m_groups[static_cast<size_t>(groupIndex)];
    group.tabs.emplace_back(std::move(info));

    const TabLocation location{groupIndex, static_cast<int>(group.tabs.size() - 1)};
    if (select) {
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
        group.collapsed = false;
    }

    EnsureSplitIntegrity(groupIndex);
    EnsureVisibleSelection();
    return location;
}

void TabManager::Remove(TabLocation location) {
    if (!location.IsValid()) {
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

    if (group.splitPrimary == location.tabIndex) {
        group.splitPrimary = -1;
    } else if (group.splitPrimary > location.tabIndex) {
        --group.splitPrimary;
    }
    if (group.splitSecondary == location.tabIndex) {
        group.splitSecondary = -1;
    } else if (group.splitSecondary > location.tabIndex) {
        --group.splitSecondary;
    }

    group.tabs.erase(group.tabs.begin() + location.tabIndex);

    if (group.tabs.empty()) {
        m_groups.erase(m_groups.begin() + location.groupIndex);
        if (m_selectedGroup == location.groupIndex) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        } else if (m_selectedGroup > location.groupIndex) {
            --m_selectedGroup;
        }
        EnsureDefaultGroup();
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

    EnsureSplitIntegrity(location.groupIndex);
    EnsureVisibleSelection();
}

void TabManager::Clear() {
    m_groups.clear();
    m_selectedGroup = -1;
    m_selectedTab = -1;
    m_groupSequence = 1;
    EnsureDefaultGroup();
}

void TabManager::Restore(std::vector<TabGroup> groups, int selectedGroup, int selectedTab, int groupSequence) {
    m_groups = std::move(groups);
    m_selectedGroup = selectedGroup;
    m_selectedTab = selectedTab;
    m_groupSequence = std::max(groupSequence, 1);
    EnsureDefaultGroup();
    EnsureVisibleSelection();
}

std::vector<TabViewItem> TabManager::BuildView() const {
    std::vector<TabViewItem> items;
    items.reserve(TotalTabCount() + static_cast<int>(m_groups.size()));

    auto& gitCache = GitStatusCache::Instance();
    const bool gitEnabled = gitCache.IsEnabled();

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        const size_t total = group.tabs.size();
        size_t visible = 0;
        size_t hidden = 0;
        bool groupHasColor = false;
        COLORREF groupColor = 0;

        for (const auto& tab : group.tabs) {
            if (tab.hidden) {
                ++hidden;
                continue;
            }
            ++visible;
            COLORREF color = 0;
            if (!tab.path.empty() && TagStore::Instance().TryGetColorForPath(tab.path, &color)) {
                if (!groupHasColor) {
                    groupColor = color;
                    groupHasColor = true;
                } else {
                    groupColor = BlendColors(groupColor, color);
                }
            }
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
            header.hasTagColor = groupHasColor;
            header.tagColor = groupColor;
            header.splitActive = group.splitView;
            header.splitEnabled = group.splitView;
            header.splitAvailable = visible > 1;
            header.hasCustomOutline = group.hasCustomOutline;
            header.outlineColor = group.outlineColor;
            header.savedGroupId = group.savedGroupId;
            header.isSavedGroup = !group.savedGroupId.empty();
            header.headerVisible = group.headerVisible;
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
            item.splitActive = group.splitView;
            item.splitEnabled = group.splitView;
            item.splitAvailable = visible > 1;
            item.splitPrimary = group.splitView && static_cast<int>(t) == group.splitPrimary;
            item.splitSecondary = group.splitView && static_cast<int>(t) == group.splitSecondary;
            item.hasCustomOutline = group.hasCustomOutline;
            item.outlineColor = group.outlineColor;
            item.savedGroupId = group.savedGroupId;
            item.isSavedGroup = !group.savedGroupId.empty();
            item.headerVisible = group.headerVisible;

            COLORREF color = 0;
            std::vector<std::wstring> tags;
            if (!tab.path.empty() && TagStore::Instance().TryGetColorAndTags(tab.path, &color, &tags)) {
                item.hasTagColor = true;
                item.tagColor = color;
                item.tags = std::move(tags);
            }

            if (gitEnabled && !tab.path.empty()) {
                const GitStatusInfo status = gitCache.Query(tab.path);
                if (status.isRepository) {
                    item.hasGitStatus = true;
                    item.gitStatus = status;
                }
            } else if (!gitEnabled && !tab.path.empty()) {
                LogMessage(LogLevel::Info, L"TabManager::BuildView skipping git status for %ls (disabled)",
                           tab.path.c_str());
            }

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
    tab->hidden = true;
    auto* group = GetGroup(location.groupIndex);
    if (group) {
        if (group->splitPrimary == location.tabIndex) {
            group->splitPrimary = -1;
        }
        if (group->splitSecondary == location.tabIndex) {
            group->splitSecondary = -1;
        }
        EnsureSplitIntegrity(location.groupIndex);
    }
    if (m_selectedGroup == location.groupIndex && m_selectedTab == location.tabIndex) {
        EnsureVisibleSelection();
    }
}

void TabManager::UnhideTab(TabLocation location) {
    auto* tab = Get(location);
    if (!tab) {
        return;
    }
    tab->hidden = false;
    if (m_selectedGroup < 0 || m_selectedGroup >= static_cast<int>(m_groups.size())) {
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
    }
    EnsureSplitIntegrity(location.groupIndex);
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
    EnsureSplitIntegrity(groupIndex);
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

    EnsureDefaultGroup();
    EnsureVisibleSelection();

    return std::clamp(insertIndex, 0, static_cast<int>(m_groups.size()) - 1);
}

void TabManager::MoveTab(TabLocation from, TabLocation to) {
    if (!from.IsValid()) {
        return;
    }
    if (from.groupIndex < 0 || from.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }

    if (to.groupIndex < 0 || to.groupIndex >= static_cast<int>(m_groups.size())) {
        to.groupIndex = from.groupIndex;
    }

    auto& sourceGroup = m_groups[static_cast<size_t>(from.groupIndex)];
    if (from.tabIndex < 0 || from.tabIndex >= static_cast<int>(sourceGroup.tabs.size())) {
        return;
    }

    if (sourceGroup.splitPrimary == from.tabIndex) {
        sourceGroup.splitPrimary = -1;
    } else if (sourceGroup.splitPrimary > from.tabIndex) {
        --sourceGroup.splitPrimary;
    }
    if (sourceGroup.splitSecondary == from.tabIndex) {
        sourceGroup.splitSecondary = -1;
    } else if (sourceGroup.splitSecondary > from.tabIndex) {
        --sourceGroup.splitSecondary;
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

    EnsureDefaultGroup();

    to.groupIndex = std::clamp(to.groupIndex, 0, static_cast<int>(m_groups.size()) - 1);
    auto& destinationGroup = m_groups[static_cast<size_t>(to.groupIndex)];

    if (to.tabIndex < 0 || to.tabIndex > static_cast<int>(destinationGroup.tabs.size())) {
        to.tabIndex = static_cast<int>(destinationGroup.tabs.size());
    }

    destinationGroup.tabs.insert(destinationGroup.tabs.begin() + to.tabIndex, std::move(movingTab));
    if (destinationGroup.splitPrimary >= to.tabIndex && destinationGroup.splitPrimary != -1) {
        ++destinationGroup.splitPrimary;
    }
    if (destinationGroup.splitSecondary >= to.tabIndex && destinationGroup.splitSecondary != -1) {
        ++destinationGroup.splitSecondary;
    }

    if (wasSelected) {
        m_selectedGroup = to.groupIndex;
        m_selectedTab = to.tabIndex;
    } else if (m_selectedGroup == to.groupIndex && m_selectedTab >= to.tabIndex) {
        ++m_selectedTab;
    }

    EnsureSplitIntegrity(from.groupIndex);
    EnsureSplitIntegrity(to.groupIndex);
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

void TabManager::ToggleSplitView(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    group->splitView = !group->splitView;
    if (!group->splitView) {
        group->splitPrimary = -1;
        group->splitSecondary = -1;
    } else {
        if (group->splitPrimary == -1) {
            group->splitPrimary = m_selectedGroup == groupIndex ? m_selectedTab : -1;
        }
        EnsureSplitIntegrity(groupIndex);
    }
}

void TabManager::SetSplitSecondary(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }
    auto* group = GetGroup(location.groupIndex);
    if (!group) {
        return;
    }
    if (!group->splitView) {
        group->splitView = true;
    }
    group->splitSecondary = location.tabIndex;
    EnsureSplitIntegrity(location.groupIndex);
}

void TabManager::ClearSplitSecondary(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    group->splitSecondary = -1;
    EnsureSplitIntegrity(groupIndex);
}

TabLocation TabManager::GetSplitSecondary(int groupIndex) const {
    const auto* group = GetGroup(groupIndex);
    if (!group || group->splitSecondary < 0 || group->splitSecondary >= static_cast<int>(group->tabs.size())) {
        return {};
    }
    if (group->splitSecondary >= 0 && group->splitSecondary < static_cast<int>(group->tabs.size()) &&
        !group->tabs[static_cast<size_t>(group->splitSecondary)].hidden) {
        return {groupIndex, group->splitSecondary};
    }
    return {};
}

bool TabManager::IsSplitViewEnabled(int groupIndex) const {
    const auto* group = GetGroup(groupIndex);
    return group && group->splitView;
}

void TabManager::SwapSplitSelection(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group || !group->splitView) {
        return;
    }
    if (group->splitSecondary < 0 || group->splitSecondary >= static_cast<int>(group->tabs.size())) {
        return;
    }
    if (group->tabs[static_cast<size_t>(group->splitSecondary)].hidden) {
        return;
    }

    std::swap(group->splitPrimary, group->splitSecondary);
    if (groupIndex == m_selectedGroup && group->splitPrimary >= 0) {
        m_selectedTab = group->splitPrimary;
    }
    EnsureSplitIntegrity(groupIndex);
}

void TabManager::EnsureDefaultGroup() {
    if (!m_groups.empty()) {
        return;
    }

    TabGroup group;
    group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
    group.headerVisible = true;
    m_groups.emplace_back(std::move(group));
    if (m_selectedGroup < 0) {
        m_selectedGroup = 0;
        m_selectedTab = -1;
    }
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

    EnsureSplitIntegrity(m_selectedGroup);
}

void TabManager::EnsureSplitIntegrity(int groupIndex) {
    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    auto& group = m_groups[static_cast<size_t>(groupIndex)];
    EnsureSplitIntegrity(group);
    if (group.splitView && group.splitPrimary >= 0 && groupIndex == m_selectedGroup) {
        m_selectedTab = group.splitPrimary;
    }
}

void TabManager::EnsureSplitIntegrity(TabGroup& group) {
    if (!group.splitView) {
        group.splitPrimary = -1;
        group.splitSecondary = -1;
        return;
    }

    const int tabCount = static_cast<int>(group.tabs.size());
    if (tabCount == 0) {
        group.splitView = false;
        group.splitPrimary = -1;
        group.splitSecondary = -1;
        return;
    }

    auto findVisible = [&](int skip) -> int {
        for (int i = 0; i < tabCount; ++i) {
            if (group.tabs[static_cast<size_t>(i)].hidden) {
                continue;
            }
            if (skip >= 0 && i == skip) {
                continue;
            }
            return i;
        }
        return -1;
    };

    if (group.splitPrimary < 0 || group.splitPrimary >= tabCount ||
        group.tabs[static_cast<size_t>(group.splitPrimary)].hidden) {
        group.splitPrimary = findVisible(-1);
    }

    const int visibleCount = static_cast<int>(std::count_if(group.tabs.begin(), group.tabs.end(), [](const TabInfo& tab) {
        return !tab.hidden;
    }));

    if (group.splitPrimary < 0) {
        group.splitView = false;
        group.splitSecondary = -1;
        return;
    }

    if (visibleCount < 2) {
        group.splitSecondary = -1;
        if (visibleCount < 1) {
            group.splitView = false;
            group.splitPrimary = -1;
        }
        return;
    }

    if (group.splitSecondary < 0 || group.splitSecondary >= tabCount ||
        group.tabs[static_cast<size_t>(group.splitSecondary)].hidden || group.splitSecondary == group.splitPrimary) {
        group.splitSecondary = findVisible(group.splitPrimary);
    }

    if (group.splitSecondary == group.splitPrimary) {
        group.splitSecondary = findVisible(group.splitPrimary);
    }

    if (group.splitSecondary < 0) {
        group.splitView = false;
    }
}

bool TabManager::NextVisibleTabIndex(const TabGroup& group, int* index) const {
    if (!index) {
        return false;
    }
    const int skip = *index;
    for (size_t i = 0; i < group.tabs.size(); ++i) {
        if (group.tabs[i].hidden) {
            continue;
        }
        if (static_cast<int>(i) == skip) {
            continue;
        }
        *index = static_cast<int>(i);
        return true;
    }
    return false;
}

}  // namespace shelltabs
