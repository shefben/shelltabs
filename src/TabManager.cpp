#include "TabManager.h"

#include "Logging.h"
#include "ShellTabsMessages.h"
#include "IconCache.h"

#include <algorithm>
#include <cmath>
#include <cwchar>
#include <cwctype>
#include <functional>
#include <limits>
#include <numeric>
#include <optional>
#include <memory>


namespace shelltabs {

namespace {
uint64_t HashCombine64(uint64_t seed, uint64_t value) noexcept {
    return seed ^ (value + 0x9e3779b97f4a7c15ull + (seed << 6) + (seed >> 2));
}

uint64_t HashWideString64(const std::wstring& value) noexcept {
    return std::hash<std::wstring>{}(value);
}
}  // namespace

namespace {
constexpr wchar_t kDefaultGroupNamePrefix[] = L"Island ";
}  // namespace

namespace {
inline double ClampProgress(double value) {
    if (value < 0.0) {
        return 0.0;
    }
    if (value > 1.0) {
        return 1.0;
    }
    return value;
}
}  // namespace

namespace {
void InvalidateTabIcon(const TabInfo& tab) {
    const std::wstring familyKey = BuildIconCacheFamilyKey(tab.pidl.get(), tab.path);
    if (!familyKey.empty()) {
        IconCache::Instance().InvalidateFamily(familyKey);
    }
}
}  // namespace

namespace {
int CountLeadingPinned(const TabGroup& group) {
    int count = 0;
    for (const auto& tab : group.tabs) {
        if (!tab.pinned) {
            break;
        }
        ++count;
    }
    return count;
}
}  // namespace

namespace {
std::wstring NormalizeLookupKey(const std::wstring& value) {
    if (value.empty()) {
        return {};
    }
    std::wstring normalized = NormalizeFileSystemPath(value);
    if (normalized.empty()) {
        normalized = value;
    }
    std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return normalized;
}

std::wstring BuildLookupKey(PCIDLIST_ABSOLUTE pidl, const std::wstring& storedPath) {
    if (!storedPath.empty()) {
        return NormalizeLookupKey(storedPath);
    }
    if (!pidl) {
        return {};
    }
    std::wstring canonical = GetCanonicalParsingName(pidl);
    if (canonical.empty()) {
        canonical = GetParsingName(pidl);
    }
    return NormalizeLookupKey(canonical);
}

std::wstring BuildLookupKey(const TabInfo& tab) {
    if (!tab.normalizedLookupKey.empty()) {
        return tab.normalizedLookupKey;
    }
    return BuildLookupKey(tab.pidl.get(), tab.path);
}

}  // namespace

void TabInfo::RefreshNormalizedLookupKey() {
    normalizedLookupKey = BuildLookupKey(pidl.get(), path);
}

namespace {

bool LocationLess(const TabLocation& lhs, const TabLocation& rhs) noexcept {
    if (lhs.groupIndex != rhs.groupIndex) {
        return lhs.groupIndex < rhs.groupIndex;
    }
    return lhs.tabIndex < rhs.tabIndex;
}
}  // namespace

bool TabManager::ActivationPrecedes(const ActivationEntry& lhs, const ActivationEntry& rhs) noexcept {
    if (lhs.epoch != rhs.epoch) {
        return lhs.epoch > rhs.epoch;
    }
    if (lhs.ordinal != rhs.ordinal) {
        return lhs.ordinal > rhs.ordinal;
    }
    if (lhs.tick != rhs.tick) {
        return lhs.tick > rhs.tick;
    }
    if (lhs.location.groupIndex != rhs.location.groupIndex) {
        return lhs.location.groupIndex < rhs.location.groupIndex;
    }
    return lhs.location.tabIndex < rhs.location.tabIndex;
}

uint64_t TabManager::EncodeActivationKey(TabLocation location) noexcept {
    const uint64_t groupPart = static_cast<uint64_t>(static_cast<uint32_t>(location.groupIndex));
    const uint64_t tabPart = static_cast<uint64_t>(static_cast<uint32_t>(location.tabIndex));
    return (groupPart << 32) | tabPart;
}

uint64_t ComputeTabViewStableId(const TabViewItem& item) noexcept {
    uint64_t hash = static_cast<uint64_t>(item.type);
    if (item.type == TabViewItemType::kTab) {
        uint64_t ordinal = item.activationOrdinal != 0 ? item.activationOrdinal : item.lastActivatedTick;
        if (ordinal == 0 && item.pidl) {
            ordinal = reinterpret_cast<uint64_t>(item.pidl);
        }
        hash = HashCombine64(hash, ordinal);
        hash = HashCombine64(hash, HashWideString64(item.savedGroupId));
        hash = HashCombine64(hash, HashWideString64(item.path));
        hash = HashCombine64(hash, HashWideString64(item.name));
    } else {
        const int groupIndex = item.location.groupIndex >= 0 ? item.location.groupIndex : 0;
        hash = HashCombine64(hash, static_cast<uint64_t>(groupIndex));
        hash = HashCombine64(hash, HashWideString64(item.savedGroupId));
        hash = HashCombine64(hash, HashWideString64(item.name));
    }
    if (hash == 0) {
        hash = 1;  // reserve 0 for "unknown" identifiers
    }
    return hash;
}

std::mutex TabManager::s_windowMutex;
std::unordered_map<TabManager::ExplorerWindowId, TabManager*, TabManager::ExplorerWindowIdHash>
    TabManager::s_windowMap;

TabManager::TabManager() {
    EnsureDefaultGroup();
    RebuildIndices();
}

TabManager::~TabManager() { ClearWindowId(); }

TabManager& TabManager::Get() {
    static TabManager instance;
    return instance;
}

TabManager* TabManager::Find(ExplorerWindowId id) {
    if (!id.IsValid()) {
        return nullptr;
    }
    std::scoped_lock lock(s_windowMutex);
    auto it = s_windowMap.find(id);
    if (it == s_windowMap.end()) {
        return nullptr;
    }
    return it->second;
}

void TabManager::SetWindowId(ExplorerWindowId id) {
    std::scoped_lock lock(s_windowMutex);
    if (m_windowId == id) {
        return;
    }
    if (m_windowId.IsValid()) {
        auto existing = s_windowMap.find(m_windowId);
        if (existing != s_windowMap.end() && existing->second == this) {
            s_windowMap.erase(existing);
        }
    }
    m_windowId = id;
    if (m_windowId.IsValid()) {
        s_windowMap[m_windowId] = this;
    }
}

void TabManager::ClearWindowId() {
    std::scoped_lock lock(s_windowMutex);
    if (!m_windowId.IsValid()) {
        return;
    }
    auto existing = s_windowMap.find(m_windowId);
    if (existing != s_windowMap.end() && existing->second == this) {
        s_windowMap.erase(existing);
    }
    m_windowId = {};
}

size_t TabManager::ActiveWindowCount() {
    std::scoped_lock lock(s_windowMutex);
    return s_windowMap.size();
}

int TabManager::TotalTabCount() const noexcept {
    int total = 0;
    for (const auto& group : m_groups) {
        total += static_cast<int>(group.tabs.size());
    }
    return total;
}

void TabManager::SetSelectedLocation(TabLocation location) {
    const TabLocation previous = SelectedLocation();
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
    EnsureVisibleSelection();
    UpdateSelectionActivation(previous);
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
    const std::wstring key = BuildLookupKey(pidl, std::wstring{});
    const TabLocation cached = ResolveFromIndex(key, pidl, false);
    if (cached.IsValid()) {
        return cached;
    }
    return ScanForPidl(pidl);
}

TabLocation TabManager::GetLastActivatedTab(bool includeHidden) const {
    const TabLocation current = SelectedLocation();
    TabLocation best;
    uint64_t bestOrdinal = 0;
    ULONGLONG bestTick = 0;
    bool hasBest = false;

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        int candidateIndex = includeHidden ? group.lastActivatedTabIndex : group.lastVisibleActivatedTabIndex;
        const int excludeIndex = (current.groupIndex == static_cast<int>(g)) ? current.tabIndex : -1;
        if (candidateIndex == excludeIndex) {
            candidateIndex = FindBestActivatedTabIndex(group, includeHidden, excludeIndex);
        }
        if (candidateIndex < 0) {
            continue;
        }

        if (current.groupIndex == static_cast<int>(g) && current.tabIndex == candidateIndex) {
            continue;
        }

        const auto& tab = group.tabs[static_cast<size_t>(candidateIndex)];
        if (!includeHidden && tab.hidden) {
            continue;
        }

        const uint64_t ordinal = tab.activationOrdinal;
        const ULONGLONG tick = tab.lastActivatedTick;
        if (!hasBest || ordinal > bestOrdinal || (ordinal == bestOrdinal && tick > bestTick) ||
            (ordinal == bestOrdinal && tick == bestTick &&
             (static_cast<int>(g) < best.groupIndex ||
              (static_cast<int>(g) == best.groupIndex && candidateIndex < best.tabIndex)))) {
            best = {static_cast<int>(g), candidateIndex};
            bestOrdinal = ordinal;
            bestTick = tick;
            hasBest = true;
        }
    }

    if (!hasBest) {
        return {};
    }
    return best;
}

std::vector<TabLocation> TabManager::GetTabsByActivationOrder(bool includeHidden) const {
    std::vector<TabLocation> order;
    order.reserve(m_activationOrder.size());
    for (const auto& entry : m_activationOrder) {
        const TabLocation location = entry.location;
        const TabInfo* tab = Get(location);
        if (!tab) {
            continue;
        }
        if (!includeHidden && tab->hidden) {
            continue;
        }
        order.emplace_back(location);
    }
    return order;
}

TabLocation TabManager::Add(UniquePidl pidl, std::wstring name, std::wstring tooltip, bool select, int groupIndex,
                            bool pinned) {
    if (!pidl) {
        return {};
    }

    EnsureDefaultGroup();

    const TabLocation previousSelection = SelectedLocation();

    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        groupIndex = (m_selectedGroup >= 0) ? m_selectedGroup : 0;
    }

    TabInfo info{
        .pidl = std::move(pidl),
        .name = std::move(name),
        .tooltip = std::move(tooltip),
        .hidden = false,
        .pinned = pinned,
    };
    std::wstring canonicalPath = GetCanonicalParsingName(info.pidl.get());
    if (canonicalPath.empty()) {
        canonicalPath = GetParsingName(info.pidl.get());
    }
    info.path = std::move(canonicalPath);
    info.RefreshNormalizedLookupKey();

    auto& group = m_groups[static_cast<size_t>(groupIndex)];
    const int desiredIndex = info.pinned ? CountLeadingPinned(group)
                                         : static_cast<int>(group.tabs.size());
    const int insertIndex = std::clamp(desiredIndex, 0, static_cast<int>(group.tabs.size()));
    group.tabs.insert(group.tabs.begin() + insertIndex, std::move(info));
    HandleTabInserted(group, insertIndex);

    const TabLocation location{groupIndex, insertIndex};
    if (select) {
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
        group.collapsed = false;
    } else if (m_selectedGroup == location.groupIndex && m_selectedTab >= location.tabIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();
    MarkLayoutDirty();
    IndexInsertTab(location);
    UpdateSelectionActivation(previousSelection);
    return location;
}

void TabManager::Remove(TabLocation location) {
    if (!location.IsValid()) {
        return;
    }
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }

    const TabLocation previousSelection = SelectedLocation();
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }

    const bool wasSelected = (m_selectedGroup == location.groupIndex && m_selectedTab == location.tabIndex);

    TabInfo removed = std::move(group.tabs[static_cast<size_t>(location.tabIndex)]);
    const bool removedHidden = removed.hidden;
    InvalidateTabIcon(removed);
    group.tabs.erase(group.tabs.begin() + location.tabIndex);

    HandleTabRemoved(group, location.tabIndex, removedHidden);
    IndexRemoveTab(location, removed);

    if (group.tabs.empty()) {
        m_groups.erase(m_groups.begin() + location.groupIndex);
        if (m_selectedGroup == location.groupIndex) {
            m_selectedGroup = -1;
            m_selectedTab = -1;
        } else if (m_selectedGroup > location.groupIndex) {
            --m_selectedGroup;
        }
        IndexShiftGroups(location.groupIndex + 1, -1);
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

    MarkLayoutDirty();
    EnsureVisibleSelection();
    UpdateSelectionActivation(previousSelection);
}

std::optional<TabInfo> TabManager::TakeTab(TabLocation location) {
    if (!location.IsValid()) {
        return std::nullopt;
    }
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return std::nullopt;
    }

    const TabLocation previousSelection = SelectedLocation();
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return std::nullopt;
    }

    TabInfo removed = std::move(group.tabs[static_cast<size_t>(location.tabIndex)]);
    InvalidateTabIcon(removed);

    const bool wasSelected = (m_selectedGroup == location.groupIndex && m_selectedTab == location.tabIndex);

    group.tabs.erase(group.tabs.begin() + location.tabIndex);
    HandleTabRemoved(group, location.tabIndex, removed.hidden);
    IndexRemoveTab(location, removed);

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
        IndexShiftGroups(location.groupIndex + 1, -1);
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

    MarkLayoutDirty();
    EnsureVisibleSelection();

    UpdateSelectionActivation(previousSelection);
    return removed;
}

TabLocation TabManager::InsertTab(TabInfo tab, int groupIndex, int tabIndex, bool select) {
    EnsureDefaultGroup();

    const TabLocation previousSelection = SelectedLocation();

    tab.RefreshNormalizedLookupKey();

    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        groupIndex = std::clamp(groupIndex, 0, static_cast<int>(m_groups.size()) - 1);
    }
    if (m_groups.empty()) {
        groupIndex = 0;
    }

    auto& group = m_groups[static_cast<size_t>(groupIndex)];
    int insertIndex = std::clamp(tabIndex, 0, static_cast<int>(group.tabs.size()));
    const int pinnedCount = CountLeadingPinned(group);
    if (tab.pinned) {
        insertIndex = std::clamp(insertIndex, 0, pinnedCount);
    } else {
        insertIndex = std::max(insertIndex, pinnedCount);
    }

    group.tabs.insert(group.tabs.begin() + insertIndex, std::move(tab));
    HandleTabInserted(group, insertIndex);

    if (select) {
        m_selectedGroup = groupIndex;
        m_selectedTab = insertIndex;
        group.collapsed = false;
    } else if (m_selectedGroup == groupIndex && m_selectedTab >= insertIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();

    MarkLayoutDirty();
    IndexInsertTab({groupIndex, insertIndex});
    UpdateSelectionActivation(previousSelection);
    return {groupIndex, insertIndex};
}

std::optional<TabGroup> TabManager::TakeGroup(int groupIndex) {
    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        return std::nullopt;
    }

    TabGroup removed = std::move(m_groups[static_cast<size_t>(groupIndex)]);
    const bool wasSelected = (m_selectedGroup == groupIndex);

    m_groups.erase(m_groups.begin() + groupIndex);
    IndexRemoveGroup(groupIndex, removed);

    if (wasSelected) {
        m_selectedGroup = -1;
        m_selectedTab = -1;
    } else if (m_selectedGroup > groupIndex) {
        --m_selectedGroup;
    }

    EnsureDefaultGroup();
    MarkLayoutDirty();
    EnsureVisibleSelection();

    return removed;
}

int TabManager::InsertGroup(TabGroup group, int insertIndex) {
    EnsureDefaultGroup();

    if (insertIndex < 0) {
        insertIndex = 0;
    }
    if (insertIndex > static_cast<int>(m_groups.size())) {
        insertIndex = static_cast<int>(m_groups.size());
    }

    const auto position = m_groups.begin() + insertIndex;
    m_groups.insert(position, std::move(group));
    for (auto& tab : m_groups[static_cast<size_t>(insertIndex)].tabs) {
        if (tab.normalizedLookupKey.empty()) {
            tab.RefreshNormalizedLookupKey();
        }
    }
    RefreshGroupAggregates(m_groups[static_cast<size_t>(insertIndex)]);

    if (m_selectedGroup >= insertIndex) {
        ++m_selectedGroup;
    }

    EnsureVisibleSelection();
    MarkLayoutDirty();
    IndexShiftGroups(insertIndex, 1);
    IndexInsertGroup(insertIndex);
    return insertIndex;
}

void TabManager::Clear() {
    for (const auto& group : m_groups) {
        for (const auto& tab : group.tabs) {
            InvalidateTabIcon(tab);
        }
    }
    m_groups.clear();
    m_selectedGroup = -1;
    m_selectedTab = -1;
    m_groupSequence = 1;
    m_nextActivationOrdinal = 1;
    m_activationEpoch = 0;
    m_lastActivationOrdinalSeen = 0;
    m_lastActivationTickSeen = 0;
    m_activationOrder.clear();
    m_activationLookup.clear();
    EnsureDefaultGroup();
    m_locationIndex.clear();
    MarkLayoutDirty();
}

void TabManager::Restore(std::vector<TabGroup> groups, int selectedGroup, int selectedTab, int groupSequence) {
    for (const auto& group : m_groups) {
        for (const auto& tab : group.tabs) {
            InvalidateTabIcon(tab);
        }
    }
    m_groups = std::move(groups);
    for (auto& group : m_groups) {
        NormalizePinnedOrder(group);
    }
    m_selectedGroup = selectedGroup;
    m_selectedTab = selectedTab;
    m_groupSequence = std::max(groupSequence, 1);
    EnsureDefaultGroup();
    RecalculateNextActivationOrdinal();
    EnsureVisibleSelection();
    RebuildIndices();
    MarkLayoutDirty();
}

std::vector<TabViewItem> TabManager::BuildView() const {
    std::vector<TabViewItem> items;
    items.reserve(TotalTabCount() + static_cast<int>(m_groups.size()));

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        const size_t total = group.tabs.size();
        const size_t visible = group.visibleCount;
        const size_t hidden = group.hiddenCount;
        const ULONGLONG groupLastTick = group.lastActivatedTick;
        const uint64_t groupLastOrdinal = group.lastActivationOrdinal;

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
            header.outlineStyle = group.outlineStyle;
            header.savedGroupId = group.savedGroupId;
            header.isSavedGroup = !group.savedGroupId.empty();
            header.headerVisible = group.headerVisible;
            header.lastActivatedTick = groupLastTick;
            header.activationOrdinal = groupLastOrdinal;
            header.stableId = ComputeTabViewStableId(header);
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
            item.outlineStyle = group.outlineStyle;
            item.savedGroupId = group.savedGroupId;
            item.isSavedGroup = !group.savedGroupId.empty();
            item.headerVisible = group.headerVisible;
            item.lastActivatedTick = tab.lastActivatedTick;
            item.activationOrdinal = tab.activationOrdinal;
            item.pinned = tab.pinned;
            if (tab.progress.active) {
                item.progress.visible = true;
                item.progress.indeterminate = tab.progress.indeterminate;
                item.progress.fraction = tab.progress.indeterminate ? 0.0 : ClampProgress(tab.progress.fraction);
            }

            item.stableId = ComputeTabViewStableId(item);
            items.emplace_back(std::move(item));
        }
    }

    return items;
}

TabProgressSnapshot TabManager::CollectProgressStates() const {
    TabProgressSnapshot snapshot;
    snapshot.reserve(TotalTabCount() + static_cast<int>(m_groups.size()));

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        const ULONGLONG groupLastTick = group.lastActivatedTick;
        const uint64_t groupLastOrdinal = group.lastActivationOrdinal;

        if (group.headerVisible) {
            TabProgressSnapshotEntry header;
            header.type = TabViewItemType::kGroupHeader;
            header.location = {static_cast<int>(g), -1};
            header.lastActivatedTick = groupLastTick;
            header.activationOrdinal = groupLastOrdinal;
            snapshot.emplace_back(std::move(header));
        }

        if (group.collapsed) {
            continue;
        }

        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            if (tab.hidden) {
                continue;
            }

            TabProgressSnapshotEntry entry;
            entry.type = TabViewItemType::kTab;
            entry.location = {static_cast<int>(g), static_cast<int>(t)};
            entry.lastActivatedTick = tab.lastActivatedTick;
            entry.activationOrdinal = tab.activationOrdinal;
            if (tab.progress.active) {
                entry.progress.visible = true;
                entry.progress.indeterminate = tab.progress.indeterminate;
                entry.progress.fraction = tab.progress.indeterminate ? 0.0 : ClampProgress(tab.progress.fraction);
            }
            snapshot.emplace_back(std::move(entry));
        }
    }

    return snapshot;
}

void TabManager::RegisterProgressListener(HWND hwnd) {
    if (!hwnd) {
        return;
    }
    if (std::find(m_progressListeners.begin(), m_progressListeners.end(), hwnd) != m_progressListeners.end()) {
        return;
    }
    m_progressListeners.push_back(hwnd);
}

void TabManager::UnregisterProgressListener(HWND hwnd) {
    if (!hwnd) {
        return;
    }
    auto it = std::remove(m_progressListeners.begin(), m_progressListeners.end(), hwnd);
    if (it != m_progressListeners.end()) {
        m_progressListeners.erase(it, m_progressListeners.end());
    }
}

void TabManager::NotifyProgressListeners() {
    if (m_pendingProgressUpdates.empty() && m_progressListeners.empty()) {
        return;
    }

    std::vector<TabProgressSnapshotEntry> updates;
    updates.reserve(m_pendingProgressUpdates.size());
    for (const auto& key : m_pendingProgressUpdates) {
        TabProgressSnapshotEntry entry;
        if (BuildProgressEntry(key, &entry)) {
            updates.emplace_back(std::move(entry));
        }
    }
    m_pendingProgressUpdates.clear();

#if defined(SHELLTABS_BUILD_TESTS)
    m_lastProgressUpdatesForTest = updates;
    m_lastProgressLayoutVersionForTest = m_layoutVersion;
#endif

    if (updates.empty() || m_progressListeners.empty()) {
        return;
    }

    const UINT message = GetProgressUpdateMessage();
    if (message == 0) {
        return;
    }

    auto it = m_progressListeners.begin();
    while (it != m_progressListeners.end()) {
        HWND hwnd = *it;
        if (!hwnd || !IsWindow(hwnd)) {
            it = m_progressListeners.erase(it);
            continue;
        }

        auto payload = std::make_unique<TabProgressUpdatePayload>();
        payload->layoutVersion = m_layoutVersion;
        payload->entries = updates;
        const WPARAM count = static_cast<WPARAM>(payload->entries.size());
        auto* rawPayload = payload.release();
        if (!PostMessageW(hwnd, message, count, reinterpret_cast<LPARAM>(rawPayload))) {
            delete rawPayload;
            it = m_progressListeners.erase(it);
            continue;
        }
        ++it;
    }
}

void TabManager::QueueProgressUpdate(TabViewItemType type, TabLocation location) {
    if (type == TabViewItemType::kTab) {
        if (!location.IsValid()) {
            return;
        }
    } else {
        if (location.groupIndex < 0) {
            return;
        }
        location.tabIndex = -1;
    }

    auto it = std::find_if(m_pendingProgressUpdates.begin(), m_pendingProgressUpdates.end(),
                           [&](const ProgressUpdateKey& existing) {
                               return existing.type == type &&
                                      existing.location.groupIndex == location.groupIndex &&
                                      existing.location.tabIndex == location.tabIndex;
                           });
    if (it == m_pendingProgressUpdates.end()) {
        ProgressUpdateKey key;
        key.type = type;
        key.location = location;
        m_pendingProgressUpdates.push_back(key);
    }
}

bool TabManager::BuildProgressEntry(const ProgressUpdateKey& key, TabProgressSnapshotEntry* entry) const {
    if (!entry) {
        return false;
    }

    entry->type = key.type;
    entry->location = key.location;
    entry->progress = {};
    entry->lastActivatedTick = 0;
    entry->activationOrdinal = 0;

    if (key.type == TabViewItemType::kGroupHeader) {
        const auto* group = GetGroup(key.location.groupIndex);
        if (!group || !group->headerVisible) {
            return false;
        }
        entry->lastActivatedTick = group->lastActivatedTick;
        entry->activationOrdinal = group->lastActivationOrdinal;
        return true;
    }

    const auto* tab = Get(key.location);
    if (!tab || tab->hidden) {
        return false;
    }
    entry->lastActivatedTick = tab->lastActivatedTick;
    entry->activationOrdinal = tab->activationOrdinal;
    if (tab->progress.active) {
        entry->progress.visible = true;
        entry->progress.indeterminate = tab->progress.indeterminate;
        entry->progress.fraction =
            tab->progress.indeterminate ? 0.0 : ClampProgress(tab->progress.fraction);
    }
    return true;
}

void TabManager::MarkLayoutDirty() noexcept {
    ++m_layoutVersion;
    if (m_layoutVersion == 0) {
        ++m_layoutVersion;
    }
}

TabLocation TabManager::FindByPath(const std::wstring& path) const {
    if (path.empty()) {
        return {};
    }
    const std::wstring key = BuildLookupKey(nullptr, path);
    const TabLocation cached = ResolveFromIndex(key, nullptr, true);
    if (cached.IsValid()) {
        return cached;
    }
    return ScanForPath(path);
}

TabLocation TabManager::ResolveFromIndex(const std::wstring& key, PCIDLIST_ABSOLUTE pidl, bool requireVisible) const {
    if (key.empty()) {
        return {};
    }
    const auto it = m_locationIndex.find(key);
    if (it == m_locationIndex.end()) {
        return {};
    }

    TabLocation candidate;
    const TabInfo* candidateTab = nullptr;
    bool canonicalMismatch = false;
    bool ambiguous = false;

    for (const auto& location : it->second) {
        const TabInfo* tab = Get(location);
        if (!tab) {
            continue;
        }
        if (requireVisible && tab->hidden) {
            continue;
        }
        if (pidl) {
            if (!ArePidlsCanonicallyEqual(tab->pidl.get(), pidl)) {
                continue;
            }
            if (ArePidlsEqual(tab->pidl.get(), pidl)) {
                return location;
            }
            if (!candidate.IsValid()) {
                candidate = location;
                candidateTab = tab;
                canonicalMismatch = true;
            } else {
                ambiguous = true;
            }
        } else {
            if (!candidate.IsValid()) {
                candidate = location;
                candidateTab = tab;
            } else {
                ambiguous = true;
            }
        }
    }

    if (candidate.IsValid() && !ambiguous) {
        if (pidl && canonicalMismatch && candidateTab) {
            std::wstring storedCanonical = candidateTab->path;
            if (storedCanonical.empty() && candidateTab->pidl) {
                storedCanonical = GetCanonicalParsingName(candidateTab->pidl.get());
                if (storedCanonical.empty()) {
                    storedCanonical = GetParsingName(candidateTab->pidl.get());
                }
            }
            std::wstring incomingCanonical = GetCanonicalParsingName(pidl);
            if (incomingCanonical.empty()) {
                incomingCanonical = GetParsingName(pidl);
            }
            LogMessage(LogLevel::Warning,
                       L"Canonical PIDL collision detected during navigation (group=%d, tab=%d, stored=%ls, incoming=%ls)",
                       candidate.groupIndex, candidate.tabIndex,
                       storedCanonical.empty() ? L"(unknown)" : storedCanonical.c_str(),
                       incomingCanonical.empty() ? L"(unknown)" : incomingCanonical.c_str());
        }
        return candidate;
    }

    return {};
}

TabLocation TabManager::ScanForPidl(PCIDLIST_ABSOLUTE pidl) const {
    if (!pidl) {
        return {};
    }
    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            if (ArePidlsCanonicallyEqual(tab.pidl.get(), pidl)) {
                if (!ArePidlsEqual(tab.pidl.get(), pidl)) {
                    std::wstring storedCanonical = tab.path;
                    if (storedCanonical.empty() && tab.pidl) {
                        storedCanonical = GetCanonicalParsingName(tab.pidl.get());
                        if (storedCanonical.empty()) {
                            storedCanonical = GetParsingName(tab.pidl.get());
                        }
                    }
                    std::wstring incomingCanonical = GetCanonicalParsingName(pidl);
                    if (incomingCanonical.empty()) {
                        incomingCanonical = GetParsingName(pidl);
                    }
                    LogMessage(LogLevel::Warning,
                               L"Canonical PIDL collision detected during navigation (group=%zu, tab=%zu, stored=%ls, incoming=%ls)",
                               g, t,
                               storedCanonical.empty() ? L"(unknown)" : storedCanonical.c_str(),
                               incomingCanonical.empty() ? L"(unknown)" : incomingCanonical.c_str());
                }
                return {static_cast<int>(g), static_cast<int>(t)};
            }
            if (ArePidlsEqual(tab.pidl.get(), pidl)) {
                return {static_cast<int>(g), static_cast<int>(t)};
            }
        }
    }
    return {};
}

TabLocation TabManager::ScanForPath(const std::wstring& path) const {
    if (path.empty()) {
        return {};
    }
    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            if (tab.hidden) {
                continue;
            }
            std::wstring candidate = tab.path;
            if (candidate.empty() && tab.pidl) {
                candidate = GetParsingName(tab.pidl.get());
            }
            if (!candidate.empty() && _wcsicmp(candidate.c_str(), path.c_str()) == 0) {
                return {static_cast<int>(g), static_cast<int>(t)};
            }
        }
    }
    return {};
}

void TabManager::RebuildIndices() {
    m_locationIndex.clear();
    for (size_t g = 0; g < m_groups.size(); ++g) {
        auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            auto& tab = group.tabs[t];
            if (tab.normalizedLookupKey.empty()) {
                tab.RefreshNormalizedLookupKey();
            }
            const std::wstring& key = tab.normalizedLookupKey;
            if (key.empty()) {
                continue;
            }
            m_locationIndex[key].push_back({static_cast<int>(g), static_cast<int>(t)});
        }
    }
    RebuildActivationOrder();
}

void TabManager::IndexShiftTabs(int groupIndex, int startTabIndex, int delta) {
    if (delta == 0) {
        return;
    }
    for (auto& [key, locations] : m_locationIndex) {
        for (auto& location : locations) {
            if (location.groupIndex == groupIndex && location.tabIndex >= startTabIndex) {
                location.tabIndex += delta;
            }
        }
    }
    ActivationShiftTabs(groupIndex, startTabIndex, delta);
}

void TabManager::IndexShiftGroups(int startGroupIndex, int delta) {
    if (delta == 0) {
        return;
    }
    for (auto& [key, locations] : m_locationIndex) {
        bool touched = false;
        for (auto& location : locations) {
            if (location.groupIndex >= startGroupIndex) {
                location.groupIndex += delta;
                touched = true;
            }
        }
        if (touched) {
            std::sort(locations.begin(), locations.end(), LocationLess);
        }
    }
    ActivationShiftGroups(startGroupIndex, delta);
}

void TabManager::IndexInsertTab(TabLocation location) {
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }
    IndexShiftTabs(location.groupIndex, location.tabIndex, 1);
    ActivationInsertTab(location);
    const std::wstring key = BuildLookupKey(group.tabs[static_cast<size_t>(location.tabIndex)]);
    if (key.empty()) {
        return;
    }
    auto& bucket = m_locationIndex[key];
    const auto position = std::lower_bound(bucket.begin(), bucket.end(), location, LocationLess);
    bucket.insert(position, location);
}

void TabManager::IndexRemoveTab(TabLocation location, const TabInfo& tab) {
    if (location.groupIndex < 0) {
        return;
    }
    const std::wstring key = BuildLookupKey(tab);
    if (!key.empty()) {
        auto it = m_locationIndex.find(key);
        if (it != m_locationIndex.end()) {
            auto& bucket = it->second;
            auto pos = std::find_if(bucket.begin(), bucket.end(), [&](const TabLocation& candidate) {
                return candidate.groupIndex == location.groupIndex && candidate.tabIndex == location.tabIndex;
            });
            if (pos != bucket.end()) {
                bucket.erase(pos);
                if (bucket.empty()) {
                    m_locationIndex.erase(it);
                }
            }
        }
    }
    ActivationRemoveTab(location);
    IndexShiftTabs(location.groupIndex, location.tabIndex + 1, -1);
}

void TabManager::IndexInsertGroup(int groupIndex) {
    if (groupIndex < 0 || groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    const auto& group = m_groups[static_cast<size_t>(groupIndex)];
    for (int i = 0; i < static_cast<int>(group.tabs.size()); ++i) {
        IndexInsertTab({groupIndex, i});
    }
}

void TabManager::IndexRemoveGroup(int groupIndex, const TabGroup& group) {
    if (groupIndex < 0) {
        return;
    }
    for (int i = static_cast<int>(group.tabs.size()) - 1; i >= 0; --i) {
        IndexRemoveTab({groupIndex, i}, group.tabs[static_cast<size_t>(i)]);
    }
    IndexShiftGroups(groupIndex + 1, -1);
}

void TabManager::RebuildActivationOrder() {
    m_activationOrder.clear();
    m_activationLookup.clear();
    m_activationEpoch = 0;
    m_lastActivationOrdinalSeen = 0;
    m_lastActivationTickSeen = 0;

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const TabLocation location{static_cast<int>(g), static_cast<int>(t)};
            ActivationInsertTab(location);
        }
    }
}

void TabManager::ActivationInsertTab(TabLocation location) {
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }

    TabInfo& tab = group.tabs[static_cast<size_t>(location.tabIndex)];
    const uint64_t key = EncodeActivationKey(location);
    auto existing = m_activationLookup.find(key);
    if (existing != m_activationLookup.end()) {
        m_activationOrder.erase(existing->second);
        m_activationLookup.erase(existing);
    }

    ActivationEntry entry;
    entry.location = location;
    entry.ordinal = tab.activationOrdinal;
    entry.tick = tab.lastActivatedTick;
    entry.epoch = tab.activationEpoch;

    const auto insertPos = std::find_if(m_activationOrder.begin(), m_activationOrder.end(), [&](const ActivationEntry& candidate) {
        return ActivationPrecedes(entry, candidate);
    });
    auto it = m_activationOrder.insert(insertPos, std::move(entry));
    m_activationLookup[key] = it;
}

void TabManager::ActivationRemoveTab(TabLocation location) {
    if (location.groupIndex < 0 || location.tabIndex < 0) {
        return;
    }
    const uint64_t key = EncodeActivationKey(location);
    auto it = m_activationLookup.find(key);
    if (it == m_activationLookup.end()) {
        return;
    }
    m_activationOrder.erase(it->second);
    m_activationLookup.erase(it);
}

void TabManager::ActivationShiftTabs(int groupIndex, int startTabIndex, int delta) {
    if (delta == 0 || groupIndex < 0) {
        return;
    }
    for (auto it = m_activationOrder.begin(); it != m_activationOrder.end(); ++it) {
        auto& entry = *it;
        if (entry.location.groupIndex == groupIndex && entry.location.tabIndex >= startTabIndex) {
            const uint64_t oldKey = EncodeActivationKey(entry.location);
            auto lookupIt = m_activationLookup.find(oldKey);
            if (lookupIt != m_activationLookup.end()) {
                m_activationLookup.erase(lookupIt);
            }
            entry.location.tabIndex += delta;
            m_activationLookup[EncodeActivationKey(entry.location)] = it;
        }
    }
}

void TabManager::ActivationShiftGroups(int startGroupIndex, int delta) {
    if (delta == 0) {
        return;
    }
    for (auto it = m_activationOrder.begin(); it != m_activationOrder.end(); ++it) {
        auto& entry = *it;
        if (entry.location.groupIndex >= startGroupIndex) {
            const uint64_t oldKey = EncodeActivationKey(entry.location);
            auto lookupIt = m_activationLookup.find(oldKey);
            if (lookupIt != m_activationLookup.end()) {
                m_activationLookup.erase(lookupIt);
            }
            entry.location.groupIndex += delta;
            m_activationLookup[EncodeActivationKey(entry.location)] = it;
        }
    }
}

void TabManager::ActivationUpdateTab(TabLocation location) {
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return;
    }
    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }

    TabInfo& tab = group.tabs[static_cast<size_t>(location.tabIndex)];
    const uint64_t key = EncodeActivationKey(location);
    ActivationEntry entry;
    auto existing = m_activationLookup.find(key);
    if (existing != m_activationLookup.end()) {
        entry = *existing->second;
        m_activationOrder.erase(existing->second);
        m_activationLookup.erase(existing);
    } else {
        entry.location = location;
    }

    if (tab.activationOrdinal < m_lastActivationOrdinalSeen ||
        (tab.activationOrdinal == m_lastActivationOrdinalSeen && tab.lastActivatedTick <= m_lastActivationTickSeen)) {
        ++m_activationEpoch;
    }

    m_lastActivationOrdinalSeen = tab.activationOrdinal;
    m_lastActivationTickSeen = tab.lastActivatedTick;

    entry.location = location;
    entry.ordinal = tab.activationOrdinal;
    entry.tick = tab.lastActivatedTick;
    entry.epoch = m_activationEpoch;
    tab.activationEpoch = entry.epoch;

    const auto insertPos = std::find_if(m_activationOrder.begin(), m_activationOrder.end(), [&](const ActivationEntry& candidate) {
        return ActivationPrecedes(entry, candidate);
    });
    auto it = m_activationOrder.insert(insertPos, std::move(entry));
    m_activationLookup[key] = it;
}

bool TabManager::ApplyProgress(TabLocation location, TabInfo* tab, std::optional<double> fraction,
                               ULONGLONG now) {
    if (!tab) {
        return false;
    }

    TabProgressState& state = tab->progress;
    bool changed = false;
    if (fraction.has_value()) {
        const double value = ClampProgress(*fraction);
        if (!state.active || state.indeterminate || std::abs(state.fraction - value) > 1e-4) {
            state.active = value < 1.0;
            state.indeterminate = false;
            state.fraction = value;
            changed = true;
        }
        if (value >= 1.0) {
            state = {};
            changed = true;
        }
    } else {
        if (!state.active || !state.indeterminate) {
            state.active = true;
            state.indeterminate = true;
            state.fraction = 0.0;
            changed = true;
        }
    }
    state.lastUpdateTick = now;
    if (changed) {
        QueueProgressUpdate(TabViewItemType::kTab, location);
    }
    return changed;
}

bool TabManager::ClearProgress(TabLocation location, TabInfo* tab) {
    if (!tab) {
        return false;
    }
    if (!tab->progress.active) {
        return false;
    }
    tab->progress = {};
    QueueProgressUpdate(TabViewItemType::kTab, location);
    return true;
}

bool TabManager::IsBetterActivation(uint64_t candidateOrdinal, ULONGLONG candidateTick, int candidateIndex,
                                   uint64_t bestOrdinal, ULONGLONG bestTick, int bestIndex) noexcept {
    if (candidateIndex < 0) {
        return false;
    }
    if (bestIndex < 0) {
        return true;
    }
    if (candidateOrdinal > bestOrdinal) {
        return true;
    }
    if (candidateOrdinal < bestOrdinal) {
        return false;
    }
    if (candidateTick > bestTick) {
        return true;
    }
    if (candidateTick < bestTick) {
        return false;
    }
    return candidateIndex < bestIndex;
}

void TabManager::ResetGroupAggregates(TabGroup& group) noexcept {
    group.visibleCount = 0;
    group.hiddenCount = 0;
    group.lastActivatedTabIndex = -1;
    group.lastActivationOrdinal = 0;
    group.lastActivatedTick = 0;
    group.lastVisibleActivatedTabIndex = -1;
    group.lastVisibleActivationOrdinal = 0;
    group.lastVisibleActivatedTick = 0;
}

void TabManager::AccumulateGroupAggregates(TabGroup& group, const TabInfo& tab, int tabIndex) noexcept {
    if (tab.hidden) {
        ++group.hiddenCount;
    } else {
        ++group.visibleCount;
        if (IsBetterActivation(tab.activationOrdinal, tab.lastActivatedTick, tabIndex,
                               group.lastVisibleActivationOrdinal, group.lastVisibleActivatedTick,
                               group.lastVisibleActivatedTabIndex)) {
            group.lastVisibleActivationOrdinal = tab.activationOrdinal;
            group.lastVisibleActivatedTick = tab.lastActivatedTick;
            group.lastVisibleActivatedTabIndex = tabIndex;
        }
    }

    if (IsBetterActivation(tab.activationOrdinal, tab.lastActivatedTick, tabIndex, group.lastActivationOrdinal,
                           group.lastActivatedTick, group.lastActivatedTabIndex)) {
        group.lastActivationOrdinal = tab.activationOrdinal;
        group.lastActivatedTick = tab.lastActivatedTick;
        group.lastActivatedTabIndex = tabIndex;
    }
}

void TabManager::RefreshGroupAggregates(TabGroup& group) noexcept {
    ResetGroupAggregates(group);
    for (size_t i = 0; i < group.tabs.size(); ++i) {
        AccumulateGroupAggregates(group, group.tabs[i], static_cast<int>(i));
    }
}

void TabManager::NormalizePinnedOrder(TabGroup& group) {
    for (auto& tab : group.tabs) {
        tab.RefreshNormalizedLookupKey();
    }
    std::stable_partition(group.tabs.begin(), group.tabs.end(), [](const TabInfo& tab) { return tab.pinned; });
    RefreshGroupAggregates(group);
}

void TabManager::HandleTabInserted(TabGroup& group, int tabIndex) {
    if (group.lastActivatedTabIndex >= tabIndex && group.lastActivatedTabIndex != -1) {
        ++group.lastActivatedTabIndex;
    }
    if (group.lastVisibleActivatedTabIndex >= tabIndex && group.lastVisibleActivatedTabIndex != -1) {
        ++group.lastVisibleActivatedTabIndex;
    }

    if (tabIndex < 0 || tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }

    AccumulateGroupAggregates(group, group.tabs[static_cast<size_t>(tabIndex)], tabIndex);
}

void TabManager::HandleTabRemoved(TabGroup& group, int tabIndex, bool wasHidden) {
    if (wasHidden) {
        if (group.hiddenCount > 0) {
            --group.hiddenCount;
        }
    } else {
        if (group.visibleCount > 0) {
            --group.visibleCount;
        }
    }

    bool requiresRefresh = false;
    if (group.lastActivatedTabIndex == tabIndex) {
        requiresRefresh = true;
    }
    if (!wasHidden && group.lastVisibleActivatedTabIndex == tabIndex) {
        requiresRefresh = true;
    }

    if (requiresRefresh) {
        RefreshGroupAggregates(group);
        return;
    }

    if (group.lastActivatedTabIndex > tabIndex) {
        --group.lastActivatedTabIndex;
    }
    if (group.lastVisibleActivatedTabIndex > tabIndex) {
        --group.lastVisibleActivatedTabIndex;
    }
}

void TabManager::HandleTabVisibilityChanged(TabGroup& group, int tabIndex, bool wasHidden, bool isHidden) {
    if (tabIndex < 0 || tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }
    if (wasHidden == isHidden) {
        return;
    }

    const TabInfo& tab = group.tabs[static_cast<size_t>(tabIndex)];
    if (!wasHidden && isHidden) {
        if (group.visibleCount > 0) {
            --group.visibleCount;
        }
        ++group.hiddenCount;
        if (group.lastVisibleActivatedTabIndex == tabIndex) {
            RefreshGroupAggregates(group);
            return;
        }
    } else if (wasHidden && !isHidden) {
        if (group.hiddenCount > 0) {
            --group.hiddenCount;
        }
        ++group.visibleCount;
        if (IsBetterActivation(tab.activationOrdinal, tab.lastActivatedTick, tabIndex,
                               group.lastVisibleActivationOrdinal, group.lastVisibleActivatedTick,
                               group.lastVisibleActivatedTabIndex)) {
            group.lastVisibleActivationOrdinal = tab.activationOrdinal;
            group.lastVisibleActivatedTick = tab.lastActivatedTick;
            group.lastVisibleActivatedTabIndex = tabIndex;
        }
    }
}

void TabManager::HandleTabActivationUpdated(TabGroup& group, int tabIndex) {
    if (tabIndex < 0 || tabIndex >= static_cast<int>(group.tabs.size())) {
        return;
    }
    const TabInfo& tab = group.tabs[static_cast<size_t>(tabIndex)];

    if (IsBetterActivation(tab.activationOrdinal, tab.lastActivatedTick, tabIndex, group.lastActivationOrdinal,
                           group.lastActivatedTick, group.lastActivatedTabIndex)) {
        group.lastActivationOrdinal = tab.activationOrdinal;
        group.lastActivatedTick = tab.lastActivatedTick;
        group.lastActivatedTabIndex = tabIndex;
    } else if (group.lastActivatedTabIndex == tabIndex) {
        group.lastActivationOrdinal = tab.activationOrdinal;
        group.lastActivatedTick = tab.lastActivatedTick;
    }

    if (tab.hidden) {
        if (group.lastVisibleActivatedTabIndex == tabIndex) {
            RefreshGroupAggregates(group);
        }
        return;
    }

    if (IsBetterActivation(tab.activationOrdinal, tab.lastActivatedTick, tabIndex,
                           group.lastVisibleActivationOrdinal, group.lastVisibleActivatedTick,
                           group.lastVisibleActivatedTabIndex)) {
        group.lastVisibleActivationOrdinal = tab.activationOrdinal;
        group.lastVisibleActivatedTick = tab.lastActivatedTick;
        group.lastVisibleActivatedTabIndex = tabIndex;
    } else if (group.lastVisibleActivatedTabIndex == tabIndex) {
        group.lastVisibleActivationOrdinal = tab.activationOrdinal;
        group.lastVisibleActivatedTick = tab.lastActivatedTick;
    }
}

int TabManager::FindBestActivatedTabIndex(const TabGroup& group, bool includeHidden, int excludeTabIndex) const {
    int bestIndex = -1;
    uint64_t bestOrdinal = 0;
    ULONGLONG bestTick = 0;
    for (size_t i = 0; i < group.tabs.size(); ++i) {
        const int index = static_cast<int>(i);
        if (index == excludeTabIndex) {
            continue;
        }
        const auto& tab = group.tabs[i];
        if (!includeHidden && tab.hidden) {
            continue;
        }
        if (IsBetterActivation(tab.activationOrdinal, tab.lastActivatedTick, index, bestOrdinal, bestTick,
                               bestIndex)) {
            bestOrdinal = tab.activationOrdinal;
            bestTick = tab.lastActivatedTick;
            bestIndex = index;
        }
    }
    return bestIndex;
}

void TabManager::UpdateSelectionActivation(TabLocation previousSelection) {
    const TabLocation current = SelectedLocation();
    if (!current.IsValid()) {
        return;
    }
    if (previousSelection.IsValid() && previousSelection.groupIndex == current.groupIndex &&
        previousSelection.tabIndex == current.tabIndex) {
        return;
    }

    TabInfo* tab = Get(current);
    if (!tab) {
        return;
    }

    tab->lastActivatedTick = GetTickCount64();
    if (m_nextActivationOrdinal == 0) {
        m_nextActivationOrdinal = 1;
    }
    tab->activationOrdinal = m_nextActivationOrdinal;
    if (m_nextActivationOrdinal < std::numeric_limits<uint64_t>::max()) {
        ++m_nextActivationOrdinal;
    }

    ActivationUpdateTab(current);

    TabGroup* group = GetGroup(current.groupIndex);
    if (group) {
        HandleTabActivationUpdated(*group, current.tabIndex);
        QueueProgressUpdate(TabViewItemType::kGroupHeader, {current.groupIndex, -1});
    }

    QueueProgressUpdate(TabViewItemType::kTab, current);
    NotifyProgressListeners();
}

void TabManager::RecalculateNextActivationOrdinal() {
    uint64_t maxOrdinal = 0;
    for (const auto& group : m_groups) {
        if (group.lastActivationOrdinal > maxOrdinal) {
            maxOrdinal = group.lastActivationOrdinal;
        }
    }

    if (maxOrdinal >= std::numeric_limits<uint64_t>::max()) {
        m_nextActivationOrdinal = std::numeric_limits<uint64_t>::max();
    } else {
        m_nextActivationOrdinal = maxOrdinal + 1;
    }

    if (m_nextActivationOrdinal == 0) {
        m_nextActivationOrdinal = 1;
    }
}

void TabManager::TouchFolderOperation(PCIDLIST_ABSOLUTE folder, std::optional<double> fraction) {
    ULONGLONG now = GetTickCount64();
    TabLocation location = Find(folder);
    if (!location.IsValid() && folder) {
        std::wstring path = GetParsingName(folder);
        location = FindByPath(path);
    }
    if (!location.IsValid()) {
        return;
    }
    TabInfo* tab = Get(location);
    if (!tab) {
        return;
    }
    if (ApplyProgress(location, tab, fraction, now)) {
        NotifyProgressListeners();
    }
}

void TabManager::ClearFolderOperation(PCIDLIST_ABSOLUTE folder) {
    TabLocation location = Find(folder);
    if (!location.IsValid() && folder) {
        std::wstring path = GetParsingName(folder);
        location = FindByPath(path);
    }
    if (!location.IsValid()) {
        return;
    }
    TabInfo* tab = Get(location);
    if (!tab) {
        return;
    }
    if (ClearProgress(location, tab)) {
        NotifyProgressListeners();
    }
}

std::vector<TabLocation> TabManager::ExpireFolderOperations(ULONGLONG now, ULONGLONG timeoutMs) {
    std::vector<TabLocation> expired;
    for (size_t g = 0; g < m_groups.size(); ++g) {
        auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            auto& tab = group.tabs[t];
            if (!tab.progress.active) {
                continue;
            }
            if (now >= tab.progress.lastUpdateTick && (now - tab.progress.lastUpdateTick) > timeoutMs) {
                tab.progress = {};
                expired.push_back({static_cast<int>(g), static_cast<int>(t)});
                QueueProgressUpdate(TabViewItemType::kTab, {static_cast<int>(g), static_cast<int>(t)});
            }
        }
    }
    if (!expired.empty()) {
        NotifyProgressListeners();
    }
    return expired;
}

bool TabManager::HasActiveProgress() const {
    for (const auto& group : m_groups) {
        for (const auto& tab : group.tabs) {
            if (tab.progress.active) {
                return true;
            }
        }
    }
    return false;
}

void TabManager::ToggleGroupCollapsed(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    group->collapsed = !group->collapsed;
    MarkLayoutDirty();
    if (!group->collapsed) {
        EnsureVisibleSelection();
    }
}

void TabManager::SetGroupCollapsed(int groupIndex, bool collapsed) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    if (group->collapsed == collapsed) {
        return;
    }
    group->collapsed = collapsed;
    MarkLayoutDirty();
    if (!collapsed) {
        EnsureVisibleSelection();
    }
}

void TabManager::HideTab(TabLocation location) {
    auto* tab = Get(location);
    if (!tab) {
        return;
    }
    TabGroup* group = GetGroup(location.groupIndex);
    if (!group) {
        return;
    }
    const TabLocation previousSelection = SelectedLocation();
    const bool wasHidden = tab->hidden;
    tab->hidden = true;
    HandleTabVisibilityChanged(*group, location.tabIndex, wasHidden, true);
    if (!wasHidden) {
        MarkLayoutDirty();
    }
    if (m_selectedGroup == location.groupIndex && m_selectedTab == location.tabIndex) {
        EnsureVisibleSelection();
        UpdateSelectionActivation(previousSelection);
    }
}

void TabManager::UnhideTab(TabLocation location) {
    auto* tab = Get(location);
    if (!tab) {
        return;
    }
    TabGroup* group = GetGroup(location.groupIndex);
    if (!group) {
        return;
    }
    const TabLocation previousSelection = SelectedLocation();
    const bool wasHidden = tab->hidden;
    tab->hidden = false;
    HandleTabVisibilityChanged(*group, location.tabIndex, wasHidden, false);
    if (wasHidden) {
        MarkLayoutDirty();
    }
    if (m_selectedGroup < 0 || m_selectedGroup >= static_cast<int>(m_groups.size())) {
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
    }
    EnsureVisibleSelection();
    UpdateSelectionActivation(previousSelection);
}

void TabManager::UnhideAllInGroup(int groupIndex) {
    auto* group = GetGroup(groupIndex);
    if (!group) {
        return;
    }
    const TabLocation previousSelection = SelectedLocation();
    bool anyChanged = false;
    for (auto& tab : group->tabs) {
        if (tab.hidden) {
            anyChanged = true;
        }
        tab.hidden = false;
    }
    RefreshGroupAggregates(*group);
    if (anyChanged) {
        MarkLayoutDirty();
    }
    if (m_selectedGroup < 0 || m_selectedGroup >= static_cast<int>(m_groups.size())) {
        m_selectedGroup = groupIndex;
        m_selectedTab = group->tabs.empty() ? -1 : 0;
    }
    EnsureVisibleSelection();
    UpdateSelectionActivation(previousSelection);
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
    return group->hiddenCount;
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
    ResetGroupAggregates(group);

    const int insertIndex = groupIndex + 1;
    const auto position = m_groups.begin() + std::clamp(insertIndex, 0, static_cast<int>(m_groups.size()));
    m_groups.insert(position, std::move(group));

    if (m_selectedGroup >= insertIndex) {
        ++m_selectedGroup;
    }

    EnsureDefaultGroup();
    MarkLayoutDirty();
    EnsureVisibleSelection();
    IndexShiftGroups(insertIndex, 1);

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

    TabInfo movingTab = std::move(sourceGroup.tabs[static_cast<size_t>(from.tabIndex)]);
    const bool wasSelected = (m_selectedGroup == from.groupIndex && m_selectedTab == from.tabIndex);

    sourceGroup.tabs.erase(sourceGroup.tabs.begin() + from.tabIndex);
    HandleTabRemoved(sourceGroup, from.tabIndex, movingTab.hidden);
    IndexRemoveTab(from, movingTab);

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
        IndexShiftGroups(from.groupIndex + 1, -1);
    }

    if (removedSourceGroup && to.groupIndex > from.groupIndex) {
        --to.groupIndex;
    }

    EnsureDefaultGroup();

    to.groupIndex = std::clamp(to.groupIndex, 0, static_cast<int>(m_groups.size()) - 1);
    auto& destinationGroup = m_groups[static_cast<size_t>(to.groupIndex)];

    const int destinationSize = static_cast<int>(destinationGroup.tabs.size());
    if (to.tabIndex < 0 || to.tabIndex > destinationSize) {
        to.tabIndex = destinationSize;
    }

    const int pinnedCount = CountLeadingPinned(destinationGroup);
    if (movingTab.pinned) {
        to.tabIndex = std::clamp(to.tabIndex, 0, pinnedCount);
    } else {
        const int lowerBound = std::min(pinnedCount, destinationSize);
        to.tabIndex = std::max(to.tabIndex, lowerBound);
        to.tabIndex = std::clamp(to.tabIndex, lowerBound, destinationSize);
    }

    destinationGroup.tabs.insert(destinationGroup.tabs.begin() + to.tabIndex, std::move(movingTab));
    HandleTabInserted(destinationGroup, to.tabIndex);
    IndexInsertTab({to.groupIndex, to.tabIndex});

    if (wasSelected) {
        m_selectedGroup = to.groupIndex;
        m_selectedTab = to.tabIndex;
    } else if (m_selectedGroup == to.groupIndex && m_selectedTab >= to.tabIndex) {
        ++m_selectedTab;
    }

    MarkLayoutDirty();
    EnsureVisibleSelection();
}

bool TabManager::SetTabPinned(TabLocation location, bool pinned) {
    if (!location.IsValid()) {
        return false;
    }
    if (location.groupIndex < 0 || location.groupIndex >= static_cast<int>(m_groups.size())) {
        return false;
    }

    auto& group = m_groups[static_cast<size_t>(location.groupIndex)];
    if (location.tabIndex < 0 || location.tabIndex >= static_cast<int>(group.tabs.size())) {
        return false;
    }

    if (group.tabs[static_cast<size_t>(location.tabIndex)].pinned == pinned) {
        return false;
    }

    group.tabs[static_cast<size_t>(location.tabIndex)].pinned = pinned;
    MoveTab(location, {location.groupIndex, pinned ? static_cast<int>(group.tabs.size()) : 0});
    return true;
}

bool TabManager::ToggleTabPinned(TabLocation location) {
    const TabInfo* tab = Get(location);
    if (!tab) {
        return false;
    }
    return SetTabPinned(location, !tab->pinned);
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
    IndexRemoveGroup(fromGroup, moving);

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

    IndexShiftGroups(toGroup, 1);
    m_groups.insert(m_groups.begin() + toGroup, std::move(moving));
    IndexInsertGroup(toGroup);

    if (selectedGroup == -1) {
        m_selectedGroup = toGroup;
    } else {
        if (selectedGroup >= toGroup) {
            ++selectedGroup;
        }
        m_selectedGroup = selectedGroup;
    }

    MarkLayoutDirty();
    EnsureVisibleSelection();
}

void TabManager::EnsureDefaultGroup() {
    if (!m_groups.empty()) {
        return;
    }

    TabGroup group;
    group.name = std::wstring(kDefaultGroupNamePrefix) + std::to_wstring(m_groupSequence);
    group.headerVisible = true;
    ResetGroupAggregates(group);
    m_groups.emplace_back(std::move(group));
    MarkLayoutDirty();
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
    MarkLayoutDirty();
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

}

}  // namespace shelltabs
