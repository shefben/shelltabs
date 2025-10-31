#include "TabManager.h"

#include "Logging.h"
#include "ShellTabsMessages.h"
#include "IconCache.h"

#include <algorithm>
#include <cmath>
#include <cwchar>
#include <cwctype>
#include <limits>
#include <numeric>
#include <optional>


namespace shelltabs {

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
    return BuildLookupKey(tab.pidl.get(), tab.path);
}
}  // namespace

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
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            if (!includeHidden && tab.hidden) {
                continue;
            }
            if (current.groupIndex == static_cast<int>(g) && current.tabIndex == static_cast<int>(t)) {
                continue;
            }
            const uint64_t ordinal = tab.activationOrdinal;
            const ULONGLONG tick = tab.lastActivatedTick;
            if (!hasBest || ordinal > bestOrdinal ||
                (ordinal == bestOrdinal && tick > bestTick) ||
                (ordinal == bestOrdinal && tick == bestTick &&
                 (static_cast<int>(g) < best.groupIndex ||
                  (static_cast<int>(g) == best.groupIndex && static_cast<int>(t) < best.tabIndex)))) {
                best = {static_cast<int>(g), static_cast<int>(t)};
                bestOrdinal = ordinal;
                bestTick = tick;
                hasBest = true;
            }
        }
    }

    if (!hasBest) {
        return {};
    }
    return best;
}

std::vector<TabLocation> TabManager::GetTabsByActivationOrder(bool includeHidden) const {
    struct Candidate {
        TabLocation location;
        uint64_t ordinal;
        ULONGLONG tick;
    };

    std::vector<Candidate> candidates;
    candidates.reserve(static_cast<size_t>(TotalTabCount()));
    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            if (!includeHidden && tab.hidden) {
                continue;
            }
            candidates.push_back({{static_cast<int>(g), static_cast<int>(t)}, tab.activationOrdinal,
                                  tab.lastActivatedTick});
        }
    }

    std::sort(candidates.begin(), candidates.end(), [](const Candidate& lhs, const Candidate& rhs) {
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
    });

    std::vector<TabLocation> order;
    order.reserve(candidates.size());
    for (const auto& candidate : candidates) {
        order.emplace_back(candidate.location);
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

    auto& group = m_groups[static_cast<size_t>(groupIndex)];
    const int desiredIndex = info.pinned ? CountLeadingPinned(group)
                                         : static_cast<int>(group.tabs.size());
    const int insertIndex = std::clamp(desiredIndex, 0, static_cast<int>(group.tabs.size()));
    group.tabs.insert(group.tabs.begin() + insertIndex, std::move(info));

    const TabLocation location{groupIndex, insertIndex};
    if (select) {
        m_selectedGroup = location.groupIndex;
        m_selectedTab = location.tabIndex;
        group.collapsed = false;
    } else if (m_selectedGroup == location.groupIndex && m_selectedTab >= location.tabIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();
    UpdateSelectionActivation(previousSelection);
    RebuildIndices();
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

    if (location.tabIndex >= 0 && location.tabIndex < static_cast<int>(group.tabs.size())) {
        InvalidateTabIcon(group.tabs[static_cast<size_t>(location.tabIndex)]);
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

    EnsureVisibleSelection();
    UpdateSelectionActivation(previousSelection);
    RebuildIndices();
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

    EnsureVisibleSelection();

    UpdateSelectionActivation(previousSelection);
    RebuildIndices();
    return removed;
}

TabLocation TabManager::InsertTab(TabInfo tab, int groupIndex, int tabIndex, bool select) {
    EnsureDefaultGroup();

    const TabLocation previousSelection = SelectedLocation();

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

    if (select) {
        m_selectedGroup = groupIndex;
        m_selectedTab = insertIndex;
        group.collapsed = false;
    } else if (m_selectedGroup == groupIndex && m_selectedTab >= insertIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();

    UpdateSelectionActivation(previousSelection);
    RebuildIndices();
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

    EnsureDefaultGroup();
    EnsureVisibleSelection();
    RebuildIndices();

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

    if (m_selectedGroup >= insertIndex) {
        ++m_selectedGroup;
    }

    EnsureVisibleSelection();
    RebuildIndices();
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
    EnsureDefaultGroup();
    RebuildIndices();
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
}

std::vector<TabViewItem> TabManager::BuildView() const {
    std::vector<TabViewItem> items;
    items.reserve(TotalTabCount() + static_cast<int>(m_groups.size()));

    for (size_t g = 0; g < m_groups.size(); ++g) {
        const auto& group = m_groups[g];
        const size_t total = group.tabs.size();
        size_t visible = 0;
        size_t hidden = 0;
        ULONGLONG groupLastTick = 0;
        uint64_t groupLastOrdinal = 0;

        for (const auto& tab : group.tabs) {
            if (tab.activationOrdinal > groupLastOrdinal ||
                (tab.activationOrdinal == groupLastOrdinal && tab.lastActivatedTick > groupLastTick)) {
                groupLastOrdinal = tab.activationOrdinal;
                groupLastTick = tab.lastActivatedTick;
            }
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
            header.outlineStyle = group.outlineStyle;
            header.savedGroupId = group.savedGroupId;
            header.isSavedGroup = !group.savedGroupId.empty();
            header.headerVisible = group.headerVisible;
            header.lastActivatedTick = groupLastTick;
            header.activationOrdinal = groupLastOrdinal;
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
        ULONGLONG groupLastTick = 0;
        uint64_t groupLastOrdinal = 0;

        for (const auto& tab : group.tabs) {
            if (tab.activationOrdinal > groupLastOrdinal ||
                (tab.activationOrdinal == groupLastOrdinal && tab.lastActivatedTick > groupLastTick)) {
                groupLastOrdinal = tab.activationOrdinal;
                groupLastTick = tab.lastActivatedTick;
            }
        }

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
    if (m_progressListeners.empty()) {
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
        PostMessageW(hwnd, message, 0, 0);
        ++it;
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
        const auto& group = m_groups[g];
        for (size_t t = 0; t < group.tabs.size(); ++t) {
            const auto& tab = group.tabs[t];
            const std::wstring key = BuildLookupKey(tab);
            if (key.empty()) {
                continue;
            }
            m_locationIndex[key].push_back({static_cast<int>(g), static_cast<int>(t)});
        }
    }
}

bool TabManager::ApplyProgress(TabInfo* tab, std::optional<double> fraction, ULONGLONG now) {
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
    return changed;
}

bool TabManager::ClearProgress(TabInfo* tab) {
    if (!tab) {
        return false;
    }
    if (!tab->progress.active) {
        return false;
    }
    tab->progress = {};
    return true;
}

void TabManager::NormalizePinnedOrder(TabGroup& group) {
    std::stable_partition(group.tabs.begin(), group.tabs.end(), [](const TabInfo& tab) { return tab.pinned; });
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

    NotifyProgressListeners();
}

void TabManager::RecalculateNextActivationOrdinal() {
    uint64_t maxOrdinal = 0;
    for (const auto& group : m_groups) {
        for (const auto& tab : group.tabs) {
            if (tab.activationOrdinal > maxOrdinal) {
                maxOrdinal = tab.activationOrdinal;
            }
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
    if (ApplyProgress(tab, fraction, now)) {
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
    if (ClearProgress(tab)) {
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
    const TabLocation previousSelection = SelectedLocation();
    tab->hidden = true;
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
    const TabLocation previousSelection = SelectedLocation();
    tab->hidden = false;
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
    for (auto& tab : group->tabs) {
        tab.hidden = false;
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
    RebuildIndices();

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

    if (wasSelected) {
        m_selectedGroup = to.groupIndex;
        m_selectedTab = to.tabIndex;
    } else if (m_selectedGroup == to.groupIndex && m_selectedTab >= to.tabIndex) {
        ++m_selectedTab;
    }

    EnsureVisibleSelection();
    RebuildIndices();
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
    RebuildIndices();
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

}

}  // namespace shelltabs
