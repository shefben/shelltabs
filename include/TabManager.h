#pragma once

#include <cmath>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>
#include <functional>
#include <list>
#include <mutex>
#include <unordered_map>

#include "Utilities.h"

namespace shelltabs {

enum class TabViewItemType {
    kGroupHeader,
    kTab,
};

enum class TabGroupOutlineStyle {
    kSolid,
    kDashed,
    kDotted,
};

struct TabLocation {
    int groupIndex = -1;
    int tabIndex = -1;

    bool IsValid() const noexcept { return groupIndex >= 0 && tabIndex >= 0; }
};

struct TabProgressState {
    bool active = false;
    bool indeterminate = false;
    double fraction = 0.0;
    ULONGLONG lastUpdateTick = 0;
};

struct TabProgressView {
    bool visible = false;
    bool indeterminate = false;
    double fraction = 0.0;

    bool operator==(const TabProgressView& other) const noexcept {
        return visible == other.visible && indeterminate == other.indeterminate &&
               std::abs(fraction - other.fraction) < 1e-4;
    }
    bool operator!=(const TabProgressView& other) const noexcept { return !(*this == other); }
};

struct NavigationHistoryEntry {
    UniquePidl pidl;
    std::wstring path;
    std::wstring name;
    ULONGLONG timestamp = 0;
};

struct NavigationHistory {
    std::vector<NavigationHistoryEntry> entries;
    int currentIndex = -1;

    bool CanGoBack() const noexcept {
        return currentIndex > 0;
    }

    bool CanGoForward() const noexcept {
        return currentIndex >= 0 && currentIndex < static_cast<int>(entries.size()) - 1;
    }

    void Clear() {
        entries.clear();
        currentIndex = -1;
    }

    bool IsEmpty() const noexcept {
        return entries.empty();
    }
};

inline double ClampProgress(double value) noexcept {
    if (value < 0.0) {
        return 0.0;
    }
    if (value > 1.0) {
        return 1.0;
    }
    return value;
}

struct TabInfo {
    UniquePidl pidl;
    std::wstring name;
    std::wstring tooltip;
    bool hidden = false;
    bool pinned = false;
    std::wstring path;
    std::wstring normalizedLookupKey;
    TabProgressState progress;
    ULONGLONG lastActivatedTick = 0;
    uint64_t activationOrdinal = 0;
    uint64_t activationEpoch = 0;
    NavigationHistory navigationHistory;

    void RefreshNormalizedLookupKey();
};

struct TabGroup {
    std::wstring name;
    bool collapsed = false;
    std::vector<TabInfo> tabs;
    bool headerVisible = true;
    std::wstring savedGroupId;
    bool hasCustomOutline = false;
    COLORREF outlineColor = RGB(0, 120, 215);
    TabGroupOutlineStyle outlineStyle = TabGroupOutlineStyle::kSolid;
    size_t visibleCount = 0;
    size_t hiddenCount = 0;
    int lastActivatedTabIndex = -1;
    uint64_t lastActivationOrdinal = 0;
    ULONGLONG lastActivatedTick = 0;
    int lastVisibleActivatedTabIndex = -1;
    uint64_t lastVisibleActivationOrdinal = 0;
    ULONGLONG lastVisibleActivatedTick = 0;
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
    std::wstring path;
    bool hasCustomOutline = false;
    COLORREF outlineColor = 0;
    TabGroupOutlineStyle outlineStyle = TabGroupOutlineStyle::kSolid;
    std::wstring savedGroupId;
    bool isSavedGroup = false;
    bool headerVisible = true;
    TabProgressView progress;
    ULONGLONG lastActivatedTick = 0;
    uint64_t activationOrdinal = 0;
    bool pinned = false;
    uint64_t stableId = 0;
};

uint64_t ComputeTabViewStableId(const TabViewItem& item) noexcept;

struct TabProgressSnapshotEntry {
    TabViewItemType type = TabViewItemType::kGroupHeader;
    TabLocation location;
    TabProgressView progress;
    ULONGLONG lastActivatedTick = 0;
    uint64_t activationOrdinal = 0;
};

using TabProgressSnapshot = std::vector<TabProgressSnapshotEntry>;

struct TabProgressUpdatePayload {
    uint32_t layoutVersion = 0;
    std::vector<TabProgressSnapshotEntry> entries;
};

class TabManager {
public:
    struct ExplorerWindowId {
        HWND hwnd = nullptr;
        uintptr_t frameCookie = 0;

        bool IsValid() const noexcept { return hwnd != nullptr && frameCookie != 0; }
        bool operator==(const ExplorerWindowId& other) const noexcept {
            return hwnd == other.hwnd && frameCookie == other.frameCookie;
        }
        bool operator!=(const ExplorerWindowId& other) const noexcept { return !(*this == other); }
    };

    struct ExplorerWindowIdHash {
        size_t operator()(const ExplorerWindowId& value) const noexcept {
            const uintptr_t hwndValue = reinterpret_cast<uintptr_t>(value.hwnd);
            size_t result = std::hash<uintptr_t>{}(hwndValue);
            const size_t cookieHash = std::hash<uintptr_t>{}(value.frameCookie);
            // Mix using a variant of boost::hash_combine
            result ^= cookieHash + 0x9e3779b97f4a7c15ULL + (result << 6) + (result >> 2);
            return result;
        }
    };

    TabManager();
    ~TabManager();

    int TotalTabCount() const noexcept;
    static TabManager& Get();
    static TabManager* Find(ExplorerWindowId id);

    void SetWindowId(ExplorerWindowId id);
    void ClearWindowId();
    ExplorerWindowId GetWindowId() const noexcept { return m_windowId; }
    static size_t ActiveWindowCount();

    TabLocation SelectedLocation() const noexcept { return {m_selectedGroup, m_selectedTab}; }
    void SetSelectedLocation(TabLocation location);

    int GroupCount() const noexcept { return static_cast<int>(m_groups.size()); }
    const TabGroup* GetGroup(int index) const noexcept;
    TabGroup* GetGroup(int index) noexcept;

    const TabInfo* Get(TabLocation location) const noexcept;
    TabInfo* Get(TabLocation location) noexcept;
    TabLocation Find(PCIDLIST_ABSOLUTE pidl) const;

    TabLocation GetLastActivatedTab(bool includeHidden = false) const;
    std::vector<TabLocation> GetTabsByActivationOrder(bool includeHidden = false) const;

    TabLocation Add(UniquePidl pidl, std::wstring name, std::wstring tooltip, bool select, int groupIndex = -1,
                    bool pinned = false);
    void Remove(TabLocation location);
    std::optional<TabInfo> TakeTab(TabLocation location);
    TabLocation InsertTab(TabInfo tab, int groupIndex, int tabIndex, bool select);
    std::optional<TabGroup> TakeGroup(int groupIndex);
    int InsertGroup(TabGroup group, int insertIndex);
    void Clear();
    void Restore(std::vector<TabGroup> groups, int selectedGroup, int selectedTab, int groupSequence);

    std::vector<TabViewItem> BuildView() const;
    TabProgressSnapshot CollectProgressStates() const;

    void RegisterProgressListener(HWND hwnd);
    void UnregisterProgressListener(HWND hwnd);
    void TouchFolderOperation(PCIDLIST_ABSOLUTE folder, std::optional<double> fraction = std::nullopt);
    void ClearFolderOperation(PCIDLIST_ABSOLUTE folder);
    std::vector<TabLocation> ExpireFolderOperations(ULONGLONG now, ULONGLONG timeoutMs);
    bool HasActiveProgress() const;

    uint32_t GetLayoutVersion() const noexcept { return m_layoutVersion; }

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

    bool SetTabPinned(TabLocation location, bool pinned);
    bool ToggleTabPinned(TabLocation location);

    int NextGroupSequence() const noexcept { return m_groupSequence; }

    // Navigation history methods
    void RecordNavigation(TabLocation location, UniquePidl pidl, std::wstring path, std::wstring name);
    std::optional<NavigationHistoryEntry> NavigateBack(TabLocation location);
    std::optional<NavigationHistoryEntry> NavigateForward(TabLocation location);
    bool CanNavigateBack(TabLocation location) const;
    bool CanNavigateForward(TabLocation location) const;
    void ClearNavigationHistory(TabLocation location);

private:
    struct ActivationEntry {
        TabLocation location;
        uint64_t ordinal = 0;
        ULONGLONG tick = 0;
        uint64_t epoch = 0;
    };

    using ActivationList = std::list<ActivationEntry>;

    void EnsureDefaultGroup();
    void EnsureVisibleSelection();
    void NotifyProgressListeners();
    TabLocation FindByPath(const std::wstring& path) const;
    TabLocation ResolveFromIndex(const std::wstring& key, PCIDLIST_ABSOLUTE pidl, bool requireVisible) const;
    TabLocation ScanForPidl(PCIDLIST_ABSOLUTE pidl) const;
    TabLocation ScanForPath(const std::wstring& path) const;
    struct ProgressUpdateKey {
        TabViewItemType type = TabViewItemType::kGroupHeader;
        TabLocation location;
    };

    void QueueProgressUpdate(TabViewItemType type, TabLocation location);
    bool BuildProgressEntry(const ProgressUpdateKey& key, TabProgressSnapshotEntry* entry) const;
    void MarkLayoutDirty() noexcept;
    bool ApplyProgress(TabLocation location, TabInfo* tab, std::optional<double> fraction, ULONGLONG now);
    bool ClearProgress(TabLocation location, TabInfo* tab);
    void UpdateSelectionActivation(TabLocation previousSelection);
    void RecalculateNextActivationOrdinal();
    void NormalizePinnedOrder(TabGroup& group);
    void RebuildIndices();
    void RebuildActivationOrder();
    void IndexInsertTab(TabLocation location);
    void IndexRemoveTab(TabLocation location, const TabInfo& tab);
    void IndexInsertGroup(int groupIndex);
    void IndexRemoveGroup(int groupIndex, const TabGroup& group);
    void IndexShiftTabs(int groupIndex, int startTabIndex, int delta);
    void IndexShiftGroups(int startGroupIndex, int delta);
    void ActivationInsertTab(TabLocation location);
    void ActivationRemoveTab(TabLocation location);
    void ActivationShiftTabs(int groupIndex, int startTabIndex, int delta);
    void ActivationShiftGroups(int startGroupIndex, int delta);
    void ActivationUpdateTab(TabLocation location);
    static bool ActivationPrecedes(const ActivationEntry& lhs, const ActivationEntry& rhs) noexcept;
    static uint64_t EncodeActivationKey(TabLocation location) noexcept;
    static bool IsBetterActivation(uint64_t candidateOrdinal, ULONGLONG candidateTick, int candidateIndex,
                                   uint64_t bestOrdinal, ULONGLONG bestTick, int bestIndex) noexcept;
    static void ResetGroupAggregates(TabGroup& group) noexcept;
    static void AccumulateGroupAggregates(TabGroup& group, const TabInfo& tab, int tabIndex) noexcept;
    static void RefreshGroupAggregates(TabGroup& group) noexcept;
    void HandleTabInserted(TabGroup& group, int tabIndex);
    void HandleTabRemoved(TabGroup& group, int tabIndex, bool wasHidden);
    void HandleTabVisibilityChanged(TabGroup& group, int tabIndex, bool wasHidden, bool isHidden);
    void HandleTabActivationUpdated(TabGroup& group, int tabIndex);
    int FindBestActivatedTabIndex(const TabGroup& group, bool includeHidden, int excludeTabIndex) const;

    std::vector<TabGroup> m_groups;
    int m_selectedGroup = -1;
    int m_selectedTab = -1;
    int m_groupSequence = 1;
    std::vector<HWND> m_progressListeners;
    uint64_t m_nextActivationOrdinal = 1;
    uint64_t m_activationEpoch = 0;
    uint64_t m_lastActivationOrdinalSeen = 0;
    ULONGLONG m_lastActivationTickSeen = 0;
    ExplorerWindowId m_windowId{};
    std::unordered_map<std::wstring, std::vector<TabLocation>> m_locationIndex;
    std::vector<ProgressUpdateKey> m_pendingProgressUpdates;
    uint32_t m_layoutVersion = 1;
#if defined(SHELLTABS_BUILD_TESTS)
    std::vector<TabProgressSnapshotEntry> m_lastProgressUpdatesForTest;
    uint32_t m_lastProgressLayoutVersionForTest = 0;
#endif
    ActivationList m_activationOrder;
    std::unordered_map<uint64_t, ActivationList::iterator> m_activationLookup;

    static std::mutex s_windowMutex;
    static std::unordered_map<ExplorerWindowId, TabManager*, ExplorerWindowIdHash> s_windowMap;
};

}  // namespace shelltabs

