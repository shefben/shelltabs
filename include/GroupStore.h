#pragma once

#include <windows.h>

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "TabManager.h"

namespace shelltabs {

struct SavedGroup {
    std::wstring name;
    COLORREF color = RGB(0, 120, 215);
    std::vector<std::wstring> tabPaths;
    TabGroupOutlineStyle outlineStyle = TabGroupOutlineStyle::kSolid;
};

class GroupStore {
public:
    static GroupStore& Instance();

    const std::vector<SavedGroup>& Groups() const noexcept { return m_groups; }
    std::vector<std::wstring> GroupNames() const;
    const SavedGroup* Find(const std::wstring& name) const;

    bool Load();
    bool Save() const;

    bool Upsert(SavedGroup group);
    bool UpdateTabs(const std::wstring& name, const std::vector<std::wstring>& tabPaths);
    bool UpdateColor(const std::wstring& name, COLORREF color);
    bool Remove(const std::wstring& name);

    void RecordChanges(const std::vector<std::pair<std::wstring, std::wstring>>& renamedGroups,
                       const std::vector<std::wstring>& removedGroupIds);
    uint64_t ChangeGeneration() const noexcept { return m_changeGeneration; }
    const std::vector<std::pair<std::wstring, std::wstring>>& LastRenamedGroups() const noexcept {
        return m_lastRenamedGroups;
    }
    const std::vector<std::wstring>& LastRemovedGroups() const noexcept { return m_lastRemovedGroups; }

private:
    GroupStore() = default;

    std::wstring ResolveStoragePath() const;
    bool EnsureLoaded() const;

    mutable bool m_loaded = false;
    mutable std::wstring m_storagePath;
    std::vector<SavedGroup> m_groups;
    uint64_t m_changeGeneration = 0;
    std::vector<std::pair<std::wstring, std::wstring>> m_lastRenamedGroups;
    std::vector<std::wstring> m_lastRemovedGroups;
};

}  // namespace shelltabs
