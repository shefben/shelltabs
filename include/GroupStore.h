#pragma once

#include <windows.h>

#include <string>
#include <vector>

namespace shelltabs {

struct SavedGroup {
    std::wstring name;
    COLORREF color = RGB(0, 120, 215);
    std::vector<std::wstring> tabPaths;
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

private:
    GroupStore() = default;

    std::wstring ResolveStoragePath() const;
    bool EnsureLoaded() const;

    mutable bool m_loaded = false;
    mutable std::wstring m_storagePath;
    std::vector<SavedGroup> m_groups;
};

}  // namespace shelltabs
