#pragma once

#include <string>
#include <vector>

namespace shelltabs {

struct SessionTab {
    std::wstring path;
    std::wstring name;
    std::wstring tooltip;
    bool hidden = false;
};

struct SessionGroup {
    std::wstring name;
    bool collapsed = false;
    std::vector<SessionTab> tabs;
    bool splitView = false;
    int splitPrimary = -1;
    int splitSecondary = -1;
};

struct SessionData {
    std::vector<SessionGroup> groups;
    int selectedGroup = -1;
    int selectedTab = -1;
    int groupSequence = 1;
};

class SessionStore {
public:
    SessionStore();

    bool Load(SessionData& data) const;
    bool Save(const SessionData& data) const;

private:
    std::wstring m_storagePath;
};

}  // namespace shelltabs
