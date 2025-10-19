#pragma once

#include <windows.h>

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
    bool headerVisible = true;
    bool hasOutline = false;
    COLORREF outlineColor = RGB(0, 120, 215);
    std::wstring savedGroupId;
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
    explicit SessionStore(std::wstring storagePath);

    bool Load(SessionData& data) const;
    bool Save(const SessionData& data) const;

    static std::wstring BuildPathForToken(const std::wstring& token);
    static bool WasPreviousSessionUnclean();
    static void MarkSessionActive();
    static void ClearSessionMarker();

private:
    std::wstring m_storagePath;
};

}  // namespace shelltabs
