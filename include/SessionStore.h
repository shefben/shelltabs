#pragma once

#include <windows.h>

#include <cstdint>
#include <string>
#include <vector>
#include <optional>
#include <atomic>

#include "OptionsStore.h"
#include "TabManager.h"

namespace shelltabs {

struct SessionTab {
    std::wstring path;
    std::wstring name;
    std::wstring tooltip;
    bool hidden = false;
    bool pinned = false;
    ULONGLONG lastActivatedTick = 0;
    uint64_t activationOrdinal = 0;
};

struct SessionGroup {
    std::wstring name;
    bool collapsed = false;
    std::vector<SessionTab> tabs;
    bool headerVisible = true;
    bool hasOutline = false;
    COLORREF outlineColor = RGB(0, 120, 215);
    std::wstring savedGroupId;
    TabGroupOutlineStyle outlineStyle = TabGroupOutlineStyle::kSolid;
};

struct SessionClosedTab {
    SessionTab tab;
    int index = -1;
};

struct SessionClosedSet {
    int groupIndex = -1;
    bool groupRemoved = false;
    int selectionIndex = -1;
    bool hasGroupInfo = false;
    SessionGroup groupInfo;
    std::vector<SessionClosedTab> tabs;
};

struct SessionData {
    std::vector<SessionGroup> groups;
    int selectedGroup = -1;
    int selectedTab = -1;
    int groupSequence = 1;
    TabBandDockMode dockMode = TabBandDockMode::kAutomatic;
    std::optional<SessionClosedSet> lastClosed;
};

class SessionStore {
public:
    struct RecoverableSessionCandidate {
        std::wstring token;
        std::wstring storagePath;
        uint64_t lastActivityTicks = 0;
        bool hasLock = false;
        bool hasTemp = false;
        bool hasCheckpoint = false;

        [[nodiscard]] int CompanionCount() const noexcept {
            return static_cast<int>(hasLock) + static_cast<int>(hasTemp) + static_cast<int>(hasCheckpoint);
        }
    };

    SessionStore();
    explicit SessionStore(std::wstring storagePath);

    bool Load(SessionData& data) const;
    bool Save(const SessionData& data) const;

    static std::wstring BuildPathForToken(const std::wstring& token);
    static std::vector<RecoverableSessionCandidate> EnumerateRecoverableSessions();
    static std::optional<RecoverableSessionCandidate> SelectRecoverableSession(
        const std::vector<RecoverableSessionCandidate>& candidates);
    static std::optional<std::wstring> LoadPersistedWindowToken();
    static bool PersistWindowToken(const std::wstring& token);
    static void ClearPersistedWindowToken();
    bool WasPreviousSessionUnclean() const;
    void MarkSessionActive() const;
    void ClearSessionMarker() const;
    void SetMarkerReady(bool ready) const;
    bool MarkerReady() const noexcept;

private:
    std::wstring m_storagePath;
    mutable std::optional<std::wstring> m_lastSerializedSnapshot;
    mutable bool m_pendingCheckpointCleanup = false;
    mutable std::atomic<bool> m_markerReady = false;
};

}  // namespace shelltabs
