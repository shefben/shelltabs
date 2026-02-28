#pragma once

#include <windows.h>

#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <vector>

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

// A single window's session within the coordinated file.
struct WindowSession {
    int slot = -1;
    bool claimed = false;
    SessionData data;
};

class TabBand;

// Process-wide singleton that manages a single session.db containing ALL windows.
// All TabBands in the same explorer.exe process register here.
class SessionCoordinator {
public:
    static SessionCoordinator& Instance();

    // Called from EnsureSessionStore. Returns a slot id.
    int Register(TabBand* band);

    // Called from DisconnectSite (via SessionStore destructor).
    void Unregister(int slot);

    // Dequeue the next unclaimed window block from a loaded session file.
    bool ClaimWindowData(int slot, SessionData& outData);

    // Report current window state (called from SaveSession).
    void UpdateWindowData(int slot, const SessionData& data);

    // Atomic write of all windows to disk.
    bool SaveAll();

    // Was the previous session unclean (crash marker present)?
    bool WasCrash() const;

    // Create session.active marker (first Register).
    void MarkActive();

    // Delete session.active marker (last Unregister).
    void MarkClean();

    // Number of currently registered slots.
    int RegisteredCount() const;

private:
    SessionCoordinator();
    ~SessionCoordinator() = default;
    SessionCoordinator(const SessionCoordinator&) = delete;
    SessionCoordinator& operator=(const SessionCoordinator&) = delete;

    void LoadSessionFile();
    void MigrateFromOldFormat();
    void CleanupOldFiles();
    std::wstring SerializeAllWindows() const;

    mutable std::recursive_mutex m_mutex;
    std::wstring m_sessionPath;       // path to session.db
    std::wstring m_markerPath;        // path to session.active
    bool m_loaded = false;
    bool m_wasCrash = false;
    int m_nextSlot = 0;

    struct SlotEntry {
        TabBand* band = nullptr;
        SessionData data;
        bool hasData = false;
    };
    std::vector<std::pair<int, SlotEntry>> m_slots;

    // Unclaimed window data loaded from the session file on startup.
    std::vector<SessionData> m_pendingWindows;

    // Last serialized snapshot for dedup.
    std::optional<std::wstring> m_lastSnapshot;
};

// Thin wrapper that holds a slot id and delegates to SessionCoordinator.
// Each TabBand owns one SessionStore instance.
class SessionStore {
public:
    SessionStore();
    ~SessionStore();

    SessionStore(const SessionStore&) = delete;
    SessionStore& operator=(const SessionStore&) = delete;

    bool Load(SessionData& data) const;
    bool Save(const SessionData& data);
    int Slot() const noexcept { return m_slot; }

private:
    int m_slot = -1;
};

}  // namespace shelltabs
