#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
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

struct SessionHistoryEntry {
    std::wstring path;
    std::wstring name;
    ULONGLONG timestamp = 0;
};

struct SessionTab {
    std::wstring path;
    std::wstring name;
    std::wstring tooltip;
    bool hidden = false;
    bool pinned = false;
    ULONGLONG lastActivatedTick = 0;
    uint64_t activationOrdinal = 0;
    bool hasScrollPosition = false;
    int32_t scrollX = 0;
    int32_t scrollY = 0;
    std::vector<SessionHistoryEntry> history;
    int historyIndex = -1;
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

    // Clear crash state: delete marker and discard any unclaimed pending data
    // that was loaded from the crashed session. Call after the first window has
    // successfully initialized (restored or not).
    void ClearCrashState();

    // Number of currently registered slots.
    int RegisteredCount() const;

    // True if the pending pool has unclaimed session data.
    bool HasPendingData() const;

    std::vector<std::wstring> GetAllTabPaths() const;

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

    // Timestamp of the last Unregister that returned data to the pending pool.
    // Used to implement a cooldown: if a window claims data and then closes
    // almost immediately (returning data to pending), we suppress the next
    // claim for a brief period so the next window doesn't enter the same
    // claim → restore → close → return cycle.
    ULONGLONG m_lastReturnTick = 0;

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

struct SavedTabSession {
    ULONGLONG timestamp = 0;
    std::vector<std::wstring> paths;
};

class SavedTabSessionManager {
public:
    static SavedTabSessionManager& Instance();
    void SaveCurrentSession();
    std::vector<SavedTabSession> GetSavedSessions() const;

private:
    SavedTabSessionManager() = default;
    std::wstring GetFilePath() const;
};

}  // namespace shelltabs
