#pragma once

#include <windows.h>
#include <string>
#include <vector>

namespace shelltabs {

struct CrashRecoveryTabInfo {
    wchar_t path[MAX_PATH];
    bool isSelected;
};

struct CrashRecoveryWindowInfo {
    bool active;
    uint32_t timestamp;
    CrashRecoveryTabInfo tabs[32]; // Max 32 tabs per window for fast logging
    uint32_t tabCount;
};

struct CrashRecoveryState {
    uint32_t signature; // 'SHCR'
    uint32_t version;   // 1
    uint32_t sequence;  // Increment on every write
    CrashRecoveryWindowInfo windows[16]; // Max 16 windows
};

class CrashRecoveryLogger {
public:
    static CrashRecoveryLogger& Instance();

    void Initialize();
    void Shutdown();

    // Fast memory-mapped write
    void LogWindowState(int windowSlot, const std::vector<std::wstring>& paths, int selectedIndex);
    
    // Clear the active flag for a window when it closes cleanly
    void ClearWindowState(int windowSlot);

    // Read the last crashed state if available
    bool ReadCrashedState(int windowSlot, std::vector<std::wstring>& outPaths, int& outSelectedIndex);

private:
    CrashRecoveryLogger() = default;
    ~CrashRecoveryLogger();

    HANDLE m_hFile = INVALID_HANDLE_VALUE;
    HANDLE m_hMap = NULL;
    CrashRecoveryState* m_state = nullptr;
    SRWLOCK m_lock = SRWLOCK_INIT;
};

} // namespace shelltabs
