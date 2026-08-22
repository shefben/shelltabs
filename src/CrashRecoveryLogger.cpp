#include "CrashRecoveryLogger.h"
#include <shlobj.h>
#include <algorithm>

namespace shelltabs {

CrashRecoveryLogger& CrashRecoveryLogger::Instance() {
    static CrashRecoveryLogger instance;
    return instance;
}

CrashRecoveryLogger::~CrashRecoveryLogger() {
    Shutdown();
}

void CrashRecoveryLogger::Initialize() {
    AcquireSRWLockExclusive(&m_lock);
    if (m_state) {
        ReleaseSRWLockExclusive(&m_lock);
        return;
    }

    wchar_t appData[MAX_PATH];
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, appData))) {
        ReleaseSRWLockExclusive(&m_lock);
        return;
    }

    std::wstring dir = std::wstring(appData) + L"\\ShellTabs";
    CreateDirectoryW(dir.c_str(), NULL);
    std::wstring file = dir + L"\\crash_state.bin";

    m_hFile = CreateFileW(file.c_str(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                          NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (m_hFile == INVALID_HANDLE_VALUE) {
        ReleaseSRWLockExclusive(&m_lock);
        return;
    }

    m_hMap = CreateFileMappingW(m_hFile, NULL, PAGE_READWRITE, 0, sizeof(CrashRecoveryState), NULL);
    if (!m_hMap) {
        CloseHandle(m_hFile);
        m_hFile = INVALID_HANDLE_VALUE;
        ReleaseSRWLockExclusive(&m_lock);
        return;
    }

    m_state = static_cast<CrashRecoveryState*>(MapViewOfFile(m_hMap, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(CrashRecoveryState)));
    if (!m_state) {
        CloseHandle(m_hMap);
        m_hMap = NULL;
        CloseHandle(m_hFile);
        m_hFile = INVALID_HANDLE_VALUE;
        ReleaseSRWLockExclusive(&m_lock);
        return;
    }

    if (m_state->signature != 0x52434853 /* 'SHCR' */) {
        // Initialize new
        ZeroMemory(m_state, sizeof(CrashRecoveryState));
        m_state->signature = 0x52434853;
        m_state->version = 1;
        m_state->sequence = 0;
    }

    ReleaseSRWLockExclusive(&m_lock);
}

void CrashRecoveryLogger::Shutdown() {
    AcquireSRWLockExclusive(&m_lock);
    if (m_state) {
        UnmapViewOfFile(m_state);
        m_state = nullptr;
    }
    if (m_hMap) {
        CloseHandle(m_hMap);
        m_hMap = NULL;
    }
    if (m_hFile != INVALID_HANDLE_VALUE) {
        CloseHandle(m_hFile);
        m_hFile = INVALID_HANDLE_VALUE;
    }
    ReleaseSRWLockExclusive(&m_lock);
}

void CrashRecoveryLogger::LogWindowState(int windowSlot, const std::vector<std::wstring>& paths, int selectedIndex) {
    if (windowSlot < 0 || windowSlot >= 16) return;

    AcquireSRWLockExclusive(&m_lock);
    if (!m_state) {
        ReleaseSRWLockExclusive(&m_lock);
        return;
    }

    auto& win = m_state->windows[windowSlot];
    win.active = true;
    win.timestamp = GetTickCount();
    win.tabCount = static_cast<uint32_t>(std::min(paths.size(), size_t(32)));
    
    for (uint32_t i = 0; i < win.tabCount; ++i) {
        wcsncpy_s(win.tabs[i].path, paths[i].c_str(), MAX_PATH);
        win.tabs[i].isSelected = (i == static_cast<uint32_t>(selectedIndex));
    }
    
    m_state->sequence++;
    ReleaseSRWLockExclusive(&m_lock);
}

void CrashRecoveryLogger::ClearWindowState(int windowSlot) {
    if (windowSlot < 0 || windowSlot >= 16) return;

    AcquireSRWLockExclusive(&m_lock);
    if (m_state) {
        m_state->windows[windowSlot].active = false;
        m_state->sequence++;
    }
    ReleaseSRWLockExclusive(&m_lock);
}

bool CrashRecoveryLogger::ReadCrashedState(int windowSlot, std::vector<std::wstring>& outPaths, int& outSelectedIndex) {
    if (windowSlot < 0 || windowSlot >= 16) return false;

    bool found = false;
    AcquireSRWLockShared(&m_lock);
    if (m_state && m_state->windows[windowSlot].active) {
        auto& win = m_state->windows[windowSlot];
        outPaths.clear();
        outSelectedIndex = -1;
        for (uint32_t i = 0; i < win.tabCount; ++i) {
            outPaths.push_back(win.tabs[i].path);
            if (win.tabs[i].isSelected) {
                outSelectedIndex = static_cast<int>(i);
            }
        }
        found = true;
    }
    ReleaseSRWLockShared(&m_lock);
    
    return found;
}

} // namespace shelltabs
