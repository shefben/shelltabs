#pragma once
#include <windows.h>
#include <string>
#include <thread>
#include <atomic>
#include <mutex>

namespace shelltabs {

class DirectoryWatcher {
public:
    DirectoryWatcher(HWND targetWindow, UINT messageId);
    ~DirectoryWatcher();

    void Watch(const std::wstring& directory);
    void Stop();

private:
    void WorkerThread();

    HWND m_targetWindow;
    UINT m_messageId;
    std::wstring m_currentDirectory;
    HANDLE m_directoryHandle = INVALID_HANDLE_VALUE;
    HANDLE m_iocp = nullptr;
    std::thread m_thread;
    std::atomic<bool> m_stop{false};
    std::mutex m_mutex;
    
    alignas(8) char m_buffer[65536];
    OVERLAPPED m_overlapped{};
};

} // namespace shelltabs
