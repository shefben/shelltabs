#include "DirectoryWatcher.h"
#include "Logging.h"

namespace shelltabs {

DirectoryWatcher::DirectoryWatcher(HWND targetWindow, UINT messageId)
    : m_targetWindow(targetWindow), m_messageId(messageId) {
    m_iocp = CreateIoCompletionPort(INVALID_HANDLE_VALUE, nullptr, 0, 1);
    m_thread = std::thread(&DirectoryWatcher::WorkerThread, this);
}

DirectoryWatcher::~DirectoryWatcher() {
    Stop();
}

void DirectoryWatcher::Stop() {
    if (m_stop.exchange(true)) {
        return;
    }
    
    if (m_iocp) {
        PostQueuedCompletionStatus(m_iocp, 0, 0, nullptr);
    }
    
    if (m_thread.joinable()) {
        m_thread.join();
    }
    
    std::lock_guard<std::mutex> lock(m_mutex);
    if (m_directoryHandle != INVALID_HANDLE_VALUE) {
        CancelIo(m_directoryHandle);
        CloseHandle(m_directoryHandle);
        m_directoryHandle = INVALID_HANDLE_VALUE;
    }
    if (m_iocp) {
        CloseHandle(m_iocp);
        m_iocp = nullptr;
    }
}

void DirectoryWatcher::Watch(const std::wstring& directory) {
    std::lock_guard<std::mutex> lock(m_mutex);
    
    if (m_currentDirectory == directory) {
        return;
    }
    m_currentDirectory = directory;
    
    if (m_directoryHandle != INVALID_HANDLE_VALUE) {
        CancelIo(m_directoryHandle);
        CloseHandle(m_directoryHandle);
        m_directoryHandle = INVALID_HANDLE_VALUE;
    }
    
    if (directory.empty()) {
        return;
    }
    
    m_directoryHandle = CreateFileW(
        directory.c_str(),
        FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED,
        nullptr
    );
    
    if (m_directoryHandle == INVALID_HANDLE_VALUE) {
        LogMessage(LogLevel::Warning, L"DirectoryWatcher: failed to open directory %s", directory.c_str());
        return;
    }
    
    CreateIoCompletionPort(m_directoryHandle, m_iocp, (ULONG_PTR)m_directoryHandle, 1);
    
    memset(&m_overlapped, 0, sizeof(m_overlapped));
    if (!ReadDirectoryChangesW(
        m_directoryHandle,
        m_buffer,
        sizeof(m_buffer),
        FALSE,
        FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_ATTRIBUTES | FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE,
        nullptr,
        &m_overlapped,
        nullptr
    )) {
        LogMessage(LogLevel::Warning, L"DirectoryWatcher: ReadDirectoryChangesW failed");
        CloseHandle(m_directoryHandle);
        m_directoryHandle = INVALID_HANDLE_VALUE;
    } else {
        LogMessage(LogLevel::Info, L"DirectoryWatcher: Watching %s", directory.c_str());
    }
}

void DirectoryWatcher::WorkerThread() {
    DWORD bytesTransferred = 0;
    ULONG_PTR completionKey = 0;
    LPOVERLAPPED overlapped = nullptr;
    
    while (!m_stop) {
        BOOL success = GetQueuedCompletionStatus(m_iocp, &bytesTransferred, &completionKey, &overlapped, INFINITE);
        
        if (m_stop) {
            break;
        }
        
        if (!success) {
            continue;
        }
        
        if (completionKey == 0 && overlapped == nullptr) {
            break;
        }
        
        if (overlapped) {
            PostMessageW(m_targetWindow, m_messageId, 0, 0);
            
            std::lock_guard<std::mutex> lock(m_mutex);
            if (m_directoryHandle != INVALID_HANDLE_VALUE && (HANDLE)completionKey == m_directoryHandle) {
                memset(&m_overlapped, 0, sizeof(m_overlapped));
                ReadDirectoryChangesW(
                    m_directoryHandle,
                    m_buffer,
                    sizeof(m_buffer),
                    FALSE,
                    FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_ATTRIBUTES | FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE,
                    nullptr,
                    &m_overlapped,
                    nullptr
                );
            }
        }
    }
}

} // namespace shelltabs
