#pragma once

#include "Utilities.h"

#include <windows.h>
#include <winhttp.h>

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <deque>
#include <functional>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <vector>

namespace shelltabs::http {

struct DirectoryEntry;  // Forward from HtmlDirectoryParser.h

struct HttpConnectionOptions {
    std::wstring host;
    std::wstring basePath;
    unsigned short port = 443;
    bool useHttps = true;
    std::chrono::seconds poolIdleTimeout{180};
};

using HttpProgressCallback = std::function<bool(ULONGLONG downloaded, ULONGLONG total)>;

class HttpClient {
public:
    HttpClient();
    ~HttpClient();

    HttpClient(const HttpClient&) = delete;
    HttpClient& operator=(const HttpClient&) = delete;

    HRESULT ListDirectory(const HttpConnectionOptions& options, const std::wstring& path,
                          std::vector<DirectoryEntry>* results);

    HRESULT DownloadFile(const HttpConnectionOptions& options, const std::wstring& remotePath,
                         const std::wstring& localPath, HttpProgressCallback callback = nullptr);

    HRESULT DownloadToMemory(const HttpConnectionOptions& options, const std::wstring& remotePath,
                             std::vector<BYTE>* data);

    HRESULT HeadRequest(const HttpConnectionOptions& options, const std::wstring& remotePath,
                        std::wstring* contentType, ULONGLONG* contentLength);

    void ClearConnectionPool();

private:
    struct WinHttpHandle {
        HINTERNET handle = nullptr;
        WinHttpHandle() = default;
        explicit WinHttpHandle(HINTERNET h) : handle(h) {}
        ~WinHttpHandle() {
            if (handle) {
                WinHttpCloseHandle(handle);
            }
        }
        WinHttpHandle(const WinHttpHandle&) = delete;
        WinHttpHandle& operator=(const WinHttpHandle&) = delete;
        WinHttpHandle(WinHttpHandle&& other) noexcept : handle(other.handle) { other.handle = nullptr; }
        WinHttpHandle& operator=(WinHttpHandle&& other) noexcept {
            if (this != &other) {
                if (handle) WinHttpCloseHandle(handle);
                handle = other.handle;
                other.handle = nullptr;
            }
            return *this;
        }
        HINTERNET get() const noexcept { return handle; }
    };

    struct SessionKey {
        std::wstring host;
        unsigned short port = 0;
        bool useHttps = false;
        bool operator==(const SessionKey& other) const {
            return host == other.host && port == other.port && useHttps == other.useHttps;
        }
    };

    struct SessionKeyHasher {
        size_t operator()(const SessionKey& key) const {
            size_t h = std::hash<std::wstring>{}(key.host);
            h ^= std::hash<unsigned short>{}(key.port) + 0x9e3779b9 + (h << 6) + (h >> 2);
            h ^= std::hash<bool>{}(key.useHttps) + 0x9e3779b9 + (h << 6) + (h >> 2);
            return h;
        }
    };

    struct PooledSession {
        WinHttpHandle session;
        WinHttpHandle connection;
        std::chrono::steady_clock::time_point lastUsed;
    };

    HRESULT AcquireSession(const HttpConnectionOptions& options, HINTERNET* session, HINTERNET* connection);
    void ReturnSession(const HttpConnectionOptions& options, WinHttpHandle session, WinHttpHandle connection);
    HRESULT SendRequest(HINTERNET connection, const wchar_t* verb, const std::wstring& path,
                        bool useHttps, WinHttpHandle* request);
    HRESULT ReadResponseBody(HINTERNET request, std::vector<BYTE>* data);

    std::mutex poolMutex_;
    std::unordered_map<SessionKey, PooledSession, SessionKeyHasher> pool_;
};

// Download task for the parallel queue.
struct DownloadTask {
    HttpConnectionOptions options;
    std::wstring remotePath;
    std::wstring localPath;
    HttpProgressCallback callback;
};

// Thread-pool based download queue with configurable concurrency and speed limiting.
class HttpDownloadQueue {
public:
    HttpDownloadQueue() = default;
    ~HttpDownloadQueue();

    HttpDownloadQueue(const HttpDownloadQueue&) = delete;
    HttpDownloadQueue& operator=(const HttpDownloadQueue&) = delete;

    void Configure(int maxConcurrent, int speedLimitKBps);
    HRESULT Enqueue(DownloadTask task);
    void CancelAll();
    size_t PendingCount() const;

private:
    void EnsureWorkers();
    void WorkerLoop();
    void ThrottleIfNeeded(size_t bytesJustRead);

    std::mutex m_mutex;
    std::condition_variable m_cv;
    std::deque<DownloadTask> m_queue;
    std::vector<std::thread> m_workers;
    std::atomic<bool> m_shutdown{false};
    int m_maxConcurrent = 4;
    int m_speedLimitKBps = 0;
    bool m_workersStarted = false;

    // Speed tracking shared across workers
    std::mutex m_speedMutex;
    std::atomic<size_t> m_bytesThisSecond{0};
    std::chrono::steady_clock::time_point m_windowStart{std::chrono::steady_clock::now()};
};

}  // namespace shelltabs::http
