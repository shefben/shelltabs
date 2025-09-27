#pragma once

#include <windows.h>

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <functional>
#include <mutex>
#include <optional>
#include <queue>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

namespace shelltabs {

struct GitStatusInfo {
    bool isRepository = false;
    bool hasChanges = false;
    bool hasUntracked = false;
    int ahead = 0;
    int behind = 0;
    std::wstring branch;
    std::wstring rootPath;
};

class GitStatusCache {
public:
    static GitStatusCache& Instance();

    GitStatusInfo Query(const std::wstring& path);
    size_t AddListener(std::function<void()> callback);
    void RemoveListener(size_t id);

private:
    struct CacheEntry {
        GitStatusInfo status;
        std::chrono::steady_clock::time_point timestamp;
        bool hasStatus = false;
        bool inFlight = false;
    };

    GitStatusCache() = default;

    std::wstring ResolveRepositoryRoot(const std::wstring& path);
    GitStatusInfo ComputeStatus(const std::wstring& repoRoot) const;
    bool HasFreshEntry(const std::wstring& repoRoot, GitStatusInfo* info);
    void EnsureWorker();
    void WorkerLoop();
    bool EnqueueWork(std::wstring repoRoot);
    void HandleWorkerFailure(std::chrono::steady_clock::time_point now);
    void ResetPendingWork();
    void ScheduleWorkerRetryLocked(std::chrono::steady_clock::time_point now);
    void NotifyListeners();
    bool IsShellNamespacePath(const std::wstring& path) const;
    bool HasGitMetadata(const std::wstring& directory) const;
    std::wstring ParentDirectory(const std::wstring& path) const;
    void CacheRepositoryRoot(const std::vector<std::wstring>& probedPaths, const std::wstring& root,
                             std::chrono::steady_clock::time_point timestamp);
    std::optional<std::wstring> LookupCachedRoot(const std::wstring& path,
                                                 std::chrono::steady_clock::time_point now);

    mutable std::mutex m_mutex;
    std::unordered_map<std::wstring, CacheEntry> m_cache;
    mutable std::mutex m_rootCacheMutex;
    struct RootCacheEntry {
        std::wstring root;
        std::chrono::steady_clock::time_point timestamp;
    };
    std::unordered_map<std::wstring, RootCacheEntry> m_rootCache;

    std::mutex m_queueMutex;
    std::condition_variable m_queueCv;
    std::queue<std::wstring> m_queue;
    std::unordered_set<std::wstring> m_pendingWork;
    std::atomic<bool> m_stop{false};

    mutable std::mutex m_workerStartMutex;
    std::atomic<bool> m_workerRunning{false};
    std::atomic<bool> m_workerFailed{false};
    std::chrono::steady_clock::time_point m_nextWorkerRetry{};

    std::mutex m_listenerMutex;
    std::vector<std::pair<size_t, std::function<void()>>> m_listeners;
    size_t m_nextListenerId = 1;
};

}  // namespace shelltabs

