#include "GitStatus.h"

#include <Shlwapi.h>

#include <PathCch.h>

#include <algorithm>
#include <chrono>
#include <cerrno>
#include <cstdio>
#include <exception>
#include <filesystem>
#include <sstream>
#include <string>
#include <system_error>
#include <vector>
#include <cwchar>
#include <wchar.h>

#include "Logging.h"
#include "Utilities.h"

namespace shelltabs {

namespace {
constexpr std::chrono::seconds kCacheTtl(5);
constexpr std::chrono::seconds kRootCacheTtl(30);
constexpr size_t kMaxRootCacheEntries = 256;
constexpr size_t kMaxQueueDepth = 64;
constexpr std::chrono::seconds kWorkerRetryDelay(5);

std::wstring QuoteCommandArgument(const std::wstring& argument) {
    std::wstring result;
    result.reserve(argument.size() + 2);
    result.push_back(L'"');

    size_t backslashCount = 0;
    for (wchar_t ch : argument) {
        if (ch == L'\\') {
            ++backslashCount;
            continue;
        }

        if (ch == L'"') {
            result.append(backslashCount * 2 + 1, L'\\');
            result.push_back(L'"');
            backslashCount = 0;
            continue;
        }

        if (backslashCount != 0) {
            result.append(backslashCount, L'\\');
            backslashCount = 0;
        }

        result.push_back(ch);
    }

    if (backslashCount != 0) {
        result.append(backslashCount * 2, L'\\');
    }

    result.push_back(L'"');
    return result;
}

std::wstring FromUtf8(const char* value) {
    if (!value || !*value) {
        return {};
    }
    int required = MultiByteToWideChar(CP_UTF8, 0, value, -1, nullptr, 0);
    if (required <= 0) {
        return {};
    }
    std::wstring buffer(static_cast<size_t>(required - 1), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value, -1, buffer.data(), required);
    return buffer;
}

std::wstring NormalizePath(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }
#ifdef _WIN32
    const size_t capacity = std::max<size_t>(path.size() + 1, MAX_PATH);
    std::vector<wchar_t> buffer(capacity, L'\0');
    const HRESULT hr = PathCchCanonicalizeEx(buffer.data(), buffer.size(), path.c_str(), PATHCCH_ALLOW_LONG_PATHS);
    if (SUCCEEDED(hr)) {
        const auto terminator = std::find(buffer.begin(), buffer.end(), L'\0');
        return std::wstring(buffer.data(), static_cast<size_t>(std::distance(buffer.begin(), terminator)));
    }
#endif
    return path;
}

bool DirectoryExists(const std::wstring& path) {
    const DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

std::wstring RunGitStatus(const std::wstring& root) {
    const std::wstring command = QuoteCommandArgument(L"git") + L" -C " + QuoteCommandArgument(root) +
                                 L" status --porcelain=2 --branch";
    LogMessage(LogLevel::Info, L"GitStatus invoking git for %ls", root.c_str());
    LogMessage(LogLevel::Info, L"GitStatus command line: %ls", command.c_str());
#ifdef _WIN32
    FILE* pipe = _wpopen(command.c_str(), L"rt, ccs=UTF-8");
#else
    FILE* pipe = popen(std::string(command.begin(), command.end()).c_str(), "r");
#endif
    if (!pipe) {
        LogMessage(LogLevel::Warning, L"GitStatus failed to start git process for %ls (errno=%d)", root.c_str(), errno);
        return {};
    }

    std::wstring result;
    wchar_t buffer[512];
    while (fgetws(buffer, ARRAYSIZE(buffer), pipe)) {
        result.append(buffer);
    }
#ifdef _WIN32
    _pclose(pipe);
#else
    pclose(pipe);
#endif
    LogMessage(LogLevel::Info, L"GitStatus received %llu characters from git for %ls",
               static_cast<unsigned long long>(result.size()), root.c_str());
    return result;
}

GitStatusInfo ParseGitStatus(const std::wstring& root, const std::wstring& output) {
    GitStatusInfo info;
    info.isRepository = true;
    info.rootPath = root;

    std::wistringstream stream(output);
    std::wstring line;
    while (std::getline(stream, line)) {
        if (line.empty()) {
            continue;
        }
        if (line.rfind(L"#", 0) == 0) {
            if (line.find(L"branch.head") != std::wstring::npos) {
                size_t pos = line.find_last_of(L' ');
                if (pos != std::wstring::npos && pos + 1 < line.size()) {
                    info.branch = line.substr(pos + 1);
                }
            } else if (line.find(L"branch.ab") != std::wstring::npos) {
                size_t plus = line.find(L'+');
                size_t minus = line.find(L'-');
                if (plus != std::wstring::npos) {
                    info.ahead = std::wcstol(line.c_str() + plus + 1, nullptr, 10);
                }
                if (minus != std::wstring::npos) {
                    info.behind = std::wcstol(line.c_str() + minus + 1, nullptr, 10);
                }
            }
            continue;
        }

        if (line.size() >= 2) {
            const wchar_t code = line[0];
            if (code == L'?') {
                info.hasUntracked = true;
            }
            if (code != L'!' && code != L'#') {
                info.hasChanges = true;
            }
        }
    }

    return info;
}

}  // namespace

GitStatusCache& GitStatusCache::Instance() {
    static GitStatusCache cache;
    return cache;
}

GitStatusCache::~GitStatusCache() {
    LogMessage(LogLevel::Info, L"GitStatusCache shutting down (workerRunning=%ls)",
               m_workerRunning.load(std::memory_order_acquire) ? L"true" : L"false");
    m_stop.store(true, std::memory_order_release);
    {
        std::lock_guard queueLock(m_queueMutex);
        while (!m_queue.empty()) {
            m_queue.pop();
        }
        m_pendingWork.clear();
    }
    m_queueCv.notify_all();

    {
        std::lock_guard startLock(m_workerStartMutex);
        JoinWorkerIfNeeded();
        m_workerRunning.store(false, std::memory_order_release);
    }
    LogMessage(LogLevel::Info, L"GitStatusCache shutdown complete");
}

GitStatusInfo GitStatusCache::Query(const std::wstring& path) {
    return GuardExplorerCall(
        L"GitStatusCache::Query",
        [&]() -> GitStatusInfo {
            LogMessage(LogLevel::Info, L"GitStatusCache::Query invoked for %ls", path.c_str());
            if (path.empty()) {
                return {};
            }

            const std::wstring repoRoot = ResolveRepositoryRoot(path);
            if (repoRoot.empty()) {
                LogMessage(LogLevel::Info, L"GitStatusCache::Query no repository found for %ls", path.c_str());
                return {};
            }

            EnsureWorker();
            const bool workerReady = m_workerRunning.load(std::memory_order_acquire);

            bool shouldQueue = false;
            GitStatusInfo cached;
            {
                std::lock_guard lock(m_mutex);
                auto& entry = m_cache[repoRoot];
                const auto now = std::chrono::steady_clock::now();
                if (entry.hasStatus && now - entry.timestamp <= kCacheTtl) {
                    LogMessage(LogLevel::Info, L"GitStatusCache::Query returning cached status for %ls",
                               repoRoot.c_str());
                    return entry.status;
                }
                if (workerReady) {
                    if (!entry.inFlight) {
                        entry.inFlight = true;
                        shouldQueue = true;
                    }
                } else {
                    entry.inFlight = false;
                }
                if (entry.hasStatus) {
                    cached = entry.status;
                }
            }

            if (shouldQueue) {
                LogMessage(LogLevel::Info, L"GitStatusCache::Query queueing status refresh for %ls", repoRoot.c_str());
                if (!EnqueueWork(repoRoot)) {
                    LogMessage(LogLevel::Warning, L"GitStatusCache::Query failed to enqueue work for %ls", repoRoot.c_str());
                    std::lock_guard lock(m_mutex);
                    auto it = m_cache.find(repoRoot);
                    if (it != m_cache.end()) {
                        it->second.inFlight = false;
                    }
                }
            } else if (!workerReady) {
                LogMessage(LogLevel::Warning, L"GitStatusCache::Query worker not ready for %ls", repoRoot.c_str());
            }

            if (cached.isRepository) {
                LogMessage(LogLevel::Info, L"GitStatusCache::Query returning stale status for %ls", repoRoot.c_str());
            }
            return cached;
        },
        []() -> GitStatusInfo { return {}; });
}

std::wstring GitStatusCache::ResolveRepositoryRoot(const std::wstring& path) {
    const auto now = std::chrono::steady_clock::now();
    const std::wstring normalized = NormalizePath(path);
    if (normalized.empty() || IsShellNamespacePath(normalized)) {
        return {};
    }

    if (auto cached = LookupCachedRoot(normalized, now)) {
        return *cached;
    }

    std::vector<std::wstring> probed;
    probed.reserve(8);

    std::wstring current = normalized;
    std::wstring resolved;

    while (!current.empty()) {
        probed.push_back(current);

        if (HasGitMetadata(current)) {
            resolved = current;
            break;
        }

        if (auto cachedParent = LookupCachedRoot(current, now)) {
            resolved = *cachedParent;
            break;
        }

        std::wstring parent = ParentDirectory(current);
        if (parent.empty() || parent == current) {
            break;
        }
        current = std::move(parent);
    }

    if (probed.empty()) {
        probed.push_back(normalized);
    }

    CacheRepositoryRoot(probed, resolved, now);
    if (resolved.empty()) {
        LogMessage(LogLevel::Info, L"GitStatusCache::ResolveRepositoryRoot %ls resolved to <none>", normalized.c_str());
    } else {
        LogMessage(LogLevel::Info, L"GitStatusCache::ResolveRepositoryRoot %ls resolved to %ls", normalized.c_str(),
                   resolved.c_str());
    }
    return resolved;
}

GitStatusInfo GitStatusCache::ComputeStatus(const std::wstring& repoRoot) const {
    GitStatusInfo info;
    if (repoRoot.empty()) {
        return info;
    }

    LogMessage(LogLevel::Info, L"GitStatusCache::ComputeStatus running for %ls", repoRoot.c_str());
    const std::wstring output = RunGitStatus(repoRoot);
    if (output.empty()) {
        LogMessage(LogLevel::Warning, L"GitStatusCache::ComputeStatus git output empty for %ls", repoRoot.c_str());
        return info;
    }
    LogMessage(LogLevel::Info, L"GitStatusCache::ComputeStatus received output for %ls", repoRoot.c_str());
    return ParseGitStatus(repoRoot, output);
}

bool GitStatusCache::IsShellNamespacePath(const std::wstring& path) const {
    if (path.empty()) {
        return true;
    }
    if (path.rfind(L"::", 0) == 0 || path.rfind(L"shell::", 0) == 0) {
        return true;
    }
    if (PathIsURLW(path.c_str())) {
        return true;
    }
    if (PathIsRelativeW(path.c_str()) && path.rfind(L"\\\\?\\", 0) != 0) {
        return true;
    }
    return false;
}

bool GitStatusCache::HasGitMetadata(const std::wstring& directory) const {
    if (directory.empty()) {
        return false;
    }

    std::wstring probe = directory;
    if (!probe.empty() && probe.back() != L'\\' && probe.back() != L'/') {
        probe.push_back(L'\\');
    }
    probe += L".git";

    const DWORD attributes = GetFileAttributesW(probe.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES;
}

std::wstring GitStatusCache::ParentDirectory(const std::wstring& path) const {
    if (path.empty()) {
        return {};
    }

    std::wstring buffer = path;
    if (buffer.empty() || buffer.back() != L'\0') {
        buffer.push_back(L'\0');
    }

    PWSTR end = nullptr;
    size_t remaining = 0;
    PathCchRemoveBackslashEx(buffer.data(), buffer.size(), &end, &remaining);

    HRESULT hr = PathCchRemoveFileSpec(buffer.data(), buffer.size());
    if (FAILED(hr)) {
        return {};
    }

    const size_t length = wcsnlen(buffer.c_str(), buffer.size());
    if (length == 0) {
        return {};
    }

    std::wstring parent(buffer.c_str(), length);
    if (parent.size() == 2 && parent[1] == L':') {
        parent.push_back(L'\\');
    }
    return NormalizePath(parent);
}

std::optional<std::wstring> GitStatusCache::LookupCachedRoot(const std::wstring& path,
                                                             std::chrono::steady_clock::time_point now) {
    std::lock_guard lock(m_rootCacheMutex);
    auto it = m_rootCache.find(path);
    if (it == m_rootCache.end()) {
        return std::nullopt;
    }
    if (now - it->second.timestamp > kRootCacheTtl) {
        m_rootCache.erase(it);
        return std::nullopt;
    }
    return it->second.root;
}

void GitStatusCache::CacheRepositoryRoot(const std::vector<std::wstring>& probedPaths, const std::wstring& root,
                                         std::chrono::steady_clock::time_point timestamp) {
    if (probedPaths.empty()) {
        return;
    }

    std::lock_guard lock(m_rootCacheMutex);
    for (const auto& path : probedPaths) {
        if (path.empty()) {
            continue;
        }
        m_rootCache[path] = {root, timestamp};
    }

    for (auto it = m_rootCache.begin(); it != m_rootCache.end();) {
        if (timestamp - it->second.timestamp > kRootCacheTtl) {
            it = m_rootCache.erase(it);
        } else {
            ++it;
        }
    }

    while (m_rootCache.size() > kMaxRootCacheEntries) {
        auto oldest = std::min_element(
            m_rootCache.begin(), m_rootCache.end(),
            [](const auto& lhs, const auto& rhs) { return lhs.second.timestamp < rhs.second.timestamp; });
        if (oldest == m_rootCache.end()) {
            break;
        }
        m_rootCache.erase(oldest);
    }
}

bool GitStatusCache::HasFreshEntry(const std::wstring& repoRoot, GitStatusInfo* info) {
    std::lock_guard lock(m_mutex);
    auto it = m_cache.find(repoRoot);
    if (it == m_cache.end()) {
        return false;
    }
    const auto now = std::chrono::steady_clock::now();
    if (!it->second.hasStatus || now - it->second.timestamp > kCacheTtl) {
        m_cache.erase(it);
        return false;
    }
    if (info) {
        *info = it->second.status;
    }
    return true;
}

size_t GitStatusCache::AddListener(std::function<void()> callback) {
    if (!callback) {
        return 0;
    }
    std::lock_guard lock(m_listenerMutex);
    const size_t id = m_nextListenerId++;
    m_listeners.emplace_back(id, std::move(callback));
    return id;
}

void GitStatusCache::RemoveListener(size_t id) {
    if (id == 0) {
        return;
    }
    std::lock_guard lock(m_listenerMutex);
    auto it = std::remove_if(m_listeners.begin(), m_listeners.end(), [id](const auto& entry) { return entry.first == id; });
    m_listeners.erase(it, m_listeners.end());
}

void GitStatusCache::EnsureWorker() {
    if (m_workerRunning.load(std::memory_order_acquire)) {
        return;
    }

    const auto now = std::chrono::steady_clock::now();
    std::lock_guard startLock(m_workerStartMutex);
    if (m_stop.load(std::memory_order_acquire)) {
        return;
    }

    if (m_workerRunning.load(std::memory_order_relaxed)) {
        return;
    }

    if (m_workerFailed.load(std::memory_order_relaxed) && now < m_nextWorkerRetry) {
        return;
    }

    JoinWorkerIfNeeded();

    try {
        m_stop.store(false, std::memory_order_release);
        m_workerThread = std::thread([this]() { WorkerLoop(); });
        LogMessage(LogLevel::Info, L"GitStatus worker thread started");
        m_workerRunning.store(true, std::memory_order_release);
        m_workerFailed.store(false, std::memory_order_release);
        m_nextWorkerRetry = std::chrono::steady_clock::time_point{};
    } catch (const std::system_error& ex) {
        const std::wstring message = FromUtf8(ex.what());
        if (!message.empty()) {
            LogMessage(LogLevel::Error, L"GitStatus worker start failed (code=%d): %ls", ex.code().value(), message.c_str());
        } else {
            LogMessage(LogLevel::Error, L"GitStatus worker start failed (code=%d)", ex.code().value());
        }
        ScheduleWorkerRetryLocked(now);
    } catch (...) {
        LogMessage(LogLevel::Error, L"GitStatus worker start encountered an unknown exception");
        ScheduleWorkerRetryLocked(now);
    }
}

void GitStatusCache::WorkerLoop() {
    struct RunningGuard {
        explicit RunningGuard(std::atomic<bool>& flag) : target(flag) {}
        ~RunningGuard() {
            if (!released) {
                target.store(false, std::memory_order_release);
            }
        }
        void Release() {
            if (!released) {
                target.store(false, std::memory_order_release);
                released = true;
            }
        }
        std::atomic<bool>& target;
        bool released = false;
    } guard(m_workerRunning);

    auto onFailure = [this]() {
        HandleWorkerFailure(std::chrono::steady_clock::now());
    };

    try {
        while (true) {
            std::wstring repo;
            {
                std::unique_lock lock(m_queueMutex);
                m_queueCv.wait(lock, [this]() { return m_stop.load() || !m_queue.empty(); });
                if (m_stop.load(std::memory_order_acquire) && m_queue.empty()) {
                    LogMessage(LogLevel::Info, L"GitStatus worker stopping gracefully");
                    guard.Release();
                    return;
                }
                repo = std::move(m_queue.front());
                m_queue.pop();
                if (!repo.empty()) {
                    m_pendingWork.erase(repo);
                }
            }

            if (repo.empty()) {
                continue;
            }

            GitStatusInfo info = ComputeStatus(repo);
            const auto now = std::chrono::steady_clock::now();
            bool notify = false;
            {
                std::lock_guard lock(m_mutex);
                auto it = m_cache.find(repo);
                if (it != m_cache.end()) {
                    it->second.status = info;
                    it->second.timestamp = now;
                    it->second.hasStatus = true;
                    it->second.inFlight = false;
                    notify = true;
                }
            }
            if (notify) {
                NotifyListeners();
            }
        }
    } catch (const std::exception& ex) {
        const std::wstring message = FromUtf8(ex.what());
        if (!message.empty()) {
            LogMessage(LogLevel::Error, L"GitStatus worker loop terminated with exception: %ls", message.c_str());
        } else {
            LogMessage(LogLevel::Error, L"GitStatus worker loop terminated with exception");
        }
        guard.Release();
        onFailure();
        return;
    } catch (...) {
        LogMessage(LogLevel::Error, L"GitStatus worker loop terminated with unknown exception");
        guard.Release();
        onFailure();
        return;
    }

    guard.Release();
}

bool GitStatusCache::EnqueueWork(std::wstring repoRoot) {
    if (repoRoot.empty() || m_stop.load(std::memory_order_acquire)) {
        return false;
    }

    bool enqueued = false;
    {
        std::lock_guard lock(m_queueMutex);
        if (m_pendingWork.find(repoRoot) != m_pendingWork.end()) {
            return true;
        }
        if (m_queue.size() >= kMaxQueueDepth) {
            return false;
        }
        m_pendingWork.insert(repoRoot);
        m_queue.push(std::move(repoRoot));
        enqueued = true;
    }
    if (enqueued) {
        m_queueCv.notify_one();
    }
    return enqueued;
}

void GitStatusCache::NotifyListeners() {
    std::vector<std::function<void()>> callbacks;
    {
        std::lock_guard lock(m_listenerMutex);
        callbacks.reserve(m_listeners.size());
        for (const auto& entry : m_listeners) {
            if (entry.second) {
                callbacks.push_back(entry.second);
            }
        }
    }

    for (auto& callback : callbacks) {
        try {
            callback();
        } catch (...) {
        }
    }
}

void GitStatusCache::HandleWorkerFailure(std::chrono::steady_clock::time_point now) {
    LogMessage(LogLevel::Warning, L"GitStatus worker failure detected; scheduling retry");
    ResetPendingWork();

    std::lock_guard startLock(m_workerStartMutex);
    ScheduleWorkerRetryLocked(now);
}

void GitStatusCache::ResetPendingWork() {
    std::vector<std::wstring> pending;
    {
        std::lock_guard lock(m_queueMutex);
        if (m_pendingWork.empty() && m_queue.empty()) {
            return;
        }
        pending.reserve(m_pendingWork.size());
        for (const auto& repo : m_pendingWork) {
            pending.push_back(repo);
        }
        m_pendingWork.clear();
        std::queue<std::wstring> empty;
        std::swap(m_queue, empty);
    }

    if (pending.empty()) {
        LogMessage(LogLevel::Info, L"GitStatus worker had no pending work to reset");
        return;
    }

    std::lock_guard lock(m_mutex);
    for (const auto& repo : pending) {
        auto it = m_cache.find(repo);
        if (it != m_cache.end()) {
            it->second.inFlight = false;
        }
    }
    LogMessage(LogLevel::Info, L"GitStatus worker reset %zu pending repositories", pending.size());
}

void GitStatusCache::ScheduleWorkerRetryLocked(std::chrono::steady_clock::time_point now) {
    m_workerFailed.store(true, std::memory_order_release);
    m_nextWorkerRetry = now + kWorkerRetryDelay;
    const auto delayMs = std::chrono::duration_cast<std::chrono::milliseconds>(kWorkerRetryDelay).count();
    LogMessage(LogLevel::Warning, L"GitStatus worker retry scheduled in %lld ms", static_cast<long long>(delayMs));
}

void GitStatusCache::JoinWorkerIfNeeded() {
    if (!m_workerThread.joinable()) {
        return;
    }
    try {
        LogMessage(LogLevel::Info, L"Joining GitStatus worker thread");
        m_workerThread.join();
        LogMessage(LogLevel::Info, L"GitStatus worker thread joined");
    } catch (...) {
        LogMessage(LogLevel::Warning, L"Exception while joining GitStatus worker thread");
    }
    m_workerThread = std::thread();
}

}  // namespace shelltabs

