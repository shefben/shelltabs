#pragma once

#include <windows.h>

#include <chrono>
#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>

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

private:
    struct CacheEntry {
        GitStatusInfo status;
        std::chrono::steady_clock::time_point timestamp;
    };

    GitStatusCache() = default;

    std::wstring FindRepositoryRoot(const std::wstring& path) const;
    GitStatusInfo ComputeStatus(const std::wstring& repoRoot) const;
    bool HasFreshEntry(const std::wstring& repoRoot, GitStatusInfo* info);

    mutable std::mutex m_mutex;
    std::unordered_map<std::wstring, CacheEntry> m_cache;
};

}  // namespace shelltabs

