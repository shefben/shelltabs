#include "GitStatus.h"

#include <Shlwapi.h>

#include <cstdio>
#include <exception>
#include <filesystem>
#include <sstream>
#include <string>
#include <vector>

#include "Utilities.h"

namespace shelltabs {

namespace {
constexpr std::chrono::seconds kCacheTtl(5);

std::wstring NormalizePath(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }
    wchar_t buffer[MAX_PATH];
    if (PathCanonicalizeW(buffer, path.c_str())) {
        return std::wstring(buffer);
    }
    return path;
}

bool DirectoryExists(const std::wstring& path) {
    const DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

std::wstring RunGitStatus(const std::wstring& root) {
    std::wstring command = L"\"git\" -C \"" + root + L"\" status --porcelain=2 --branch";
#ifdef _WIN32
    FILE* pipe = _wpopen(command.c_str(), L"rt, ccs=UTF-8");
#else
    FILE* pipe = popen(std::string(command.begin(), command.end()).c_str(), "r");
#endif
    if (!pipe) {
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

GitStatusInfo GitStatusCache::Query(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    const std::wstring repoRoot = FindRepositoryRoot(path);
    if (repoRoot.empty()) {
        return {};
    }

    GitStatusInfo cached;
    if (HasFreshEntry(repoRoot, &cached)) {
        return cached;
    }

    GitStatusInfo info = ComputeStatus(repoRoot);

    std::lock_guard lock(m_mutex);
    m_cache[repoRoot] = CacheEntry{info, std::chrono::steady_clock::now()};
    return info;
}

std::wstring GitStatusCache::FindRepositoryRoot(const std::wstring& path) const {
    try {
        std::filesystem::path current(path);
        if (current.empty()) {
            return {};
        }

        if (!std::filesystem::exists(current)) {
            current = current.parent_path();
        }

        while (!current.empty()) {
            const std::wstring gitDir = (current / L".git").wstring();
            if (DirectoryExists(gitDir)) {
                return NormalizePath(current.wstring());
            }
            current = current.parent_path();
        }
    } catch (const std::exception&) {
        return {};
    }
    return {};
}

GitStatusInfo GitStatusCache::ComputeStatus(const std::wstring& repoRoot) const {
    GitStatusInfo info;
    if (repoRoot.empty()) {
        return info;
    }

    const std::wstring output = RunGitStatus(repoRoot);
    if (output.empty()) {
        return info;
    }
    return ParseGitStatus(repoRoot, output);
}

bool GitStatusCache::HasFreshEntry(const std::wstring& repoRoot, GitStatusInfo* info) {
    std::lock_guard lock(m_mutex);
    auto it = m_cache.find(repoRoot);
    if (it == m_cache.end()) {
        return false;
    }
    const auto now = std::chrono::steady_clock::now();
    if (now - it->second.timestamp > kCacheTtl) {
        m_cache.erase(it);
        return false;
    }
    if (info) {
        *info = it->second.status;
    }
    return true;
}

}  // namespace shelltabs

