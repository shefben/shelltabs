#include "BackgroundCache.h"

#include "Logging.h"
#include "Utilities.h"

#include <Shlwapi.h>
#include <objbase.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cwctype>
#include <limits>
#include <memory>
#include <mutex>
#include <string>
#include <system_error>
#include <utility>
#include <unordered_set>
#include <thread>
#include <vector>

namespace shelltabs {
namespace {

constexpr auto kUnusedCacheExpiration = std::chrono::hours(24 * 30);
constexpr auto kCacheMaintenanceThrottle = std::chrono::minutes(5);

std::atomic<ULONGLONG> g_lastCachePurgeTick{0};
std::atomic<bool> g_cachePurgeInFlight{false};
std::mutex g_cachePurgeMutex;

struct CacheMaintenanceTask {
    std::unordered_set<std::wstring> referenced;
    ULONGLONG expirationTicks = 0;
};

void PurgeUnusedCacheEntries(const std::unordered_set<std::wstring>& referenced,
                             ULONGLONG expirationTicks);

ULONGLONG ThrottleIntervalInMilliseconds() {
    return static_cast<ULONGLONG>(
        std::chrono::duration_cast<std::chrono::milliseconds>(kCacheMaintenanceThrottle).count());
}

void RunCacheMaintenanceNow(const std::unordered_set<std::wstring>& referenced, ULONGLONG expirationTicks) {
    std::lock_guard<std::mutex> lock(g_cachePurgeMutex);
    PurgeUnusedCacheEntries(referenced, expirationTicks);
    g_lastCachePurgeTick.store(GetTickCount64(), std::memory_order_relaxed);
}

void DispatchCacheMaintenance(std::unordered_set<std::wstring> referenced,
                              ULONGLONG expirationTicks,
                              bool forceMaintenance) {
    if (forceMaintenance) {
        RunCacheMaintenanceNow(referenced, expirationTicks);
        return;
    }

    const ULONGLONG throttleMillis = ThrottleIntervalInMilliseconds();
    if (throttleMillis > 0) {
        const ULONGLONG now = GetTickCount64();
        const ULONGLONG last = g_lastCachePurgeTick.load(std::memory_order_relaxed);
        if (last != 0) {
            const ULONGLONG elapsed = now >= last ? (now - last) : std::numeric_limits<ULONGLONG>::max();
            if (elapsed < throttleMillis) {
                return;
            }
        }
    }

    bool expected = false;
    if (!g_cachePurgeInFlight.compare_exchange_strong(expected, true, std::memory_order_acq_rel)) {
        return;
    }

    auto task = std::make_shared<CacheMaintenanceTask>();
    task->referenced = std::move(referenced);
    task->expirationTicks = expirationTicks;

    struct FlagResetter {
        ~FlagResetter() { g_cachePurgeInFlight.store(false, std::memory_order_release); }
    };

    auto worker = [task]() {
        FlagResetter flagResetter;

        RunCacheMaintenanceNow(task->referenced, task->expirationTicks);
    };

    try {
        std::thread(worker).detach();
    } catch (const std::system_error& ex) {
        LogMessage(LogLevel::Warning,
                   L"Failed to queue background cache purge (%ls); running synchronously.",
                   Utf8ToWide(ex.what()).c_str());

        FlagResetter flagResetter;

        RunCacheMaintenanceNow(task->referenced, task->expirationTicks);
    }
}

std::wstring NormalizeAndEnsureTrailingSlash(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }

    if (normalized.back() != L'\\') {
        normalized.push_back(L'\\');
    }
    return normalized;
}

std::wstring ToCaseInsensitiveKey(const std::wstring& path) {
    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }
    std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(towlower(ch));
    });
    return normalized;
}

std::wstring FormatSystemErrorMessage(DWORD error) {
    if (error == 0) {
        return {};
    }
    wchar_t* buffer = nullptr;
    DWORD copied = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
                                      FORMAT_MESSAGE_IGNORE_INSERTS,
                                  nullptr, error, 0, reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
    std::wstring message;
    if (copied != 0 && buffer) {
        message.assign(buffer, copied);
        LocalFree(buffer);
        while (!message.empty() && (message.back() == L'\r' || message.back() == L'\n')) {
            message.pop_back();
        }
    } else {
        message = L"Error code " + std::to_wstring(error);
    }
    return message;
}

std::vector<std::wstring> CollectReferencedPaths(const ShellTabsOptions& options) {
    std::vector<std::wstring> paths;
    if (!options.universalFolderBackgroundImage.cachedImagePath.empty()) {
        paths.push_back(options.universalFolderBackgroundImage.cachedImagePath);
    }
    for (const auto& entry : options.folderBackgroundEntries) {
        if (!entry.image.cachedImagePath.empty()) {
            paths.push_back(entry.image.cachedImagePath);
        }
    }
    return paths;
}

ULONGLONG DurationToFileTimeTicks(const std::chrono::milliseconds& duration) {
    return static_cast<ULONGLONG>(duration.count()) * 10'000ULL;
}

ULONGLONG ExpirationToTicks(const std::chrono::hours& expiration) {
    return DurationToFileTimeTicks(std::chrono::duration_cast<std::chrono::milliseconds>(expiration));
}

void PurgeUnusedCacheEntries(const std::unordered_set<std::wstring>& referenced,
                             ULONGLONG expirationTicks) {
    std::wstring cacheDirectory = EnsureBackgroundCacheDirectory();
    if (cacheDirectory.empty()) {
        return;
    }

    std::wstring normalizedDirectory = NormalizeAndEnsureTrailingSlash(cacheDirectory);
    if (normalizedDirectory.empty()) {
        return;
    }

    std::wstring searchPattern = normalizedDirectory + L"*";

    WIN32_FIND_DATAW findData{};
    HANDLE find = FindFirstFileW(searchPattern.c_str(), &findData);
    if (find == INVALID_HANDLE_VALUE) {
        return;
    }

    FILETIME nowFileTime{};
    GetSystemTimeAsFileTime(&nowFileTime);
    ULARGE_INTEGER now{};
    now.LowPart = nowFileTime.dwLowDateTime;
    now.HighPart = nowFileTime.dwHighDateTime;

    do {
        if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            continue;
        }

        std::wstring filePath = normalizedDirectory + findData.cFileName;
        std::wstring key = ToCaseInsensitiveKey(filePath);
        if (key.empty() || referenced.find(key) != referenced.end()) {
            continue;
        }

        ULARGE_INTEGER lastWrite{};
        lastWrite.LowPart = findData.ftLastWriteTime.dwLowDateTime;
        lastWrite.HighPart = findData.ftLastWriteTime.dwHighDateTime;

        bool expired = expirationTicks == 0;
        if (!expired && now.QuadPart >= lastWrite.QuadPart) {
            expired = (now.QuadPart - lastWrite.QuadPart) >= expirationTicks;
        }

        if (!expired) {
            continue;
        }

        if (DeleteFileW(filePath.c_str())) {
            LogMessage(LogLevel::Info, L"Removed stale cached background image %ls", filePath.c_str());
        } else {
            LogLastError(L"DeleteFileW(background cache purge)", GetLastError());
        }
    } while (FindNextFileW(find, &findData));

    FindClose(find);
}

bool IsPathInDirectory(const std::wstring& path, const std::wstring& directory) {
    if (path.size() < directory.size()) {
        return false;
    }
    return _wcsnicmp(path.c_str(), directory.c_str(), directory.size()) == 0;
}

}  // namespace

std::wstring EnsureBackgroundCacheDirectory() {
    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return {};
    }

    if (directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += L"Backgrounds";
    CreateDirectoryW(directory.c_str(), nullptr);
    return NormalizeFileSystemPath(directory);
}

bool CopyImageToBackgroundCache(const std::wstring& sourcePath,
                                const std::wstring& displayName,
                                CachedImageMetadata* metadata,
                                std::wstring* createdPath,
                                std::wstring* errorMessage) {
    if (errorMessage) {
        errorMessage->clear();
    }

    if (!metadata) {
        if (errorMessage) {
            *errorMessage = L"Invalid image metadata destination.";
        }
        return false;
    }

    if (createdPath) {
        createdPath->clear();
    }

    std::wstring normalizedSource = NormalizeFileSystemPath(sourcePath);
    if (normalizedSource.empty()) {
        if (errorMessage) {
            *errorMessage = L"The selected image path could not be normalized.";
        }
        return false;
    }

    std::wstring cacheDirectory = EnsureBackgroundCacheDirectory();
    if (cacheDirectory.empty()) {
        if (errorMessage) {
            *errorMessage = L"Unable to resolve the ShellTabs cache directory.";
        }
        return false;
    }

    std::wstring normalizedCacheDirectory = NormalizeAndEnsureTrailingSlash(cacheDirectory);
    if (normalizedCacheDirectory.empty()) {
        if (errorMessage) {
            *errorMessage = L"Unable to normalize the ShellTabs cache directory.";
        }
        return false;
    }

    std::wstring targetPath = normalizedSource;
    bool copiedIntoCache = false;

    if (!IsPathInDirectory(targetPath, normalizedCacheDirectory)) {
        GUID guid{};
        if (FAILED(CoCreateGuid(&guid))) {
            if (errorMessage) {
                *errorMessage = L"Unable to create a cache identifier for the image.";
            }
            return false;
        }

        wchar_t guidBuffer[64];
        if (StringFromGUID2(guid, guidBuffer, ARRAYSIZE(guidBuffer)) <= 0) {
            if (errorMessage) {
                *errorMessage = L"Unable to create a cache filename for the image.";
            }
            return false;
        }

        std::wstring fileName(guidBuffer);
        fileName.erase(std::remove(fileName.begin(), fileName.end(), L'{'), fileName.end());
        fileName.erase(std::remove(fileName.begin(), fileName.end(), L'}'), fileName.end());
        fileName.erase(std::remove(fileName.begin(), fileName.end(), L'-'), fileName.end());

        const wchar_t* extension = PathFindExtensionW(normalizedSource.c_str());
        std::wstring ext = (extension && *extension) ? extension : L".img";

        if (normalizedCacheDirectory.back() != L'\\') {
            normalizedCacheDirectory.push_back(L'\\');
        }
        targetPath = normalizedCacheDirectory + fileName + ext;

        if (!CopyFileW(normalizedSource.c_str(), targetPath.c_str(), FALSE)) {
            DWORD copyError = GetLastError();
            DeleteFileW(targetPath.c_str());
            if (errorMessage) {
                *errorMessage = FormatSystemErrorMessage(copyError);
            }
            LogLastError(L"CopyFileW(background cache)", copyError);
            return false;
        }

        copiedIntoCache = true;
        if (createdPath) {
            *createdPath = targetPath;
        }
    }

    std::wstring normalizedTarget = NormalizeFileSystemPath(targetPath);
    if (normalizedTarget.empty()) {
        if (copiedIntoCache) {
            DeleteFileW(targetPath.c_str());
        }
        if (errorMessage) {
            *errorMessage = L"Unable to normalize the cached image path.";
        }
        return false;
    }

    metadata->cachedImagePath = normalizedTarget;
    metadata->displayName = !displayName.empty() ? displayName :
                                                       (PathFindFileNameW(normalizedSource.c_str())
                                                            ? PathFindFileNameW(normalizedSource.c_str())
                                                            : normalizedSource);

    if (createdPath && !createdPath->empty()) {
        *createdPath = normalizedTarget;
    }

    TouchCachedImage(normalizedTarget);
    return true;
}

std::unique_ptr<Gdiplus::Bitmap> LoadBackgroundBitmap(const std::wstring& path) {
    if (path.empty()) {
        return nullptr;
    }

    auto bitmap = std::make_unique<Gdiplus::Bitmap>(path.c_str());
    if (!bitmap || bitmap->GetLastStatus() != Gdiplus::Ok) {
        return nullptr;
    }

    TouchCachedImage(path);
    return bitmap;
}

void TouchCachedImage(const std::wstring& path) noexcept {
    if (path.empty()) {
        return;
    }

    HANDLE file = CreateFileW(path.c_str(), FILE_WRITE_ATTRIBUTES,
                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING,
                              FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return;
    }

    FILETIME now{};
    GetSystemTimeAsFileTime(&now);
    SetFileTime(file, nullptr, &now, &now);
    CloseHandle(file);
}

std::vector<std::wstring> CollectCachedImageReferences(const ShellTabsOptions& options) {
    return CollectReferencedPaths(options);
}

void UpdateCachedImageUsage(const ShellTabsOptions& options, bool forceMaintenance) {
    const std::vector<std::wstring> references = CollectReferencedPaths(options);
    if (references.empty()) {
        // Still perform an expiration sweep so abandoned files disappear eventually.
        DispatchCacheMaintenance({}, ExpirationToTicks(kUnusedCacheExpiration), forceMaintenance);
        return;
    }

    std::unordered_set<std::wstring> referencedKeys;
    referencedKeys.reserve(references.size());
    for (const auto& path : references) {
        TouchCachedImage(path);
        std::wstring key = ToCaseInsensitiveKey(path);
        if (!key.empty()) {
            referencedKeys.insert(std::move(key));
        }
    }

    DispatchCacheMaintenance(std::move(referencedKeys), ExpirationToTicks(kUnusedCacheExpiration),
                             forceMaintenance);
}

void ForceBackgroundCacheMaintenance(const ShellTabsOptions& options) {
    UpdateCachedImageUsage(options, true);
}

CacheMaintenanceResult RemoveOrphanedCacheEntries(const ShellTabsOptions& options,
                                                  const std::vector<std::wstring>& protectedPaths) {
    CacheMaintenanceResult result;

    const std::vector<std::wstring> persistedReferences = CollectReferencedPaths(options);
    std::unordered_set<std::wstring> referencedKeys;
    referencedKeys.reserve(persistedReferences.size() + protectedPaths.size());

    for (const auto& path : persistedReferences) {
        std::wstring key = ToCaseInsensitiveKey(path);
        if (!key.empty()) {
            referencedKeys.insert(std::move(key));
        }
    }
    for (const auto& path : protectedPaths) {
        std::wstring key = ToCaseInsensitiveKey(path);
        if (!key.empty()) {
            referencedKeys.insert(std::move(key));
        }
    }

    std::wstring cacheDirectory = EnsureBackgroundCacheDirectory();
    if (cacheDirectory.empty()) {
        return result;
    }

    std::wstring normalizedDirectory = NormalizeAndEnsureTrailingSlash(cacheDirectory);
    if (normalizedDirectory.empty()) {
        return result;
    }

    std::wstring searchPattern = normalizedDirectory + L"*";
    WIN32_FIND_DATAW findData{};
    HANDLE find = FindFirstFileW(searchPattern.c_str(), &findData);
    if (find == INVALID_HANDLE_VALUE) {
        return result;
    }

    do {
        if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            continue;
        }
        std::wstring filePath = normalizedDirectory + findData.cFileName;
        std::wstring key = ToCaseInsensitiveKey(filePath);
        if (key.empty() || referencedKeys.find(key) != referencedKeys.end()) {
            continue;
        }

        if (DeleteFileW(filePath.c_str())) {
            result.removedPaths.push_back(filePath);
        } else {
            CacheMaintenanceFailure failure;
            failure.path = filePath;
            failure.error = GetLastError();
            failure.message = FormatSystemErrorMessage(failure.error);
            result.failures.push_back(std::move(failure));
        }
    } while (FindNextFileW(find, &findData));

    FindClose(find);
    return result;
}

}  // namespace shelltabs

