#pragma once

#include <atomic>
#include <condition_variable>
#include <deque>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>
#include <thread>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>

#include <shlobj.h>

namespace shelltabs {

struct PreviewImage {
    HBITMAP bitmap = nullptr;
    SIZE size{};
};

inline constexpr SIZE kPreviewImageSize{192, 128};

// Provides cached previews for PIDLs captured from Explorer folder views.
class PreviewCache {
public:
    static PreviewCache& Instance();

    std::optional<PreviewImage> GetPreview(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize);
    void StorePreviewFromWindow(PCIDLIST_ABSOLUTE pidl, HWND window, const SIZE& desiredSize);
    uint64_t RequestPreviewAsync(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize, HWND notifyHwnd, UINT message);
    void CancelRequest(uint64_t requestId);
    void Clear();

private:
    struct AsyncRequest;

    PreviewCache() = default;
    ~PreviewCache();

    PreviewCache(const PreviewCache&) = delete;
    PreviewCache& operator=(const PreviewCache&) = delete;

    struct Entry {
        HBITMAP bitmap = nullptr;
        SIZE size{};
        ULONGLONG lastAccess = 0;
    };

    void EnsureWorkerThread();
    void ProcessRequests();
    void StoreBitmapForKey(const std::wstring& key, HBITMAP bitmap, const SIZE& size);
    void TrimCacheLocked();
    static std::wstring BuildCacheKey(PCIDLIST_ABSOLUTE pidl);

    std::mutex m_mutex;
    std::unordered_map<std::wstring, Entry> m_entries;
    static constexpr size_t kMaxEntries = 64;

    std::mutex m_requestMutex;
    std::condition_variable m_requestCv;
    std::deque<std::shared_ptr<AsyncRequest>> m_requestQueue;
    std::unordered_map<uint64_t, std::shared_ptr<AsyncRequest>> m_requestMap;
    std::thread m_workerThread;
    bool m_shutdown = false;
    uint64_t m_nextRequestId = 1;
};

}  // namespace shelltabs

