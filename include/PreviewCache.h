#pragma once

#include <atomic>
#include <condition_variable>
#include <deque>
#include <cstdint>
#include <list>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <vector>

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
    void StorePreviewFromWindow(PCIDLIST_ABSOLUTE pidl, HWND window, const SIZE& desiredSize,
                                std::wstring_view ownerToken = {});
    uint64_t RequestPreviewAsync(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize, HWND notifyHwnd, UINT message);
    void CancelRequest(uint64_t requestId);
    void CancelPendingCapturesForKey(PCIDLIST_ABSOLUTE pidl);
    void CancelPendingCapturesForOwner(std::wstring_view ownerToken);
    void Clear();

private:
    enum class RequestKind {
        kShellPreview,
        kWindowCapture,
    };

    struct AsyncRequest;

    struct PendingKeyEntry {
        uint64_t shellPreviewId = 0;
        uint64_t windowCaptureId = 0;

        [[nodiscard]] bool Empty() const {
            return shellPreviewId == 0 && windowCaptureId == 0;
        }
    };

    PreviewCache() = default;
    ~PreviewCache();

    PreviewCache(const PreviewCache&) = delete;
    PreviewCache& operator=(const PreviewCache&) = delete;

    struct Entry {
        HBITMAP bitmap = nullptr;
        SIZE size{};
        std::list<std::wstring>::iterator lruPosition{};
        bool inLruList = false;
    };

    void EnsureWorkerThread();
    void ProcessRequests();
    void StoreBitmapForKey(const std::wstring& key, HBITMAP bitmap, const SIZE& size);
    void TouchEntryLocked(Entry& entry, const std::wstring& key);
    void TrimCacheLocked();
    static std::wstring BuildCacheKey(PCIDLIST_ABSOLUTE pidl);
    uint64_t GetPendingRequestIdLocked(const std::wstring& key, RequestKind kind);
    void SetPendingRequestIdLocked(const std::wstring& key, RequestKind kind, uint64_t requestId);
    void ClearPendingRequestIdLocked(const std::wstring& key, RequestKind kind, uint64_t requestId);

    std::mutex m_mutex;
    std::unordered_map<std::wstring, Entry> m_entries;
    std::list<std::wstring> m_lruList;
    static constexpr size_t kMaxEntries = 64;

    std::mutex m_requestMutex;
    std::condition_variable m_requestCv;
    std::deque<std::shared_ptr<AsyncRequest>> m_requestQueue;
    std::unordered_map<uint64_t, std::shared_ptr<AsyncRequest>> m_requestMap;
    std::unordered_map<std::wstring, PendingKeyEntry> m_requestsByKey;
    std::thread m_workerThread;
    bool m_shutdown = false;
    uint64_t m_nextRequestId = 1;
};

}  // namespace shelltabs

