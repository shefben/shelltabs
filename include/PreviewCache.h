#pragma once

#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>

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
    void Clear();

private:
    PreviewCache() = default;
    ~PreviewCache();

    PreviewCache(const PreviewCache&) = delete;
    PreviewCache& operator=(const PreviewCache&) = delete;

    struct Entry {
        HBITMAP bitmap = nullptr;
        SIZE size{};
        ULONGLONG lastAccess = 0;
    };

    void TrimCacheLocked();
    static std::wstring BuildCacheKey(PCIDLIST_ABSOLUTE pidl);

    std::mutex m_mutex;
    std::unordered_map<std::wstring, Entry> m_entries;
    static constexpr size_t kMaxEntries = 64;
};

}  // namespace shelltabs

