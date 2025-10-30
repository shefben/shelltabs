#include "PreviewCache.h"

#include <algorithm>
#include <cmath>
#include <mutex>
#include <optional>
#include <vector>

#include <shlwapi.h>
#include <shobjidl_core.h>
#include <wrl/client.h>

#include "Utilities.h"

namespace shelltabs {

namespace {
constexpr SIZE kDefaultPreviewSize{160, 120};

HBITMAP CreatePreviewBitmap(IShellItemImageFactory* factory, const SIZE& desiredSize) {
    if (!factory) {
        return nullptr;
    }

    SIZE size = desiredSize;
    if (size.cx <= 0 || size.cy <= 0) {
        size = kDefaultPreviewSize;
    }

    HBITMAP bitmap = nullptr;
    constexpr DWORD kFlags = SIIGBF_RESIZETOFIT | SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK;
    if (FAILED(factory->GetImage(size, kFlags, &bitmap))) {
        return nullptr;
    }
    return bitmap;
}

SIZE GetBitmapSize(HBITMAP bitmap) {
    SIZE size{0, 0};
    if (!bitmap) {
        return size;
    }

    BITMAP info{};
    if (GetObjectW(bitmap, sizeof(info), &info) == sizeof(info)) {
        size.cx = info.bmWidth;
        size.cy = info.bmHeight;
    }
    return size;
}

}  // namespace

PreviewCache& PreviewCache::Instance() {
    static PreviewCache cache;
    return cache;
}

PreviewCache::~PreviewCache() {
    Clear();
}

std::optional<PreviewImage> PreviewCache::GetPreview(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize) {
    if (!pidl) {
        return std::nullopt;
    }

    const std::wstring key = BuildCacheKey(pidl);
    if (key.empty()) {
        return std::nullopt;
    }

    {
        std::scoped_lock lock(m_mutex);
        auto it = m_entries.find(key);
        if (it != m_entries.end()) {
            it->second.lastAccess = GetTickCount64();
            return PreviewImage{it->second.bitmap, it->second.size};
        }
    }

    Microsoft::WRL::ComPtr<IShellItem> item;
    if (FAILED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&item)))) {
        return std::nullopt;
    }

    Microsoft::WRL::ComPtr<IShellItemImageFactory> factory;
    if (FAILED(item.As(&factory))) {
        return std::nullopt;
    }

    HBITMAP bitmap = CreatePreviewBitmap(factory.Get(), desiredSize);
    if (!bitmap) {
        return std::nullopt;
    }

    SIZE size = GetBitmapSize(bitmap);
    if (size.cx <= 0 || size.cy <= 0) {
        DeleteObject(bitmap);
        return std::nullopt;
    }

    {
        std::scoped_lock lock(m_mutex);
        Entry entry;
        entry.bitmap = bitmap;
        entry.size = size;
        entry.lastAccess = GetTickCount64();
        m_entries.emplace(key, entry);
        TrimCacheLocked();
    }

    return PreviewImage{bitmap, size};
}

void PreviewCache::Clear() {
    std::scoped_lock lock(m_mutex);
    for (auto& [_, entry] : m_entries) {
        if (entry.bitmap) {
            DeleteObject(entry.bitmap);
            entry.bitmap = nullptr;
        }
    }
    m_entries.clear();
}

void PreviewCache::TrimCacheLocked() {
    if (m_entries.size() <= kMaxEntries) {
        return;
    }

    std::vector<std::pair<std::wstring, Entry*>> ordered;
    ordered.reserve(m_entries.size());
    for (auto& [key, entry] : m_entries) {
        ordered.emplace_back(key, &entry);
    }

    std::sort(ordered.begin(), ordered.end(), [](const auto& left, const auto& right) {
        return left.second->lastAccess < right.second->lastAccess;
    });

    while (m_entries.size() > kMaxEntries && !ordered.empty()) {
        const auto& victim = ordered.front();
        auto it = m_entries.find(victim.first);
        if (it != m_entries.end()) {
            if (it->second.bitmap) {
                DeleteObject(it->second.bitmap);
            }
            m_entries.erase(it);
        }
        ordered.erase(ordered.begin());
    }
}

std::wstring PreviewCache::BuildCacheKey(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return {};
    }
    std::wstring key = GetParsingName(pidl);
    if (!key.empty()) {
        return key;
    }
    wchar_t buffer[32];
    _snwprintf_s(buffer, _TRUNCATE, L"pidl:%p", pidl);
    return buffer;
}

}  // namespace shelltabs

