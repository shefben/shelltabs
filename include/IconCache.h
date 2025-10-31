#pragma once

#include <windows.h>

#include <functional>
#include <list>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>
#include <vector>

#include <shlobj.h>

namespace shelltabs {

std::wstring BuildIconCacheFamilyKey(PCIDLIST_ABSOLUTE pidl, const std::wstring& canonicalPath);

class IconCache {
private:
    struct Entry;

public:
    class Reference {
    public:
        Reference() = default;
        Reference(const Reference& other);
        Reference& operator=(const Reference& other);
        Reference(Reference&& other) noexcept;
        Reference& operator=(Reference&& other) noexcept;
        ~Reference();

        HICON Get() const noexcept { return m_icon; }
        explicit operator bool() const noexcept { return m_icon != nullptr; }
        void Reset() noexcept;
        std::optional<SIZE> GetMetrics() const noexcept;

    private:
        friend class IconCache;
        Reference(class IconCache* cache, struct IconCache::Entry* entry, HICON icon,
                   bool addRef) noexcept;
        void Attach(class IconCache* cache, struct IconCache::Entry* entry, HICON icon,
                    bool addRef) noexcept;
        void ReleaseCurrent() noexcept;

        class IconCache* m_cache = nullptr;
        struct IconCache::Entry* m_entry = nullptr;
        HICON m_icon = nullptr;
    };

    static IconCache& Instance();

    Reference Acquire(const std::wstring& familyKey, UINT iconFlags, const std::function<HICON()>& loader);
    void InvalidateFamily(const std::wstring& familyKey);
    void InvalidatePidl(PCIDLIST_ABSOLUTE pidl);
    void InvalidatePath(const std::wstring& path);
    void LogStatsNow();

private:
    IconCache() = default;

    struct Entry {
        std::wstring key;
        std::wstring family;
        HICON icon = nullptr;
        SIZE metrics{0, 0};
        size_t refCount = 0;
        std::list<std::wstring>::iterator lruIt;
        bool stale = false;
        bool hasMetrics = false;
    };

    struct StatsSnapshot {
        uint64_t hits = 0;
        uint64_t misses = 0;
        uint64_t evictions = 0;
        size_t size = 0;
    };

    Reference MakeUncachedReference(HICON icon);
    void AddRef(Entry* entry);
    void Release(Entry* entry);
    static std::wstring BuildVariantKey(const std::wstring& familyKey, UINT iconFlags);
    void Touch(Entry* entry);
    std::optional<StatsSnapshot> LogStatsIfNeededLocked();
    void TrimLocked(std::vector<HICON>* destroyList);

    std::unordered_map<std::wstring, std::unique_ptr<Entry>> m_entries;
    std::list<std::wstring> m_lru;
    std::mutex m_mutex;
    size_t m_capacity = 128;
    uint64_t m_hits = 0;
    uint64_t m_misses = 0;
    uint64_t m_evictions = 0;
    uint64_t m_requestsSinceLog = 0;
    uint64_t m_nextToken = 1;
};

}  // namespace shelltabs

