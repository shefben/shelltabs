#include "IconCache.h"

#include "Logging.h"
#include "Utilities.h"

#include <algorithm>
#include <cwctype>
#include <sstream>

namespace shelltabs {

namespace {
constexpr uint64_t kLogInterval = 50;

size_t HashPidlBytes(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return 0;
    }
    const BYTE* data = reinterpret_cast<const BYTE*>(pidl);
    size_t hash = 1469598103934665603ull;
    for (;;) {
        auto* item = reinterpret_cast<const SHITEMID*>(data);
        if (!item || item->cb == 0) {
            break;
        }
        for (USHORT i = 0; i < item->cb; ++i) {
            hash ^= data[i];
            hash *= 1099511628211ull;
        }
        data += item->cb;
    }
    return hash;
}

std::wstring NormalizeKey(std::wstring key) {
    if (key.empty()) {
        return key;
    }
    std::wstring result = key;
    if (IsLikelyFileSystemPath(result)) {
        std::wstring normalized = NormalizeFileSystemPath(result);
        if (!normalized.empty()) {
            result = std::move(normalized);
        }
    }
    std::transform(result.begin(), result.end(), result.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(ch));
    });
    return result;
}

}  // namespace

std::wstring BuildIconCacheFamilyKey(PCIDLIST_ABSOLUTE pidl, const std::wstring& canonicalPath) {
    std::wstring key = canonicalPath;
    if (key.empty() && pidl) {
        key = GetCanonicalParsingName(pidl);
        if (key.empty()) {
            key = GetParsingName(pidl);
        }
    }
    if (key.empty() && pidl) {
        std::wstringstream stream;
        stream << L"pidl:" << std::hex << HashPidlBytes(pidl);
        key = stream.str();
    }
    if (key.empty()) {
        return {};
    }
    return NormalizeKey(std::move(key));
}

IconCache& IconCache::Instance() {
    static IconCache cache;
    return cache;
}

IconCache::Reference::Reference(const Reference& other) {
    Attach(other.m_cache, other.m_entry, other.m_icon, true);
}

IconCache::Reference& IconCache::Reference::operator=(const Reference& other) {
    if (this != &other) {
        ReleaseCurrent();
        Attach(other.m_cache, other.m_entry, other.m_icon, true);
    }
    return *this;
}

IconCache::Reference::Reference(Reference&& other) noexcept {
    m_cache = other.m_cache;
    m_entry = other.m_entry;
    m_icon = other.m_icon;
    other.m_cache = nullptr;
    other.m_entry = nullptr;
    other.m_icon = nullptr;
}

IconCache::Reference& IconCache::Reference::operator=(Reference&& other) noexcept {
    if (this != &other) {
        ReleaseCurrent();
        m_cache = other.m_cache;
        m_entry = other.m_entry;
        m_icon = other.m_icon;
        other.m_cache = nullptr;
        other.m_entry = nullptr;
        other.m_icon = nullptr;
    }
    return *this;
}

IconCache::Reference::~Reference() { ReleaseCurrent(); }

void IconCache::Reference::Reset() noexcept { ReleaseCurrent(); }

IconCache::Reference::Reference(IconCache* cache, Entry* entry, HICON icon, bool addRef) noexcept {
    Attach(cache, entry, icon, addRef);
}

void IconCache::Reference::Attach(IconCache* cache, Entry* entry, HICON icon, bool addRef) noexcept {
    m_cache = cache;
    m_entry = entry;
    m_icon = icon;
    if (addRef && m_cache && m_entry) {
        m_cache->AddRef(m_entry);
    }
}

void IconCache::Reference::ReleaseCurrent() noexcept {
    if (m_cache && m_entry) {
        m_cache->Release(m_entry);
    } else if (m_icon) {
        DestroyIcon(m_icon);
    }
    m_cache = nullptr;
    m_entry = nullptr;
    m_icon = nullptr;
}

IconCache::Reference IconCache::MakeUncachedReference(HICON icon) {
    return Reference(nullptr, nullptr, icon, false);
}

IconCache::Reference IconCache::Acquire(const std::wstring& familyKey, UINT iconFlags,
                                        const std::function<HICON()>& loader) {
    if (!loader) {
        return {};
    }

    const UINT variantFlags = (iconFlags & (SHGFI_LARGEICON | SHGFI_SMALLICON)) != 0
                                  ? (iconFlags & (SHGFI_LARGEICON | SHGFI_SMALLICON))
                                  : SHGFI_SMALLICON;

    if (familyKey.empty()) {
        return MakeUncachedReference(loader());
    }

    const std::wstring variantKey = BuildVariantKey(familyKey, variantFlags);

    std::vector<HICON> destroyList;
    std::optional<StatsSnapshot> snapshot;
    Entry* entry = nullptr;
    {
        std::unique_lock lock(m_mutex);
        auto it = m_entries.find(variantKey);
        if (it != m_entries.end()) {
            entry = it->second.get();
            ++entry->refCount;
            Touch(entry);
            ++m_hits;
            ++m_requestsSinceLog;
            snapshot = LogStatsIfNeededLocked();
        } else {
            ++m_misses;
            ++m_requestsSinceLog;
            lock.unlock();
            HICON icon = loader();
            if (!icon) {
                return {};
            }
            lock.lock();
            auto retry = m_entries.find(variantKey);
            if (retry != m_entries.end()) {
                destroyList.push_back(icon);
                entry = retry->second.get();
                ++entry->refCount;
                Touch(entry);
            } else {
                auto newEntry = std::make_unique<Entry>();
                newEntry->key = variantKey;
                newEntry->family = familyKey;
                newEntry->icon = icon;
                newEntry->refCount = 1;
                newEntry->lruIt = m_lru.insert(m_lru.begin(), variantKey);
                entry = newEntry.get();
                m_entries.emplace(variantKey, std::move(newEntry));
                TrimLocked(&destroyList);
            }
            snapshot = LogStatsIfNeededLocked();
        }
    }

    for (HICON icon : destroyList) {
        if (icon) {
            DestroyIcon(icon);
        }
    }

    if (snapshot.has_value()) {
        const double total = static_cast<double>(snapshot->hits + snapshot->misses);
        const double hitRate = total > 0.0 ? (static_cast<double>(snapshot->hits) / total) * 100.0 : 0.0;
        LogMessage(LogLevel::Info,
                   L"IconCache stats: size=%zu hits=%llu misses=%llu evictions=%llu hitRate=%.1f%%",
                   snapshot->size, static_cast<unsigned long long>(snapshot->hits),
                   static_cast<unsigned long long>(snapshot->misses),
                   static_cast<unsigned long long>(snapshot->evictions), hitRate);
    }

    if (!entry) {
        return {};
    }

    return Reference(this, entry, entry->icon, false);
}

void IconCache::InvalidateFamily(const std::wstring& familyKey) {
    if (familyKey.empty()) {
        return;
    }

    std::vector<HICON> destroyList;
    {
        std::unique_lock lock(m_mutex);
        for (auto it = m_entries.begin(); it != m_entries.end();) {
            Entry* entry = it->second.get();
            if (!entry || entry->family != familyKey) {
                ++it;
                continue;
            }
            if (entry->refCount > 0) {
                entry->stale = true;
                std::wstring newKey = entry->key + L"#stale" + std::to_wstring(m_nextToken++);
                if (entry->lruIt != m_lru.end()) {
                    *entry->lruIt = newKey;
                }
                auto owned = std::move(it->second);
                it = m_entries.erase(it);
                entry->key = std::move(newKey);
                it = m_entries.emplace(entry->key, std::move(owned)).first;
                ++it;
                continue;
            }
            entry->stale = false;
            destroyList.push_back(entry->icon);
            m_lru.erase(entry->lruIt);
            it = m_entries.erase(it);
            ++m_evictions;
        }
    }

    for (HICON icon : destroyList) {
        if (icon) {
            DestroyIcon(icon);
        }
    }
}

void IconCache::InvalidatePidl(PCIDLIST_ABSOLUTE pidl) {
    InvalidateFamily(BuildIconCacheFamilyKey(pidl, {}));
}

void IconCache::InvalidatePath(const std::wstring& path) {
    InvalidateFamily(BuildIconCacheFamilyKey(nullptr, path));
}

void IconCache::LogStatsNow() {
    StatsSnapshot snapshot;
    {
        std::unique_lock lock(m_mutex);
        snapshot.hits = m_hits;
        snapshot.misses = m_misses;
        snapshot.evictions = m_evictions;
        snapshot.size = m_entries.size();
    }
    const double total = static_cast<double>(snapshot.hits + snapshot.misses);
    const double hitRate = total > 0.0 ? (static_cast<double>(snapshot.hits) / total) * 100.0 : 0.0;
    LogMessage(LogLevel::Info, L"IconCache stats: size=%zu hits=%llu misses=%llu evictions=%llu hitRate=%.1f%%",
               snapshot.size, static_cast<unsigned long long>(snapshot.hits),
               static_cast<unsigned long long>(snapshot.misses),
               static_cast<unsigned long long>(snapshot.evictions), hitRate);
}

std::wstring IconCache::BuildVariantKey(const std::wstring& familyKey, UINT iconFlags) {
    std::wstring result = familyKey;
    result.push_back(L'|');
    if ((iconFlags & SHGFI_LARGEICON) != 0) {
        result.push_back(L'L');
    } else {
        result.push_back(L'S');
    }
    return result;
}

void IconCache::Touch(Entry* entry) {
    if (!entry) {
        return;
    }
    if (entry->lruIt != m_lru.begin()) {
        m_lru.splice(m_lru.begin(), m_lru, entry->lruIt);
    }
}

std::optional<IconCache::StatsSnapshot> IconCache::LogStatsIfNeededLocked() {
    if (m_requestsSinceLog < kLogInterval) {
        return std::nullopt;
    }
    StatsSnapshot snapshot;
    snapshot.hits = m_hits;
    snapshot.misses = m_misses;
    snapshot.evictions = m_evictions;
    snapshot.size = m_entries.size();
    m_requestsSinceLog = 0;
    return snapshot;
}

void IconCache::TrimLocked(std::vector<HICON>* destroyList) {
    if (!destroyList) {
        return;
    }
    while (m_entries.size() > m_capacity && !m_lru.empty()) {
        auto listIt = std::prev(m_lru.end());
        auto mapIt = m_entries.find(*listIt);
        if (mapIt == m_entries.end()) {
            m_lru.erase(listIt);
            continue;
        }
        Entry* entry = mapIt->second.get();
        if (!entry || entry->refCount > 0) {
            break;
        }
        entry->stale = false;
        destroyList->push_back(entry->icon);
        m_entries.erase(mapIt);
        m_lru.erase(listIt);
        ++m_evictions;
    }
}

void IconCache::AddRef(Entry* entry) {
    if (!entry) {
        return;
    }
    std::unique_lock lock(m_mutex);
    ++entry->refCount;
    Touch(entry);
}

void IconCache::Release(Entry* entry) {
    if (!entry) {
        return;
    }
    std::vector<HICON> destroyList;
    {
        std::unique_lock lock(m_mutex);
        if (entry->refCount > 0) {
            --entry->refCount;
        }
        if (entry->refCount == 0) {
            if (entry->stale) {
                auto it = m_entries.find(entry->key);
                if (it != m_entries.end()) {
                    destroyList.push_back(entry->icon);
                    m_lru.erase(entry->lruIt);
                    m_entries.erase(it);
                    ++m_evictions;
                }
            } else {
                TrimLocked(&destroyList);
            }
        }
    }
    for (HICON icon : destroyList) {
        if (icon) {
            DestroyIcon(icon);
        }
    }
}

}  // namespace shelltabs

