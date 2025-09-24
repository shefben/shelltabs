#pragma once

#include <windows.h>

#include <mutex>
#include <string>
#include <unordered_map>
#include <vector>

namespace shelltabs {

class TagStore {
public:
    static TagStore& Instance();

    std::vector<std::wstring> GetTagsForPath(const std::wstring& path) const;
    std::wstring GetTagListForPath(const std::wstring& path) const;
    bool TryGetColorForPath(const std::wstring& path, COLORREF* color) const;
    bool TryGetColorAndTags(const std::wstring& path, COLORREF* color, std::vector<std::wstring>* tags) const;

private:
    TagStore();

    std::wstring NormalizeKey(const std::wstring& path) const;
    void EnsureLoadedLocked() const;
    void LoadLocked() const;
    std::wstring StoragePath() const;
    static std::wstring Utf8ToWide(const std::string& utf8);
    static std::wstring Trim(const std::wstring& value);
    static std::vector<std::wstring> SplitTags(const std::wstring& tags);
    static COLORREF CalculateColor(const std::wstring& tag);
    static COLORREF BlendColors(COLORREF a, COLORREF b);

    mutable std::mutex m_mutex;
    mutable bool m_loaded = false;
    mutable std::unordered_map<std::wstring, std::vector<std::wstring>> m_tagMap;
};

}  // namespace shelltabs

