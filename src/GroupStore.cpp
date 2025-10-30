#include "GroupStore.h"

#include "StringUtils.h"

#include "ColorSerialization.h"

#include <algorithm>

#include "Utilities.h"

namespace shelltabs {
namespace {
constexpr wchar_t kStorageFile[] = L"groups.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kGroupToken[] = L"group";
constexpr wchar_t kTabToken[] = L"tab";
constexpr wchar_t kCommentChar = L'#';

}  // namespace

GroupStore& GroupStore::Instance() {
    static GroupStore instance;
    return instance;
}

std::vector<std::wstring> GroupStore::GroupNames() const {
    std::vector<std::wstring> names;
    names.reserve(m_groups.size());
    for (const auto& group : m_groups) {
        names.push_back(group.name);
    }
    std::sort(names.begin(), names.end(), [](const std::wstring& a, const std::wstring& b) {
        return _wcsicmp(a.c_str(), b.c_str()) < 0;
    });
    return names;
}

const SavedGroup* GroupStore::Find(const std::wstring& name) const {
    if (!EnsureLoaded()) {
        return nullptr;
    }
    for (const auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            return &group;
        }
    }
    return nullptr;
}

bool GroupStore::Load() {
    if (m_loaded) {
        return true;
    }

    const std::wstring path = ResolveStoragePath();
    if (path.empty()) {
        return false;
    }

    m_storagePath = path;
    m_groups.clear();

    std::wstring content;
    if (!ReadUtf8File(path, &content)) {
        return false;
    }

    if (content.empty()) {
        m_loaded = true;
        return true;
    }

    size_t lineStart = 0;
    SavedGroup* currentGroup = nullptr;
    int version = 1;

    while (lineStart < content.size()) {
        size_t lineEnd = content.find(L'\n', lineStart);
        std::wstring line = content.substr(lineStart, lineEnd == std::wstring::npos ? std::wstring::npos : lineEnd - lineStart);
        if (lineEnd == std::wstring::npos) {
            lineStart = content.size();
        } else {
            lineStart = lineEnd + 1;
        }

        line = Trim(line);
        if (line.empty() || line.front() == kCommentChar) {
            continue;
        }

        auto tokens = Split(line, L'|');
        if (tokens.empty()) {
            continue;
        }

        for (auto& token : tokens) {
            token = Trim(token);
        }

        if (tokens[0] == kVersionToken) {
            if (tokens.size() >= 2) {
                version = _wtoi(tokens[1].c_str());
                if (version < 1) {
                    version = 1;
                }
            }
            continue;
        }

        if (tokens[0] == kGroupToken) {
            if (tokens.size() < 3) {
                currentGroup = nullptr;
                continue;
            }
            SavedGroup group;
            group.name = tokens[1];
            group.color = ParseColor(tokens[2], RGB(0, 120, 215));
            if (version >= 2 && tokens.size() >= 4) {
                group.outlineStyle = ParseOutlineStyle(tokens[3], TabGroupOutlineStyle::kSolid);
            }
            m_groups.emplace_back(std::move(group));
            currentGroup = &m_groups.back();
            continue;
        }

        if (tokens[0] == kTabToken) {
            if (!currentGroup || tokens.size() < 2) {
                continue;
            }
            currentGroup->tabPaths.emplace_back(tokens[1]);
        }
    }

    m_loaded = true;
    return true;
}

bool GroupStore::Save() const {
    if (!EnsureLoaded()) {
        return false;
    }

    std::wstring path = m_storagePath;
    if (path.empty()) {
        path = ResolveStoragePath();
        if (path.empty()) {
            return false;
        }
        m_storagePath = path;
    }

    const size_t separator = path.find_last_of(L"\\/");
    if (separator != std::wstring::npos) {
        std::wstring directory = path.substr(0, separator);
        if (!directory.empty()) {
            CreateDirectoryW(directory.c_str(), nullptr);
        }
    }

    std::wstring content;
    content += kVersionToken;
    content += L"|2\n";
    for (const auto& group : m_groups) {
        content += kGroupToken;
        content += L"|" + group.name + L"|" + ColorToString(group.color) + L"|" +
                   OutlineStyleToString(group.outlineStyle) + L"\n";
        for (const auto& pathEntry : group.tabPaths) {
            content += kTabToken;
            content += L"|" + pathEntry + L"\n";
        }
    }

    return WriteUtf8File(path, content);
}

bool GroupStore::Upsert(SavedGroup group) {
    if (!EnsureLoaded()) {
        return false;
    }

    for (auto& existing : m_groups) {
        if (_wcsicmp(existing.name.c_str(), group.name.c_str()) == 0) {
            existing.name = group.name;
            existing.color = group.color;
            existing.tabPaths = std::move(group.tabPaths);
            existing.outlineStyle = group.outlineStyle;
            return Save();
        }
    }

    m_groups.emplace_back(std::move(group));
    return Save();
}

bool GroupStore::UpdateTabs(const std::wstring& name, const std::vector<std::wstring>& tabPaths) {
    if (!EnsureLoaded()) {
        return false;
    }

    for (auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            group.tabPaths = tabPaths;
            return Save();
        }
    }
    return false;
}

bool GroupStore::UpdateColor(const std::wstring& name, COLORREF color) {
    if (!EnsureLoaded()) {
        return false;
    }

    for (auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            group.color = color;
            return Save();
        }
    }
    return false;
}

bool GroupStore::Remove(const std::wstring& name) {
    if (!EnsureLoaded()) {
        return false;
    }
    auto it = std::remove_if(m_groups.begin(), m_groups.end(), [&](const SavedGroup& group) {
        return _wcsicmp(group.name.c_str(), name.c_str()) == 0;
    });
    if (it == m_groups.end()) {
        return false;
    }
    m_groups.erase(it, m_groups.end());
    return Save();
}

std::wstring GroupStore::ResolveStoragePath() const {
    std::wstring base = GetShellTabsDataDirectory();
    if (base.empty()) {
        return {};
    }

    if (!base.empty() && base.back() != L'\\') {
        base.push_back(L'\\');
    }
    base += kStorageFile;
    return base;
}

bool GroupStore::EnsureLoaded() const {
    if (!m_loaded) {
        const_cast<GroupStore*>(this)->Load();
    }
    return m_loaded;
}

}  // namespace shelltabs
