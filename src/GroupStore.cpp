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
    if (!EnsureLoaded()) {
        return {};
    }

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

namespace {

std::wstring FormatLoadError(std::wstring message, DWORD error) {
    if (error != ERROR_SUCCESS) {
        message += L" (error=";
        message += std::to_wstring(static_cast<unsigned long long>(error));
        message += L")";
    }
    return message;
}

}  // namespace

bool GroupStore::Load(std::wstring* errorContext) {
    if (m_loaded) {
        if (errorContext) {
            errorContext->clear();
        }
        return true;
    }

    const std::wstring path = ResolveStoragePath();
    if (path.empty()) {
        if (errorContext) {
            *errorContext = L"Group store path unavailable";
        }
        m_loaded = false;
        return false;
    }

    m_storagePath = path;
    m_groups.clear();

    std::wstring content;
    if (!ReadUtf8File(path, &content)) {
        const DWORD readError = GetLastError();
        if (errorContext) {
            std::wstring message = L"Failed to read ";
            message += path;
            *errorContext = FormatLoadError(std::move(message), readError);
        }
        m_loaded = false;
        return false;
    }

    if (content.empty()) {
        m_loaded = true;
        if (errorContext) {
            errorContext->clear();
        }
        return true;
    }

    SavedGroup* currentGroup = nullptr;
    int version = 1;

    ParseConfigLines(content, kCommentChar, L'|', [&](const std::vector<std::wstring_view>& tokens) {
        if (tokens.empty()) {
            return true;
        }

        const std::wstring_view header = tokens[0];
        if (header == kVersionToken) {
            if (tokens.size() >= 2) {
                version = ParseInt(tokens[1]);
                if (version < 1) {
                    version = 1;
                }
            }
            return true;
        }

        if (header == kGroupToken) {
            if (tokens.size() < 3) {
                currentGroup = nullptr;
                return true;
            }
            SavedGroup group;
            group.name = tokens[1];
            const std::wstring colorToken(tokens[2]);
            group.color = ParseColor(colorToken, RGB(0, 120, 215));
            if (version >= 2 && tokens.size() >= 4) {
                const std::wstring outlineToken(tokens[3]);
                group.outlineStyle =
                    ParseOutlineStyle(outlineToken, TabGroupOutlineStyle::kSolid);
            }
            m_groups.emplace_back(std::move(group));
            currentGroup = &m_groups.back();
            return true;
        }

        if (header == kTabToken) {
            if (!currentGroup || tokens.size() < 2) {
                return true;
            }
            currentGroup->tabPaths.emplace_back(tokens[1]);
        }

        return true;
    });

    m_loaded = true;
    if (errorContext) {
        errorContext->clear();
    }
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
            auto previous = std::move(existing);
            existing = std::move(group);
            if (Save()) {
                return true;
            }
            existing = std::move(previous);
            return false;
        }
    }

    m_groups.emplace_back(std::move(group));
    if (Save()) {
        return true;
    }
    m_groups.pop_back();
    return false;
}

bool GroupStore::AddTabPath(const std::wstring& name, const std::wstring& path) {
    if (path.empty() || !EnsureLoaded()) {
        return false;
    }
    const auto normalize = [](std::wstring value) {
        // URLs can contain case-sensitive paths; retain their exact spelling.
        if (value.find(L"://") == std::wstring::npos) {
            std::replace(value.begin(), value.end(), L'/', L'\\');
            while (value.size() > 1 && value.back() == L'\\') {
                value.pop_back();
            }
        }
        return value;
    };
    const auto normalizedPath = normalize(path);
    for (auto& group : m_groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) != 0) {
            continue;
        }
        for (const auto& existing : group.tabPaths) {
            const auto normalizedExisting = normalize(existing);
            const bool duplicate = path.find(L"://") == std::wstring::npos
                ? _wcsicmp(normalizedExisting.c_str(), normalizedPath.c_str()) == 0
                : normalizedExisting == normalizedPath;
            if (duplicate) {
                return true;
            }
        }
        group.tabPaths.push_back(path);
        if (Save()) {
            return true;
        }
        group.tabPaths.pop_back();
        return false;
    }
    return false;
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

void GroupStore::RecordChanges(const std::vector<std::pair<std::wstring, std::wstring>>& renamedGroups,
                               const std::vector<std::wstring>& removedGroupIds) {
    ++m_changeGeneration;
    m_lastRenamedGroups = renamedGroups;
    m_lastRemovedGroups = removedGroupIds;
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

bool GroupStore::EnsureLoaded(std::wstring* errorContext) const {
    if (!m_loaded) {
        if (!const_cast<GroupStore*>(this)->Load(errorContext)) {
            return false;
        }
    } else if (errorContext) {
        errorContext->clear();
    }
    return m_loaded;
}

}  // namespace shelltabs
