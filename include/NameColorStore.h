#pragma once

#include <windows.h>

#include <optional>
#include <shared_mutex>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

namespace shelltabs {

// Persistent per-item and per-extension name colors for Explorer's folder view and tree.
//
// Stored as a flat JSON object in %APPDATA%\ShellTabs\namecolors.json:
//   { "c:\\some\\folder": "#FF8000", "*.cpp": "#269ED9" }
// Path keys are lower-cased full parsing names; extension keys are "*.ext" (lower-case).
class NameColorStore {
public:
    static NameColorStore& Instance();

    // Cheap check used on paint paths to skip all work when no colors are configured.
    bool IsEmpty() const;

    // Resolves the color for an item. Exact path wins; files fall back to their extension.
    std::optional<COLORREF> Lookup(const std::wstring& parsingPath, bool isFolder) const;

    std::optional<COLORREF> GetPathColor(const std::wstring& parsingPath) const;
    std::optional<COLORREF> GetExtensionColor(const std::wstring& extension) const;
    std::vector<std::pair<std::wstring, COLORREF>> GetExtensionColors() const;

    bool SetPathColors(const std::vector<std::wstring>& parsingPaths, std::optional<COLORREF> color);
    bool SetExtensionColor(const std::wstring& extension, std::optional<COLORREF> color);

    void Reload();

    // Registered message posted to every Explorer frame after a change so all processes reload.
    static UINT ChangedMessage();
    static void BroadcastChanged();

    // ".TXT", "txt" and "*.txt" all normalize to "*.txt". Returns empty for invalid input.
    static std::wstring NormalizeExtensionKey(const std::wstring& extension);
    static std::wstring ExtensionFromPath(const std::wstring& path);

private:
    NameColorStore();

    void LoadLocked();
    bool SaveLocked() const;
    std::wstring ResolveStoragePath() const;

    mutable std::shared_mutex m_mutex;
    std::unordered_map<std::wstring, COLORREF> m_colors;
};

}  // namespace shelltabs
