#include "OptionsStore.h"

#include "BackgroundCache.h"
#include "StringUtils.h"
#include "Utilities.h"

#include <Shlwapi.h>
#include <shellapi.h>

#include <algorithm>
#include <array>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <sstream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace shelltabs {
namespace {
constexpr int kCurrentOptionsVersion = 3;
constexpr wchar_t kStorageFile[] = L"options.db";
constexpr wchar_t kVersionToken[] = L"version";
constexpr wchar_t kReopenToken[] = L"reopen_on_crash";
constexpr wchar_t kPersistToken[] = L"persist_group_paths";
constexpr wchar_t kNewTabTemplateToken[] = L"new_tab_template";
constexpr wchar_t kNewTabCustomPathToken[] = L"new_tab_custom_path";
constexpr wchar_t kNewTabSavedGroupToken[] = L"new_tab_saved_group";
constexpr wchar_t kBreadcrumbGradientToken[] = L"breadcrumb_gradient";
constexpr wchar_t kBreadcrumbFontGradientToken[] = L"breadcrumb_font_gradient";
constexpr wchar_t kBreadcrumbGradientTransparencyToken[] = L"breadcrumb_gradient_transparency";
constexpr wchar_t kBreadcrumbFontBrightnessToken[] = L"breadcrumb_font_brightness";
constexpr wchar_t kBreadcrumbFontTransparencyToken[] = L"breadcrumb_font_transparency";  // legacy
constexpr wchar_t kBreadcrumbHighlightAlphaMultiplierToken[] =
    L"breadcrumb_highlight_alpha_multiplier";
constexpr wchar_t kBreadcrumbDropdownAlphaMultiplierToken[] =
    L"breadcrumb_dropdown_alpha_multiplier";
constexpr wchar_t kBreadcrumbGradientColorsToken[] = L"breadcrumb_gradient_colors";
constexpr wchar_t kBreadcrumbFontGradientColorsToken[] = L"breadcrumb_font_gradient_colors";
constexpr wchar_t kProgressGradientColorsToken[] = L"progress_gradient_colors";
constexpr wchar_t kGlowEnabledToken[] = L"neon_glow_enabled";
constexpr wchar_t kGlowGradientToken[] = L"neon_glow_gradient";
constexpr wchar_t kGlowColorsToken[] = L"neon_glow_colors";
constexpr wchar_t kBitmapInterceptToken[] = L"glow_bitmap_intercept";
constexpr wchar_t kFileGradientFontToken[] = L"file_gradient_font";
constexpr wchar_t kDirectUiReplacementToken[] = L"directui_replacement";
constexpr wchar_t kGlowSurfaceToken[] = L"glow_surface";
constexpr wchar_t kTabSelectedColorToken[] = L"tab_selected_color";
constexpr wchar_t kTabUnselectedColorToken[] = L"tab_unselected_color";
constexpr wchar_t kExplorerAccentColorsToken[] = L"explorer_listview_accents";
constexpr wchar_t kFolderBackgroundsEnabledToken[] = L"folder_backgrounds_enabled";
constexpr wchar_t kFolderBackgroundUniversalToken[] = L"folder_background_universal";
constexpr wchar_t kFolderBackgroundEntryToken[] = L"folder_background_entry";
constexpr wchar_t kTabDockingToken[] = L"tab_docking";
constexpr wchar_t kContextMenuCommandToken[] = L"context_menu_command";
constexpr wchar_t kContextMenuSubmenuToken[] = L"context_menu_submenu";
constexpr wchar_t kContextMenuSeparatorToken[] = L"context_menu_separator";
constexpr wchar_t kContextMenuEndToken[] = L"context_menu_end";
constexpr wchar_t kContextMenuScopeDelimiter = L',';
constexpr wchar_t kContextMenuExtensionDelimiter = L';';
constexpr int kMaxSelectionCount = 4096;
constexpr wchar_t kCommentChar = L'#';

std::wstring FormatLoadError(std::wstring message, DWORD error) {
    if (error != ERROR_SUCCESS) {
        message += L" (error=";
        message += std::to_wstring(static_cast<unsigned long long>(error));
        message += L")";
    }
    return message;
}

int ParseIntInRange(std::wstring_view token, int minimum, int maximum, int fallback) {
    if (token.empty()) {
        return fallback;
    }
    int value = ParseInt(token);
    if (value < minimum) {
        return minimum;
    }
    if (value > maximum) {
        return maximum;
    }
    return value;
}

COLORREF ParseColorValue(std::wstring_view token, COLORREF fallback) {
    if (token.empty()) {
        return fallback;
    }

    unsigned int parsed = 0;
    const std::wstring buffer(token);
    if (swscanf_s(buffer.c_str(), L"%x", &parsed) != 1) {
        return fallback;
    }

    return RGB((parsed >> 16) & 0xFF, (parsed >> 8) & 0xFF, parsed & 0xFF);
}

std::wstring ColorToHexString(COLORREF color) {
    std::wostringstream stream;
    stream << std::uppercase << std::hex;
    stream.width(6);
    stream.fill(L'0');
    const unsigned int packed = (static_cast<unsigned int>(GetRValue(color)) << 16) |
                                (static_cast<unsigned int>(GetGValue(color)) << 8) |
                                static_cast<unsigned int>(GetBValue(color));
    stream << packed;
    return stream.str();
}

constexpr COLORREF kDefaultGlowPrimaryColor = RGB(0, 120, 215);
constexpr COLORREF kDefaultGlowSecondaryColor = RGB(0, 153, 255);

struct GlowSurfaceMapping {
    const wchar_t* token;
    GlowSurfaceOptions GlowSurfacePalette::*member;
    bool supportsExplorerAccent;
};

constexpr std::array<GlowSurfaceMapping, 9> kGlowSurfaceMappings = {{{L"header", &GlowSurfacePalette::header, false},
                                                                     {L"list_view", &GlowSurfacePalette::listView, true},
                                                                     {L"direct_ui", &GlowSurfacePalette::directUi, true},
                                                                     {L"toolbar", &GlowSurfacePalette::toolbar, false},
                                                                     {L"rebar", &GlowSurfacePalette::rebar, false},
                                                                     {L"edits", &GlowSurfacePalette::edits, false},
                                                                     {L"scrollbar", &GlowSurfacePalette::scrollbars, true},
                                                                     {L"popup_menu", &GlowSurfacePalette::popupMenus, true},
                                                                     {L"tooltip", &GlowSurfacePalette::tooltips, true}}};

const GlowSurfaceMapping* FindGlowSurfaceMapping(std::wstring_view token, size_t* index) {
    if (token.empty()) {
        return nullptr;
    }
    for (size_t i = 0; i < kGlowSurfaceMappings.size(); ++i) {
        if (EqualsIgnoreCase(token, kGlowSurfaceMappings[i].token)) {
            if (index) {
                *index = i;
            }
            return &kGlowSurfaceMappings[i];
        }
    }
    return nullptr;
}

GlowSurfaceOptions* GetGlowSurfaceOptions(ShellTabsOptions* options, const GlowSurfaceMapping& mapping) {
    if (!options) {
        return nullptr;
    }
    return &(options->glowPalette.*(mapping.member));
}

const GlowSurfaceOptions* GetGlowSurfaceOptions(const ShellTabsOptions* options,
                                                const GlowSurfaceMapping& mapping) {
    if (!options) {
        return nullptr;
    }
    return &(options->glowPalette.*(mapping.member));
}

GlowSurfaceMode ParseGlowMode(std::wstring_view token, GlowSurfaceMode fallback) {
    if (token.empty()) {
        return fallback;
    }
    if (EqualsIgnoreCase(token, L"accent")) {
        return GlowSurfaceMode::kExplorerAccent;
    }
    if (EqualsIgnoreCase(token, L"solid")) {
        return GlowSurfaceMode::kSolid;
    }
    if (EqualsIgnoreCase(token, L"gradient")) {
        return GlowSurfaceMode::kGradient;
    }
    return fallback;
}

const wchar_t* GlowModeToString(GlowSurfaceMode mode) {
    switch (mode) {
        case GlowSurfaceMode::kExplorerAccent:
            return L"accent";
        case GlowSurfaceMode::kSolid:
            return L"solid";
        case GlowSurfaceMode::kGradient:
            return L"gradient";
        default:
            return L"gradient";
    }
}

bool HasDirectoryPrefix(const std::wstring& path, const std::wstring& directory) {
    if (path.size() < directory.size()) {
        return false;
    }
    if (_wcsnicmp(path.c_str(), directory.c_str(), directory.size()) != 0) {
        return false;
    }
    if (path.size() == directory.size()) {
        return true;
    }
    const wchar_t separator = path[directory.size()];
    return separator == L'\\';
}

std::wstring NormalizeCachePath(const std::wstring& path, const std::wstring& directory) {
    if (path.empty() || directory.empty()) {
        return {};
    }

    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }

    if (!HasDirectoryPrefix(normalized, directory)) {
        return {};
    }

    return normalized;
}

ContextMenuInsertionAnchor ParseContextMenuAnchor(std::wstring_view token) {
    if (token.empty()) {
        return ContextMenuInsertionAnchor::kDefault;
    }
    if (EqualsIgnoreCase(token, L"top")) {
        return ContextMenuInsertionAnchor::kTop;
    }
    if (EqualsIgnoreCase(token, L"bottom")) {
        return ContextMenuInsertionAnchor::kBottom;
    }
    if (EqualsIgnoreCase(token, L"before_shell")) {
        return ContextMenuInsertionAnchor::kBeforeShellItems;
    }
    if (EqualsIgnoreCase(token, L"after_shell")) {
        return ContextMenuInsertionAnchor::kAfterShellItems;
    }
    return ContextMenuInsertionAnchor::kDefault;
}

const wchar_t* ContextMenuAnchorToString(ContextMenuInsertionAnchor anchor) {
    switch (anchor) {
        case ContextMenuInsertionAnchor::kTop:
            return L"top";
        case ContextMenuInsertionAnchor::kBottom:
            return L"bottom";
        case ContextMenuInsertionAnchor::kBeforeShellItems:
            return L"before_shell";
        case ContextMenuInsertionAnchor::kAfterShellItems:
            return L"after_shell";
        default:
            return L"default";
    }
}

ContextMenuItemScope ParseContextMenuScope(std::wstring_view token) {
    ContextMenuItemScope scope;
    scope.includeAllFiles = true;
    scope.includeAllFolders = true;

    if (token.empty()) {
        return scope;
    }

    scope.includeAllFiles = false;
    scope.includeAllFolders = false;
    bool anyToken = false;
    for (std::wstring_view part : Split(token, kContextMenuScopeDelimiter)) {
        const std::wstring trimmed = Trim(part);
        if (trimmed.empty()) {
            continue;
        }
        anyToken = true;
        if (EqualsIgnoreCase(trimmed, L"all")) {
            scope.includeAllFiles = true;
            scope.includeAllFolders = true;
            break;
        }
        if (EqualsIgnoreCase(trimmed, L"files") || EqualsIgnoreCase(trimmed, L"file")) {
            scope.includeAllFiles = true;
            continue;
        }
        if (EqualsIgnoreCase(trimmed, L"folders") || EqualsIgnoreCase(trimmed, L"folder")) {
            scope.includeAllFolders = true;
            continue;
        }
        if (EqualsIgnoreCase(trimmed, L"none")) {
            scope.includeAllFiles = false;
            scope.includeAllFolders = false;
            continue;
        }
        if (EqualsIgnoreCase(trimmed, L"extensions") || EqualsIgnoreCase(trimmed, L"ext")) {
            continue;
        }
    }

    if (!anyToken) {
        scope.includeAllFiles = true;
        scope.includeAllFolders = true;
    }

    return scope;
}

std::wstring BuildContextMenuScopeString(const ContextMenuItemScope& scope) {
    std::vector<std::wstring> tokens;
    if (scope.includeAllFiles && scope.includeAllFolders && scope.extensions.empty()) {
        tokens.emplace_back(L"all");
    } else {
        if (scope.includeAllFiles) {
            tokens.emplace_back(L"files");
        }
        if (scope.includeAllFolders) {
            tokens.emplace_back(L"folders");
        }
        if (!scope.extensions.empty()) {
            tokens.emplace_back(L"extensions");
        }
        if (tokens.empty()) {
            tokens.emplace_back(L"none");
        }
    }

    std::wstring result;
    for (size_t i = 0; i < tokens.size(); ++i) {
        if (i > 0) {
            result.push_back(kContextMenuScopeDelimiter);
        }
        result += tokens[i];
    }
    return result;
}

std::vector<std::wstring> ParseContextMenuExtensions(std::wstring_view token) {
    std::vector<std::wstring> extensions;
    if (token.empty()) {
        return extensions;
    }

    for (std::wstring_view part : Split(token, kContextMenuExtensionDelimiter)) {
        std::wstring trimmed = Trim(part);
        if (trimmed.empty()) {
            continue;
        }
        extensions.emplace_back(std::move(trimmed));
    }

    return NormalizeContextMenuPatterns(extensions);
}

std::wstring BuildContextMenuExtensionsToken(const std::vector<std::wstring>& extensions) {
    const std::vector<std::wstring> normalized = NormalizeContextMenuPatterns(extensions);
    std::wstring result;
    for (size_t i = 0; i < normalized.size(); ++i) {
        if (i > 0) {
            result.push_back(kContextMenuExtensionDelimiter);
        }
        result += normalized[i];
    }
    return result;
}

int ParseSelectionCount(std::wstring_view token) {
    if (token.empty()) {
        return 0;
    }
    int value = ParseInt(token);
    if (value <= 0) {
        return 0;
    }
    if (value > kMaxSelectionCount) {
        return kMaxSelectionCount;
    }
    return value;
}

void NormalizeContextMenuItem(ContextMenuItem* item) {
    if (!item) {
        return;
    }

    if (item->type == ContextMenuItemType::kSeparator) {
        item->label.clear();
        item->executable.clear();
        item->arguments.clear();
    } else {
        item->label = Trim(item->label);
        item->executable = Trim(item->executable);
        item->arguments = Trim(item->arguments);
    }

    item->iconSource = NormalizeContextMenuIconSource(item->iconSource);
    item->workingDirectory = Trim(item->workingDirectory);
    item->description = Trim(item->description);
    item->id = Trim(item->id);

    // Normalize visibility patterns
    item->visibility.filePatterns = NormalizeContextMenuPatterns(item->visibility.filePatterns);
    item->visibility.excludePatterns = NormalizeContextMenuPatterns(item->visibility.excludePatterns);

    if (item->visibility.minimumSelection < 0) {
        item->visibility.minimumSelection = 0;
    }
    if (item->visibility.maximumSelection < 0) {
        item->visibility.maximumSelection = 0;
    }
    if (item->visibility.maximumSelection > 0 &&
        item->visibility.maximumSelection < item->visibility.minimumSelection) {
        item->visibility.maximumSelection = item->visibility.minimumSelection;
    }
}

void NormalizeContextMenuItems(std::vector<ContextMenuItem>* items) {
    if (!items) {
        return;
    }
    for (auto& item : *items) {
        NormalizeContextMenuItem(&item);
        NormalizeContextMenuItems(&item.children);
    }
}

// Helper to join patterns with delimiter
std::wstring JoinPatterns(const std::vector<std::wstring>& patterns, wchar_t delimiter) {
    std::wstring result;
    for (size_t i = 0; i < patterns.size(); ++i) {
        if (i > 0) result += delimiter;
        result += patterns[i];
    }
    return result;
}

void AppendContextMenuItems(std::wstring& content, const std::vector<ContextMenuItem>& items) {
    for (const auto& item : items) {
        switch (item.type) {
            case ContextMenuItemType::kCommand: {
                content += kContextMenuCommandToken;
                content += L"|";
                content += item.label;
                content += L"|";
                content += item.executable;
                content += L"|";
                content += item.arguments;
                content += L"|";
                content += NormalizeContextMenuIconSource(item.iconSource);
                content += L"|";
                content += item.workingDirectory;
                content += L"|";
                content += std::to_wstring(static_cast<int>(item.windowState));
                content += L"|";
                content += item.runAsAdmin ? L"1" : L"0";
                content += L"|";
                content += item.waitForCompletion ? L"1" : L"0";
                content += L"|";
                content += ContextMenuAnchorToString(item.anchor);
                content += L"|";
                content += item.enabled ? L"1" : L"0";
                content += L"|";
                content += std::to_wstring(std::max(item.visibility.minimumSelection, 0));
                content += L"|";
                content += std::to_wstring(item.visibility.maximumSelection > 0 ? item.visibility.maximumSelection : 0);
                content += L"|";
                content += item.visibility.showForFiles ? L"1" : L"0";
                content += L"|";
                content += item.visibility.showForFolders ? L"1" : L"0";
                content += L"|";
                content += item.visibility.showForMultiple ? L"1" : L"0";
                content += L"|";
                content += JoinPatterns(item.visibility.filePatterns, L';');
                content += L"|";
                content += JoinPatterns(item.visibility.excludePatterns, L';');
                content += L"|";
                content += item.description;
                content += L"|";
                content += item.id;
                content += L"\n";
                break;
            }
            case ContextMenuItemType::kSubmenu: {
                content += kContextMenuSubmenuToken;
                content += L"|";
                content += item.label;
                content += L"|";
                content += NormalizeContextMenuIconSource(item.iconSource);
                content += L"|";
                content += ContextMenuAnchorToString(item.anchor);
                content += L"|";
                content += item.enabled ? L"1" : L"0";
                content += L"|";
                content += std::to_wstring(std::max(item.visibility.minimumSelection, 0));
                content += L"|";
                content += std::to_wstring(item.visibility.maximumSelection > 0 ? item.visibility.maximumSelection : 0);
                content += L"|";
                content += item.visibility.showForFiles ? L"1" : L"0";
                content += L"|";
                content += item.visibility.showForFolders ? L"1" : L"0";
                content += L"|";
                content += item.visibility.showForMultiple ? L"1" : L"0";
                content += L"|";
                content += JoinPatterns(item.visibility.filePatterns, L';');
                content += L"|";
                content += JoinPatterns(item.visibility.excludePatterns, L';');
                content += L"|";
                content += item.description;
                content += L"|";
                content += item.id;
                content += L"\n";
                AppendContextMenuItems(content, item.children);
                content += kContextMenuEndToken;
                content += L"\n";
                break;
            }
            case ContextMenuItemType::kSeparator: {
                content += kContextMenuSeparatorToken;
                content += L"|";
                content += ContextMenuAnchorToString(item.anchor);
                content += L"|";
                content += item.enabled ? L"1" : L"0";
                content += L"\n";
                break;
            }
        }
    }
}

}  // namespace

// ==================== Context Menu System Implementation ====================

// Backward compatibility: split old commandTemplate into executable + arguments
void MigrateCommandTemplate(ContextMenuItem* item) {
    if (!item || item->type != ContextMenuItemType::kCommand) {
        return;
    }

    // If we already have an executable, no need to migrate
    if (!item->executable.empty()) {
        return;
    }

    // This would be populated by old format - we'll leave the interface for now
    // The old code can call SplitCommandTemplate manually if needed
}

// Helper to build command template from executable + arguments (for old code)
std::wstring BuildCommandTemplateFromParts(const std::wstring& executable, const std::wstring& arguments) {
    std::wstring result = executable;
    if (!arguments.empty()) {
        if (!result.empty()) {
            result += L" ";
        }
        result += arguments;
    }
    return result;
}

// ==================== Context Menu System Implementation ====================

// Pattern matching helper - supports wildcards (* and ?)
bool MatchesContextMenuPattern(const std::wstring& filename, const std::wstring& pattern) {
    if (pattern.empty()) {
        return true;
    }
    if (filename.empty()) {
        return false;
    }

    // Convert to lowercase for case-insensitive matching
    std::wstring lowerFilename = filename;
    std::wstring lowerPattern = pattern;
    std::transform(lowerFilename.begin(), lowerFilename.end(), lowerFilename.begin(),
                  [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });
    std::transform(lowerPattern.begin(), lowerPattern.end(), lowerPattern.begin(),
                  [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });

    // Simple wildcard matching implementation
    size_t fnPos = 0;
    size_t patPos = 0;
    size_t starPos = std::wstring::npos;
    size_t matchPos = 0;

    while (fnPos < lowerFilename.size()) {
        if (patPos < lowerPattern.size() && (lowerPattern[patPos] == L'?' ||
            lowerPattern[patPos] == lowerFilename[fnPos])) {
            ++fnPos;
            ++patPos;
        } else if (patPos < lowerPattern.size() && lowerPattern[patPos] == L'*') {
            starPos = patPos++;
            matchPos = fnPos;
        } else if (starPos != std::wstring::npos) {
            patPos = starPos + 1;
            fnPos = ++matchPos;
        } else {
            return false;
        }
    }

    while (patPos < lowerPattern.size() && lowerPattern[patPos] == L'*') {
        ++patPos;
    }

    return patPos == lowerPattern.size();
}

// Normalize file patterns (convert to lowercase, remove duplicates)
std::vector<std::wstring> NormalizeContextMenuPatterns(const std::vector<std::wstring>& patterns) {
    std::vector<std::wstring> normalized;
    normalized.reserve(patterns.size());

    for (const auto& pattern : patterns) {
        std::wstring trimmed = Trim(pattern);
        if (trimmed.empty()) {
            continue;
        }

        // Convert to lowercase
        std::transform(trimmed.begin(), trimmed.end(), trimmed.begin(),
                      [](wchar_t c) { return static_cast<wchar_t>(std::towlower(c)); });

        // Add if not already present
        if (std::find(normalized.begin(), normalized.end(), trimmed) == normalized.end()) {
            normalized.emplace_back(std::move(trimmed));
        }
    }

    std::sort(normalized.begin(), normalized.end());
    return normalized;
}

// Legacy extension normalization (kept for compatibility)
std::wstring NormalizeContextMenuExtensions(const std::vector<std::wstring>& extensions) {
    std::wstring result;
    auto normalized = NormalizeContextMenuPatterns(extensions);
    for (size_t i = 0; i < normalized.size(); ++i) {
        if (i > 0) result += L";";
        result += normalized[i];
    }
    return result;
}

// Check if a context menu item matches the current selection
bool ContextMenuItemMatchesSelection(const ContextMenuItem& item, int selectionCount,
                                    const std::vector<std::wstring>& selectedPaths,
                                    bool hasFiles, bool hasFolders) {
    const auto& rules = item.visibility;

    // Check selection count constraints
    if (rules.minimumSelection > 0 && selectionCount < rules.minimumSelection) {
        return false;
    }
    if (rules.maximumSelection > 0 && selectionCount > rules.maximumSelection) {
        return false;
    }

    // Check multiple selection constraint
    if (selectionCount > 1 && !rules.showForMultiple) {
        return false;
    }

    // Check file/folder type constraints
    if (hasFiles && !rules.showForFiles) {
        return false;
    }
    if (hasFolders && !rules.showForFolders) {
        return false;
    }

    // Check file patterns if specified
    if (!rules.filePatterns.empty() || !rules.excludePatterns.empty()) {
        bool matchesPattern = false;

        for (const auto& path : selectedPaths) {
            size_t lastSlash = path.find_last_of(L"\\/");
            std::wstring filename = (lastSlash != std::wstring::npos) ?
                                   path.substr(lastSlash + 1) : path;

            // Check exclude patterns first
            bool excluded = false;
            for (const auto& exclude : rules.excludePatterns) {
                if (MatchesContextMenuPattern(filename, exclude)) {
                    excluded = true;
                    break;
                }
            }

            if (excluded) {
                continue;
            }

            // Check include patterns
            if (rules.filePatterns.empty()) {
                matchesPattern = true; // No patterns = match all (unless excluded)
            } else {
                for (const auto& pattern : rules.filePatterns) {
                    if (MatchesContextMenuPattern(filename, pattern)) {
                        matchesPattern = true;
                        break;
                    }
                }
            }

            if (matchesPattern) {
                break;
            }
        }

        if (!matchesPattern) {
            return false;
        }
    }

    return true;
}

// Expand placeholders in strings
std::wstring ExpandContextMenuPlaceholders(const std::wstring& text,
                                          const std::vector<std::wstring>& selectedPaths) {
    if (text.empty() || selectedPaths.empty()) {
        return text;
    }

    std::wstring result = text;

    // %1 = first selected item
    if (result.find(L"%1") != std::wstring::npos && !selectedPaths.empty()) {
        size_t pos = 0;
        while ((pos = result.find(L"%1", pos)) != std::wstring::npos) {
            result.replace(pos, 2, L"\"" + selectedPaths[0] + L"\"");
            pos += selectedPaths[0].length() + 2;
        }
    }

    // %V = all selected items (space-separated, quoted)
    if (result.find(L"%V") != std::wstring::npos) {
        std::wstring allPaths;
        for (size_t i = 0; i < selectedPaths.size(); ++i) {
            if (i > 0) allPaths += L" ";
            allPaths += L"\"" + selectedPaths[i] + L"\"";
        }
        size_t pos = 0;
        while ((pos = result.find(L"%V", pos)) != std::wstring::npos) {
            result.replace(pos, 2, allPaths);
            pos += allPaths.length();
        }
    }

    // %N = number of selected items
    if (result.find(L"%N") != std::wstring::npos) {
        std::wstring count = std::to_wstring(selectedPaths.size());
        size_t pos = 0;
        while ((pos = result.find(L"%N", pos)) != std::wstring::npos) {
            result.replace(pos, 2, count);
            pos += count.length();
        }
    }

    // %P = parent directory of first selected item
    if (result.find(L"%P") != std::wstring::npos && !selectedPaths.empty()) {
        std::wstring parentDir = selectedPaths[0];
        size_t lastSlash = parentDir.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos) {
            parentDir = parentDir.substr(0, lastSlash);
        }
        size_t pos = 0;
        while ((pos = result.find(L"%P", pos)) != std::wstring::npos) {
            result.replace(pos, 2, L"\"" + parentDir + L"\"");
            pos += parentDir.length() + 2;
        }
    }

    return result;
}

// Validate a context menu item configuration
void ValidateContextMenuItem(const ContextMenuItem& item, std::vector<std::wstring>& errors) {
    if (item.type == ContextMenuItemType::kSeparator) {
        return; // Separators don't need validation
    }

    if (item.label.empty()) {
        errors.push_back(L"Menu item label cannot be empty");
    }

    if (item.type == ContextMenuItemType::kCommand) {
        if (item.executable.empty()) {
            errors.push_back(L"Command executable path cannot be empty");
        }

        // Check for invalid selection constraints
        if (item.visibility.minimumSelection < 0) {
            errors.push_back(L"Minimum selection cannot be negative");
        }
        if (item.visibility.maximumSelection < 0) {
            errors.push_back(L"Maximum selection cannot be negative");
        }
        if (item.visibility.maximumSelection > 0 &&
            item.visibility.maximumSelection < item.visibility.minimumSelection) {
            errors.push_back(L"Maximum selection cannot be less than minimum selection");
        }
    }

    if (item.type == ContextMenuItemType::kSubmenu) {
        if (item.children.empty()) {
            errors.push_back(L"Submenu must contain at least one child item");
        } else {
            // Recursively validate children
            for (const auto& child : item.children) {
                ValidateContextMenuItem(child, errors);
            }
        }
    }
}

// Create a context menu item from a template
ContextMenuItem CreateContextMenuTemplate(const std::wstring& templateType) {
    ContextMenuItem item;
    item.enabled = true;
    item.anchor = ContextMenuInsertionAnchor::kDefault;

    if (templateType == L"open_with") {
        item.type = ContextMenuItemType::kCommand;
        item.label = L"Open with...";
        item.executable = L"rundll32.exe";
        item.arguments = L"shell32.dll,OpenAs_RunDLL %1";
        item.iconSource = L"shell32.dll,3";
        item.visibility.maximumSelection = 1;
        item.visibility.showForFiles = true;
        item.visibility.showForFolders = false;
    } else if (templateType == L"cmd_here") {
        item.type = ContextMenuItemType::kCommand;
        item.label = L"Command Prompt Here";
        item.executable = L"cmd.exe";
        item.arguments = L"/k cd /d %P";
        item.iconSource = L"cmd.exe,0";
        item.visibility.showForFiles = true;
        item.visibility.showForFolders = true;
        item.workingDirectory = L"%P";
    } else if (templateType == L"powershell_here") {
        item.type = ContextMenuItemType::kCommand;
        item.label = L"PowerShell Here";
        item.executable = L"powershell.exe";
        item.arguments = L"-NoExit -Command Set-Location -Path %P";
        item.iconSource = L"powershell.exe,0";
        item.visibility.showForFiles = true;
        item.visibility.showForFolders = true;
        item.workingDirectory = L"%P";
    } else if (templateType == L"copy_path") {
        item.type = ContextMenuItemType::kCommand;
        item.label = L"Copy Full Path";
        item.executable = L"cmd.exe";
        item.arguments = L"/c echo %1 | clip";
        item.iconSource = L"shell32.dll,134";
        item.windowState = ContextMenuWindowState::kHidden;
    } else if (templateType == L"properties") {
        item.type = ContextMenuItemType::kCommand;
        item.label = L"Properties";
        item.executable = L"rundll32.exe";
        item.arguments = L"shell32.dll,Control_RunDLL shell32.dll,,Properties %1";
        item.iconSource = L"shell32.dll,21";
        item.visibility.maximumSelection = 1;
    } else {
        // Default empty command
        item.type = ContextMenuItemType::kCommand;
        item.label = L"New Command";
        item.executable = L"";
        item.arguments = L"";
    }

    return item;
}

std::wstring NormalizeContextMenuIconSource(const std::wstring& iconSource) {
    std::wstring trimmed = Trim(iconSource);
    if (trimmed.empty()) {
        return {};
    }

    std::wstring location = trimmed;
    if (location.size() >= 2 && location.front() == L'"' && location.back() == L'"') {
        location = location.substr(1, location.size() - 2);
    }

    std::vector<wchar_t> buffer(location.begin(), location.end());
    buffer.push_back(L'\0');
    const int iconIndex = PathParseIconLocationW(buffer.data());
    std::wstring path(buffer.data());
    path = Trim(path);

    if (!path.empty() && IsLikelyFileSystemPath(path)) {
        std::wstring normalizedPath = NormalizeFileSystemPath(path);
        if (!normalizedPath.empty()) {
            path = std::move(normalizedPath);
        }
    }

    if (path.empty()) {
        return {};
    }

    std::wstring result = path;
    if (location.find(L',') != std::wstring::npos || iconIndex != 0) {
        result += L",";
        result += std::to_wstring(static_cast<long long>(iconIndex));
    }
    return result;
}

IconCache::Reference ResolveContextMenuIcon(const std::wstring& iconSource, UINT iconFlags) {
    const std::wstring normalizedSource = NormalizeContextMenuIconSource(iconSource);
    if (normalizedSource.empty()) {
        return {};
    }

    std::vector<wchar_t> buffer(normalizedSource.begin(), normalizedSource.end());
    buffer.push_back(L'\0');
    const int iconIndex = PathParseIconLocationW(buffer.data());
    std::wstring path(buffer.data());
    path = Trim(path);
    if (path.empty()) {
        return {};
    }

    const bool wantLarge = (iconFlags & SHGFI_LARGEICON) != 0;
    const bool wantSmall = (iconFlags & SHGFI_SMALLICON) != 0;
    const bool requestLarge = wantLarge || !wantSmall;
    const bool requestSmall = wantSmall || !wantLarge;

    return IconCache::Instance().Acquire(normalizedSource, iconFlags, [path, iconIndex, requestLarge, requestSmall,
                                                                      wantLarge, wantSmall]() -> HICON {
        HICON largeIcon = nullptr;
        HICON smallIcon = nullptr;
        const UINT extracted = ExtractIconExW(path.c_str(), iconIndex, requestLarge ? &largeIcon : nullptr,
                                              requestSmall ? &smallIcon : nullptr, 1);
        if (extracted == 0) {
            if (largeIcon) {
                DestroyIcon(largeIcon);
            }
            if (smallIcon) {
                DestroyIcon(smallIcon);
            }
            return nullptr;
        }

        HICON result = nullptr;
        if (wantSmall && smallIcon) {
            result = smallIcon;
            if (largeIcon) {
                DestroyIcon(largeIcon);
            }
        } else if (wantLarge && largeIcon) {
            result = largeIcon;
            if (smallIcon) {
                DestroyIcon(smallIcon);
            }
        } else if (largeIcon) {
            result = largeIcon;
            if (smallIcon) {
                DestroyIcon(smallIcon);
            }
        } else if (smallIcon) {
            result = smallIcon;
        }

        if (!result) {
            if (largeIcon) {
                DestroyIcon(largeIcon);
            }
            if (smallIcon) {
                DestroyIcon(smallIcon);
            }
        } else if (result != smallIcon && smallIcon) {
            DestroyIcon(smallIcon);
        }

        return result;
    });
}

void UpdateGlowPaletteFromLegacySettings(ShellTabsOptions& options) {
    const bool gradient = options.useNeonGlowGradient;
    const COLORREF primary = options.neonGlowPrimaryColor;
    const COLORREF secondary = options.useNeonGlowGradient ? options.neonGlowSecondaryColor : primary;

    const auto applyLegacy = [&](GlowSurfaceOptions& surface, bool allowAccent) {
        if (allowAccent && options.useExplorerAccentColors) {
            surface.mode = GlowSurfaceMode::kExplorerAccent;
        } else if (gradient) {
            surface.mode = GlowSurfaceMode::kGradient;
        } else {
            surface.mode = GlowSurfaceMode::kSolid;
        }
        surface.solidColor = primary;
        surface.gradientStartColor = primary;
        surface.gradientEndColor = gradient ? secondary : primary;
    };

    applyLegacy(options.glowPalette.header, false);
    applyLegacy(options.glowPalette.toolbar, false);
    applyLegacy(options.glowPalette.rebar, false);
    applyLegacy(options.glowPalette.edits, false);
    applyLegacy(options.glowPalette.listView, true);
    applyLegacy(options.glowPalette.directUi, true);
    applyLegacy(options.glowPalette.scrollbars, true);
    applyLegacy(options.glowPalette.popupMenus, true);
    applyLegacy(options.glowPalette.tooltips, true);
}

void UpdateLegacyGlowSettingsFromPalette(ShellTabsOptions& options) {
    const GlowSurfaceOptions& header = options.glowPalette.header;

    options.useExplorerAccentColors =
        options.glowPalette.listView.mode == GlowSurfaceMode::kExplorerAccent;

    options.useNeonGlowGradient = (header.mode == GlowSurfaceMode::kGradient);
    options.neonGlowPrimaryColor = header.solidColor;
    options.neonGlowSecondaryColor = options.useNeonGlowGradient ? header.gradientEndColor : header.solidColor;

    options.useCustomNeonGlowColors =
        header.solidColor != kDefaultGlowPrimaryColor ||
        header.gradientStartColor != kDefaultGlowPrimaryColor ||
        header.gradientEndColor != kDefaultGlowSecondaryColor;
}

OptionsStore& OptionsStore::Instance() {
    static OptionsStore store;
    return store;
}

bool OptionsStore::EnsureLoaded(std::wstring* errorContext) const {
    if (m_loaded) {
        if (errorContext) {
            errorContext->clear();
        }
        return true;
    }
    return const_cast<OptionsStore*>(this)->Load(errorContext);
}

std::wstring OptionsStore::ResolveStoragePath() const {
    if (!m_storagePath.empty()) {
        return m_storagePath;
    }

    std::wstring directory = GetShellTabsDataDirectory();
    if (directory.empty()) {
        return {};
    }

    if (!directory.empty() && directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += kStorageFile;
    return directory;
}

bool OptionsStore::Load(std::wstring* errorContext) {
    m_options = {};

    m_storagePath = ResolveStoragePath();
    if (m_storagePath.empty()) {
        m_loaded = false;
        if (errorContext) {
            *errorContext = L"Options store path unavailable";
        }
        return false;
    }

    std::wstring content;
    bool fileExists = false;
    if (!ReadUtf8File(m_storagePath, &content, &fileExists)) {
        m_loaded = false;
        const DWORD readError = GetLastError();
        if (errorContext) {
            std::wstring message = L"Failed to read ";
            message += m_storagePath;
            *errorContext = FormatLoadError(std::move(message), readError);
        }
        return false;
    }

    const std::wstring storageDirectory = GetShellTabsDataDirectory();

    if (!fileExists || content.empty()) {
        m_loaded = true;
        if (errorContext) {
            errorContext->clear();
        }
        return true;
    }

    int version = 1;
    std::array<bool, kGlowSurfaceMappings.size()> glowSurfaceSpecified{};
    glowSurfaceSpecified.fill(false);
    bool anyGlowSurfaceToken = false;
    std::vector<std::vector<ContextMenuItem>*> contextMenuStack;
    contextMenuStack.push_back(&m_options.contextMenuItems);
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

        if (header == kReopenToken) {
            if (tokens.size() >= 2) {
                m_options.reopenOnCrash = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kPersistToken) {
            if (tokens.size() >= 2) {
                m_options.persistGroupPaths = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kNewTabTemplateToken) {
            if (tokens.size() >= 2) {
                m_options.newTabTemplate = ParseNewTabTemplate(tokens[1]);
            }
            return true;
        }

        if (header == kNewTabCustomPathToken) {
            if (tokens.size() >= 2) {
                m_options.newTabCustomPath = Trim(tokens[1]);
            }
            return true;
        }

        if (header == kNewTabSavedGroupToken) {
            if (tokens.size() >= 2) {
                m_options.newTabSavedGroup = Trim(tokens[1]);
            }
            return true;
        }

        if (header == kBreadcrumbGradientToken) {
            if (tokens.size() >= 2) {
                m_options.enableBreadcrumbGradient = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kBreadcrumbFontGradientToken) {
            if (tokens.size() >= 2) {
                m_options.enableBreadcrumbFontGradient = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kBreadcrumbGradientTransparencyToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbGradientTransparency =
                    ParseIntInRange(tokens[1], 0, 100, m_options.breadcrumbGradientTransparency);
            }
            return true;
        }

        if (header == kBreadcrumbFontBrightnessToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbFontBrightness =
                    ParseIntInRange(tokens[1], 0, 100, m_options.breadcrumbFontBrightness);
            }
            return true;
        }

        if (header == kBreadcrumbHighlightAlphaMultiplierToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbHighlightAlphaMultiplier =
                    ParseIntInRange(tokens[1], 0, 200, m_options.breadcrumbHighlightAlphaMultiplier);
            }
            return true;
        }

        if (header == kBreadcrumbDropdownAlphaMultiplierToken) {
            if (tokens.size() >= 2) {
                m_options.breadcrumbDropdownAlphaMultiplier =
                    ParseIntInRange(tokens[1], 0, 200, m_options.breadcrumbDropdownAlphaMultiplier);
            }
            return true;
        }

        if (header == kBreadcrumbFontTransparencyToken) {
            if (tokens.size() >= 2) {
                const int defaultBrightness = m_options.breadcrumbFontBrightness;
                const int legacyTransparency =
                    ParseIntInRange(tokens[1], 0, 100, 100 - defaultBrightness);
                const int legacyOpacity = 100 - legacyTransparency;
                m_options.breadcrumbFontBrightness =
                    std::clamp(legacyOpacity * defaultBrightness / 100, 0, 100);
            }
            return true;
        }

        if (header == kBreadcrumbGradientColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomBreadcrumbGradientColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.breadcrumbGradientStartColor =
                    ParseColorValue(tokens[2], m_options.breadcrumbGradientStartColor);
            }
            if (tokens.size() >= 4) {
                m_options.breadcrumbGradientEndColor =
                    ParseColorValue(tokens[3], m_options.breadcrumbGradientEndColor);
            }
            return true;
        }

        if (header == kBreadcrumbFontGradientColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomBreadcrumbFontColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.breadcrumbFontGradientStartColor =
                    ParseColorValue(tokens[2], m_options.breadcrumbFontGradientStartColor);
            }
            if (tokens.size() >= 4) {
                m_options.breadcrumbFontGradientEndColor =
                    ParseColorValue(tokens[3], m_options.breadcrumbFontGradientEndColor);
            }
            return true;
        }

        if (header == kProgressGradientColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomProgressBarGradientColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.progressBarGradientStartColor =
                    ParseColorValue(tokens[2], m_options.progressBarGradientStartColor);
            }
            if (tokens.size() >= 4) {
                m_options.progressBarGradientEndColor =
                    ParseColorValue(tokens[3], m_options.progressBarGradientEndColor);
            }
            return true;
        }

        if (header == kGlowEnabledToken) {
            if (tokens.size() >= 2) {
                m_options.enableNeonGlow = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kGlowGradientToken) {
            if (tokens.size() >= 2) {
                m_options.useNeonGlowGradient = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kGlowColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomNeonGlowColors = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.neonGlowPrimaryColor =
                    ParseColorValue(tokens[2], m_options.neonGlowPrimaryColor);
            }
            if (tokens.size() >= 4) {
                m_options.neonGlowSecondaryColor =
                    ParseColorValue(tokens[3], m_options.neonGlowSecondaryColor);
            }
            return true;
        }

        if (header == kBitmapInterceptToken) {
            if (tokens.size() >= 2) {
                m_options.enableBitmapIntercept = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kFileGradientFontToken) {
            if (tokens.size() >= 2) {
                m_options.enableFileGradientFont = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kDirectUiReplacementToken) {
            if (tokens.size() >= 2) {
                m_options.enableDirectUiReplacement = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kGlowSurfaceToken) {
            if (tokens.size() >= 2) {
                size_t surfaceIndex = 0;
                const GlowSurfaceMapping* mapping = FindGlowSurfaceMapping(tokens[1], &surfaceIndex);
                if (mapping) {
                    if (GlowSurfaceOptions* surface = GetGlowSurfaceOptions(&m_options, *mapping)) {
                        anyGlowSurfaceToken = true;
                        glowSurfaceSpecified[surfaceIndex] = true;
                        if (tokens.size() >= 3) {
                            surface->mode = ParseGlowMode(tokens[2], surface->mode);
                        }
                        if (tokens.size() >= 4) {
                            surface->solidColor = ParseColorValue(tokens[3], surface->solidColor);
                        }
                        if (tokens.size() >= 5) {
                            surface->gradientStartColor =
                                ParseColorValue(tokens[4], surface->gradientStartColor);
                        }
                        if (tokens.size() >= 6) {
                            surface->gradientEndColor =
                                ParseColorValue(tokens[5], surface->gradientEndColor);
                        }
                        if (tokens.size() >= 7) {
                            surface->enabled = ParseBool(tokens[6]);
                        }
                    }
                }
            }
            return true;
        }

        if (header == kTabSelectedColorToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomTabSelectedColor = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.customTabSelectedColor =
                    ParseColorValue(tokens[2], m_options.customTabSelectedColor);
            }
            return true;
        }

        if (header == kTabUnselectedColorToken) {
            if (tokens.size() >= 2) {
                m_options.useCustomTabUnselectedColor = ParseBool(tokens[1]);
            }
            if (tokens.size() >= 3) {
                m_options.customTabUnselectedColor =
                    ParseColorValue(tokens[2], m_options.customTabUnselectedColor);
            }
            return true;
        }

        if (header == kExplorerAccentColorsToken) {
            if (tokens.size() >= 2) {
                m_options.useExplorerAccentColors = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kFolderBackgroundsEnabledToken) {
            if (tokens.size() >= 2) {
                m_options.enableFolderBackgrounds = ParseBool(tokens[1]);
            }
            return true;
        }

        if (header == kFolderBackgroundUniversalToken) {
            if (tokens.size() >= 2) {
                const std::wstring cachePath =
                    NormalizeCachePath(std::wstring(tokens[1]), storageDirectory);
                if (!cachePath.empty()) {
                    m_options.universalFolderBackgroundImage.cachedImagePath = cachePath;
                    if (tokens.size() >= 3) {
                        m_options.universalFolderBackgroundImage.displayName = tokens[2];
                    }
                }
            }
            return true;
        }

        if (header == kFolderBackgroundEntryToken) {
            if (tokens.size() >= 3) {
                FolderBackgroundEntry entry;
                entry.folderPath = NormalizeFileSystemPath(std::wstring(tokens[1]));
                if (!entry.folderPath.empty()) {
                    entry.image.cachedImagePath =
                        NormalizeCachePath(std::wstring(tokens[2]), storageDirectory);
                    if (!entry.image.cachedImagePath.empty()) {
                        if (tokens.size() >= 4) {
                            entry.image.displayName = tokens[3];
                        }
                        m_options.folderBackgroundEntries.emplace_back(std::move(entry));
                    }
                }
            }
            return true;
        }

        if (header == kContextMenuCommandToken) {
            if (!contextMenuStack.empty()) {
                ContextMenuItem item;
                item.type = ContextMenuItemType::kCommand;
                size_t idx = 1;

                if (tokens.size() > idx) item.label = Trim(tokens[idx++]);
                if (tokens.size() > idx) item.executable = Trim(tokens[idx++]);
                if (tokens.size() > idx) item.arguments = Trim(tokens[idx++]);
                if (tokens.size() > idx) item.iconSource = std::wstring(tokens[idx++]);
                if (tokens.size() > idx) item.workingDirectory = Trim(tokens[idx++]);
                if (tokens.size() > idx) item.windowState = static_cast<ContextMenuWindowState>(ParseInt(tokens[idx++]));
                if (tokens.size() > idx) item.runAsAdmin = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) item.waitForCompletion = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) item.anchor = ParseContextMenuAnchor(tokens[idx++]);
                if (tokens.size() > idx) item.enabled = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.minimumSelection = ParseSelectionCount(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.maximumSelection = ParseSelectionCount(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.showForFiles = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.showForFolders = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.showForMultiple = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.filePatterns = ParseContextMenuExtensions(tokens[idx++]);
                if (tokens.size() > idx) item.visibility.excludePatterns = ParseContextMenuExtensions(tokens[idx++]);
                if (tokens.size() > idx) item.description = Trim(tokens[idx++]);
                if (tokens.size() > idx) item.id = Trim(tokens[idx++]);

                contextMenuStack.back()->push_back(std::move(item));
            }
            return true;
        }

        if (header == kContextMenuSubmenuToken) {
            if (!contextMenuStack.empty()) {
                ContextMenuItem submenu;
                submenu.type = ContextMenuItemType::kSubmenu;
                size_t idx = 1;

                if (tokens.size() > idx) submenu.label = Trim(tokens[idx++]);
                if (tokens.size() > idx) submenu.iconSource = std::wstring(tokens[idx++]);
                if (tokens.size() > idx) submenu.anchor = ParseContextMenuAnchor(tokens[idx++]);
                if (tokens.size() > idx) submenu.enabled = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.minimumSelection = ParseSelectionCount(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.maximumSelection = ParseSelectionCount(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.showForFiles = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.showForFolders = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.showForMultiple = ParseBool(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.filePatterns = ParseContextMenuExtensions(tokens[idx++]);
                if (tokens.size() > idx) submenu.visibility.excludePatterns = ParseContextMenuExtensions(tokens[idx++]);
                if (tokens.size() > idx) submenu.description = Trim(tokens[idx++]);
                if (tokens.size() > idx) submenu.id = Trim(tokens[idx++]);

                contextMenuStack.back()->push_back(std::move(submenu));
                ContextMenuItem& inserted = contextMenuStack.back()->back();
                contextMenuStack.push_back(&inserted.children);
            }
            return true;
        }

        if (header == kContextMenuSeparatorToken) {
            if (!contextMenuStack.empty()) {
                ContextMenuItem separator;
                separator.type = ContextMenuItemType::kSeparator;
                size_t idx = 1;

                if (tokens.size() > idx) separator.anchor = ParseContextMenuAnchor(tokens[idx++]);
                if (tokens.size() > idx) separator.enabled = ParseBool(tokens[idx++]);

                contextMenuStack.back()->push_back(std::move(separator));
            }
            return true;
        }

        if (header == kContextMenuEndToken) {
            if (contextMenuStack.size() > 1) {
                contextMenuStack.pop_back();
            }
            return true;
        }

        if (header == kTabDockingToken) {
            if (tokens.size() >= 2) {
                m_options.tabDockMode = ParseDockMode(tokens[1]);
            }
            return true;
        }

        return true;
    });

    NormalizeContextMenuItems(&m_options.contextMenuItems);

    if (anyGlowSurfaceToken) {
        ShellTabsOptions fallbackOptions = m_options;
        UpdateLegacyGlowSettingsFromPalette(fallbackOptions);
        UpdateGlowPaletteFromLegacySettings(fallbackOptions);

        for (size_t i = 0; i < glowSurfaceSpecified.size(); ++i) {
            if (glowSurfaceSpecified[i]) {
                continue;
            }
            const GlowSurfaceMapping& mapping = kGlowSurfaceMappings[i];
            GlowSurfaceOptions* target = GetGlowSurfaceOptions(&m_options, mapping);
            const GlowSurfaceOptions* fallback = GetGlowSurfaceOptions(&fallbackOptions, mapping);
            if (target && fallback) {
                *target = *fallback;
            }
        }

        UpdateLegacyGlowSettingsFromPalette(m_options);
    } else {
        UpdateGlowPaletteFromLegacySettings(m_options);
    }

    m_loaded = true;
    if (errorContext) {
        errorContext->clear();
    }
    return true;
}

bool OptionsStore::Save() const {
    if (!EnsureLoaded()) {
        return false;
    }

    ShellTabsOptions persistedOptions = m_options;
    UpdateLegacyGlowSettingsFromPalette(persistedOptions);
    const_cast<OptionsStore*>(this)->m_options = persistedOptions;
    NormalizeContextMenuItems(&const_cast<OptionsStore*>(this)->m_options.contextMenuItems);
    const ShellTabsOptions& options = const_cast<OptionsStore*>(this)->m_options;

    const_cast<OptionsStore*>(this)->m_storagePath = ResolveStoragePath();
    if (m_storagePath.empty()) {
        return false;
    }

    const std::wstring storageDirectory = GetShellTabsDataDirectory();
    if (storageDirectory.empty()) {
        return false;
    }

    const size_t separator = m_storagePath.find_last_of(L"\\/");
    if (separator != std::wstring::npos) {
        std::wstring directory = m_storagePath.substr(0, separator);
        if (!directory.empty()) {
            CreateDirectoryW(directory.c_str(), nullptr);
        }
    }

    std::wstring content = L"version|" + std::to_wstring(kCurrentOptionsVersion) + L"\n";
    content += kReopenToken;
    content += L"|";
    content += options.reopenOnCrash ? L"1" : L"0";
    content += L"\n";
    content += kPersistToken;
    content += L"|";
    content += options.persistGroupPaths ? L"1" : L"0";
    content += L"\n";
    content += kNewTabTemplateToken;
    content += L"|";
    content += NewTabTemplateToString(options.newTabTemplate);
    content += L"\n";
    content += kNewTabCustomPathToken;
    content += L"|";
    content += Trim(options.newTabCustomPath);
    content += L"\n";
    content += kNewTabSavedGroupToken;
    content += L"|";
    content += Trim(options.newTabSavedGroup);
    content += L"\n";
    content += kBreadcrumbGradientToken;
    content += L"|";
    content += options.enableBreadcrumbGradient ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbFontGradientToken;
    content += L"|";
    content += options.enableBreadcrumbFontGradient ? L"1" : L"0";
    content += L"\n";
    content += kBreadcrumbGradientTransparencyToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbGradientTransparency, 0, 100));
    content += L"\n";
    content += kBreadcrumbFontBrightnessToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbFontBrightness, 0, 100));
    content += L"\n";
    content += kBreadcrumbHighlightAlphaMultiplierToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbHighlightAlphaMultiplier, 0, 200));
    content += L"\n";
    content += kBreadcrumbDropdownAlphaMultiplierToken;
    content += L"|";
    content += std::to_wstring(std::clamp(options.breadcrumbDropdownAlphaMultiplier, 0, 200));
    content += L"\n";
    content += kBreadcrumbGradientColorsToken;
    content += L"|";
    content += options.useCustomBreadcrumbGradientColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.breadcrumbGradientStartColor);
    content += L"|";
    content += ColorToHexString(options.breadcrumbGradientEndColor);
    content += L"\n";
    content += kBreadcrumbFontGradientColorsToken;
    content += L"|";
    content += options.useCustomBreadcrumbFontColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.breadcrumbFontGradientStartColor);
    content += L"|";
    content += ColorToHexString(options.breadcrumbFontGradientEndColor);
    content += L"\n";
    content += kProgressGradientColorsToken;
    content += L"|";
    content += options.useCustomProgressBarGradientColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.progressBarGradientStartColor);
    content += L"|";
    content += ColorToHexString(options.progressBarGradientEndColor);
    content += L"\n";
    content += kGlowEnabledToken;
    content += L"|";
    content += options.enableNeonGlow ? L"1" : L"0";
    content += L"\n";
    content += kGlowGradientToken;
    content += L"|";
    content += options.useNeonGlowGradient ? L"1" : L"0";
    content += L"\n";
    content += kGlowColorsToken;
    content += L"|";
    content += options.useCustomNeonGlowColors ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.neonGlowPrimaryColor);
    content += L"|";
    content += ColorToHexString(options.neonGlowSecondaryColor);
    content += L"\n";
    content += kBitmapInterceptToken;
    content += L"|";
    content += options.enableBitmapIntercept ? L"1" : L"0";
    content += L"\n";
    content += kFileGradientFontToken;
    content += L"|";
    content += options.enableFileGradientFont ? L"1" : L"0";
    content += L"\n";
    content += kDirectUiReplacementToken;
    content += L"|";
    content += options.enableDirectUiReplacement ? L"1" : L"0";
    content += L"\n";
    for (const auto& mapping : kGlowSurfaceMappings) {
        const GlowSurfaceOptions* surface = GetGlowSurfaceOptions(&options, mapping);
        if (!surface) {
            continue;
        }
        content += kGlowSurfaceToken;
        content += L"|";
        content += mapping.token;
        content += L"|";
        content += GlowModeToString(surface->mode);
        content += L"|";
        content += ColorToHexString(surface->solidColor);
        content += L"|";
        content += ColorToHexString(surface->gradientStartColor);
        content += L"|";
        content += ColorToHexString(surface->gradientEndColor);
        content += L"|";
        content += surface->enabled ? L"1" : L"0";
        content += L"\n";
    }
    content += kTabSelectedColorToken;
    content += L"|";
    content += options.useCustomTabSelectedColor ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.customTabSelectedColor);
    content += L"\n";
    content += kTabUnselectedColorToken;
    content += L"|";
    content += options.useCustomTabUnselectedColor ? L"1" : L"0";
    content += L"|";
    content += ColorToHexString(options.customTabUnselectedColor);
    content += L"\n";

    content += kExplorerAccentColorsToken;
    content += L"|";
    content += options.useExplorerAccentColors ? L"1" : L"0";
    content += L"\n";

    content += kFolderBackgroundsEnabledToken;
    content += L"|";
    content += options.enableFolderBackgrounds ? L"1" : L"0";
    content += L"\n";

    const auto appendCachedImageLine = [&](const wchar_t* token, const std::wstring& path,
                                           const std::wstring& displayName) {
        const std::wstring normalizedPath = NormalizeCachePath(path, storageDirectory);
        if (normalizedPath.empty()) {
            return;
        }

        content += token;
        content += L"|";
        content += normalizedPath;
        content += L"|";
        content += displayName;
        content += L"\n";
    };

    appendCachedImageLine(kFolderBackgroundUniversalToken, options.universalFolderBackgroundImage.cachedImagePath,
                          options.universalFolderBackgroundImage.displayName);

    for (const auto& entry : options.folderBackgroundEntries) {
        const std::wstring normalizedFolder = NormalizeFileSystemPath(entry.folderPath);
        const std::wstring normalizedCache =
            NormalizeCachePath(entry.image.cachedImagePath, storageDirectory);
        if (normalizedFolder.empty() || normalizedCache.empty()) {
            continue;
        }

        content += kFolderBackgroundEntryToken;
        content += L"|";
        content += normalizedFolder;
        content += L"|";
        content += normalizedCache;
        content += L"|";
        content += entry.image.displayName;
        content += L"\n";
    }

    AppendContextMenuItems(content, options.contextMenuItems);

    content += kTabDockingToken;
    content += L"|";
    content += DockModeToString(options.tabDockMode);
    content += L"\n";

    if (!WriteUtf8File(m_storagePath, content)) {
        return false;
    }

    UpdateCachedImageUsage(options);
    return true;
}

bool operator==(const ContextMenuVisibilityRules& left, const ContextMenuVisibilityRules& right) noexcept {
    return left.minimumSelection == right.minimumSelection &&
           left.maximumSelection == right.maximumSelection &&
           left.showForFiles == right.showForFiles &&
           left.showForFolders == right.showForFolders &&
           left.showForMultiple == right.showForMultiple &&
           left.filePatterns == right.filePatterns &&
           left.excludePatterns == right.excludePatterns;
}

bool operator==(const ContextMenuItem& left, const ContextMenuItem& right) noexcept {
    return left.type == right.type &&
           left.label == right.label &&
           left.iconSource == right.iconSource &&
           left.executable == right.executable &&
           left.arguments == right.arguments &&
           left.workingDirectory == right.workingDirectory &&
           left.windowState == right.windowState &&
           left.runAsAdmin == right.runAsAdmin &&
           left.waitForCompletion == right.waitForCompletion &&
           left.visibility == right.visibility &&
           left.anchor == right.anchor &&
           left.enabled == right.enabled &&
           left.children == right.children &&
           left.description == right.description &&
           left.id == right.id;
}

bool operator==(const CachedImageMetadata& left, const CachedImageMetadata& right) noexcept {
    return left.cachedImagePath == right.cachedImagePath && left.displayName == right.displayName;
}

bool operator==(const FolderBackgroundEntry& left, const FolderBackgroundEntry& right) noexcept {
    return left.folderPath == right.folderPath && left.image == right.image;
}

bool operator==(const GlowSurfaceOptions& left, const GlowSurfaceOptions& right) noexcept {
    return left.enabled == right.enabled && left.mode == right.mode && left.solidColor == right.solidColor &&
           left.gradientStartColor == right.gradientStartColor &&
           left.gradientEndColor == right.gradientEndColor;
}

bool operator==(const GlowSurfacePalette& left, const GlowSurfacePalette& right) noexcept {
    return left.header == right.header && left.listView == right.listView &&
           left.directUi == right.directUi && left.toolbar == right.toolbar && left.rebar == right.rebar &&
           left.edits == right.edits && left.scrollbars == right.scrollbars &&
           left.popupMenus == right.popupMenus && left.tooltips == right.tooltips;
}

bool operator==(const ShellTabsOptions& left, const ShellTabsOptions& right) noexcept {
    return left.reopenOnCrash == right.reopenOnCrash && left.persistGroupPaths == right.persistGroupPaths &&
           left.enableBreadcrumbGradient == right.enableBreadcrumbGradient &&
           left.enableBreadcrumbFontGradient == right.enableBreadcrumbFontGradient &&
           left.breadcrumbGradientTransparency == right.breadcrumbGradientTransparency &&
           left.breadcrumbFontBrightness == right.breadcrumbFontBrightness &&
           left.breadcrumbHighlightAlphaMultiplier == right.breadcrumbHighlightAlphaMultiplier &&
           left.breadcrumbDropdownAlphaMultiplier == right.breadcrumbDropdownAlphaMultiplier &&
           left.useCustomBreadcrumbGradientColors == right.useCustomBreadcrumbGradientColors &&
           left.breadcrumbGradientStartColor == right.breadcrumbGradientStartColor &&
           left.breadcrumbGradientEndColor == right.breadcrumbGradientEndColor &&
           left.useCustomBreadcrumbFontColors == right.useCustomBreadcrumbFontColors &&
           left.breadcrumbFontGradientStartColor == right.breadcrumbFontGradientStartColor &&
           left.breadcrumbFontGradientEndColor == right.breadcrumbFontGradientEndColor &&
           left.useCustomProgressBarGradientColors == right.useCustomProgressBarGradientColors &&
           left.progressBarGradientStartColor == right.progressBarGradientStartColor &&
           left.progressBarGradientEndColor == right.progressBarGradientEndColor &&
           left.enableNeonGlow == right.enableNeonGlow &&
           left.useNeonGlowGradient == right.useNeonGlowGradient &&
           left.useCustomNeonGlowColors == right.useCustomNeonGlowColors &&
           left.neonGlowPrimaryColor == right.neonGlowPrimaryColor &&
           left.neonGlowSecondaryColor == right.neonGlowSecondaryColor &&
           left.enableBitmapIntercept == right.enableBitmapIntercept &&
           left.enableFileGradientFont == right.enableFileGradientFont &&
           left.enableDirectUiReplacement == right.enableDirectUiReplacement &&
           left.useCustomTabSelectedColor == right.useCustomTabSelectedColor &&
           left.customTabSelectedColor == right.customTabSelectedColor &&
           left.useCustomTabUnselectedColor == right.useCustomTabUnselectedColor &&
           left.customTabUnselectedColor == right.customTabUnselectedColor &&
           left.useExplorerAccentColors == right.useExplorerAccentColors &&
           left.glowPalette == right.glowPalette &&
           left.enableFolderBackgrounds == right.enableFolderBackgrounds &&
           left.universalFolderBackgroundImage == right.universalFolderBackgroundImage &&
           left.folderBackgroundEntries == right.folderBackgroundEntries &&
           left.contextMenuItems == right.contextMenuItems &&
           left.tabDockMode == right.tabDockMode &&
           left.newTabTemplate == right.newTabTemplate &&
           left.newTabCustomPath == right.newTabCustomPath &&
           left.newTabSavedGroup == right.newTabSavedGroup;
}

}  // namespace shelltabs

