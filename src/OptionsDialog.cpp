#include "OptionsDialog.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <prsht.h>

#include <algorithm>
#include <array>
#include <cwchar>
#include <malloc.h>
#include <memory>
#include <random>
#include <string>
#include <thread>
#include <utility>
#include <vector>
#include <cstring>
#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <commdlg.h>
#include <wrl/client.h>
#include <objbase.h>

#include "BackgroundCache.h"
#include "GroupStore.h"
#include "Logging.h"
#include "Module.h"
#include "OptionsStore.h"
#include "StringUtils.h"
#include "ShellTabsMessages.h"
#include "TabBandWindow.h"
#include "Utilities.h"

namespace shelltabs {
namespace {

constexpr int kMainCheckboxWidth = 210;
constexpr int kMainDialogWidth = 260;
constexpr int kMainDialogHeight = 430;
constexpr int kGroupDialogWidth = 320;
constexpr int kGroupDialogHeight = 200;
constexpr int kEditorWidth = 340;
constexpr int kEditorHeight = 220;
constexpr int kGlowDialogWidth = 260;
constexpr int kGlowDialogHeight = 260;
constexpr int kGlowCheckboxWidth = 210;
constexpr int kCustomizationScrollLineStep = 16;
constexpr int kCustomizationScrollPageStep = 80;
constexpr SIZE kUniversalPreviewSize = {96, 72};
constexpr SIZE kFolderPreviewSize = {64, 64};
constexpr UINT WM_PREVIEW_BITMAP_READY = WM_APP + 101;
constexpr int kContextDialogWidth = 360;
constexpr int kContextDialogHeight = 430;

struct PreviewBitmapResult {
    UINT64 token = 0;
    HBITMAP bitmap = nullptr;
};

enum ControlIds : int {
    IDC_MAIN_REOPEN = 5001,
    IDC_MAIN_PERSIST = 5002,
    IDC_MAIN_BREADCRUMB = 5003,
    IDC_MAIN_BREADCRUMB_FONT = 5004,
    IDC_MAIN_EXAMPLE = 5005,
    IDC_MAIN_BREADCRUMB_BG_LABEL = 5006,
    IDC_MAIN_BREADCRUMB_BG_SLIDER = 5007,
    IDC_MAIN_BREADCRUMB_BG_VALUE = 5008,
    IDC_MAIN_BREADCRUMB_FONT_LABEL = 5009,
    IDC_MAIN_BREADCRUMB_FONT_SLIDER = 5010,
    IDC_MAIN_BREADCRUMB_FONT_VALUE = 5011,
    IDC_MAIN_BREADCRUMB_HIGHLIGHT_LABEL = 5012,
    IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER = 5013,
    IDC_MAIN_BREADCRUMB_HIGHLIGHT_VALUE = 5014,
    IDC_MAIN_BREADCRUMB_DROPDOWN_LABEL = 5015,
    IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER = 5016,
    IDC_MAIN_BREADCRUMB_DROPDOWN_VALUE = 5017,
    IDC_MAIN_BREADCRUMB_BG_CUSTOM = 5018,
    IDC_MAIN_BREADCRUMB_BG_START_LABEL = 5019,
    IDC_MAIN_BREADCRUMB_BG_START_PREVIEW = 5020,
    IDC_MAIN_BREADCRUMB_BG_START_BUTTON = 5021,
    IDC_MAIN_BREADCRUMB_BG_END_LABEL = 5022,
    IDC_MAIN_BREADCRUMB_BG_END_PREVIEW = 5023,
    IDC_MAIN_BREADCRUMB_BG_END_BUTTON = 5024,
    IDC_MAIN_BREADCRUMB_FONT_CUSTOM = 5025,
    IDC_MAIN_BREADCRUMB_FONT_START_LABEL = 5026,
    IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW = 5027,
    IDC_MAIN_BREADCRUMB_FONT_START_BUTTON = 5028,
    IDC_MAIN_BREADCRUMB_FONT_END_LABEL = 5029,
    IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW = 5030,
    IDC_MAIN_BREADCRUMB_FONT_END_BUTTON = 5031,
    IDC_MAIN_TAB_SELECTED_CHECK = 5032,
    IDC_MAIN_TAB_SELECTED_PREVIEW = 5033,
    IDC_MAIN_TAB_SELECTED_BUTTON = 5034,
    IDC_MAIN_TAB_UNSELECTED_CHECK = 5035,
    IDC_MAIN_TAB_UNSELECTED_PREVIEW = 5036,
    IDC_MAIN_TAB_UNSELECTED_BUTTON = 5037,
    IDC_MAIN_PROGRESS_CUSTOM = 5038,
    IDC_MAIN_PROGRESS_START_LABEL = 5039,
    IDC_MAIN_PROGRESS_START_PREVIEW = 5040,
    IDC_MAIN_PROGRESS_START_BUTTON = 5041,
    IDC_MAIN_PROGRESS_END_LABEL = 5042,
    IDC_MAIN_PROGRESS_END_PREVIEW = 5043,
    IDC_MAIN_PROGRESS_END_BUTTON = 5044,
    IDC_MAIN_NEW_TAB_LABEL = 5045,
    IDC_MAIN_NEW_TAB_COMBO = 5046,
    IDC_MAIN_NEW_TAB_PATH_LABEL = 5047,
    IDC_MAIN_NEW_TAB_PATH_EDIT = 5048,
    IDC_MAIN_NEW_TAB_BROWSE = 5049,
    IDC_MAIN_NEW_TAB_GROUP_LABEL = 5050,
    IDC_MAIN_NEW_TAB_GROUP_COMBO = 5051,
    IDC_MAIN_DOCK_LABEL = 5052,
    IDC_MAIN_DOCK_COMBO = 5053,
    IDC_MAIN_LISTVIEW_ACCENT = 5054,

    IDC_CUSTOM_BACKGROUND_ENABLE = 5301,
    IDC_CUSTOM_BACKGROUND_BROWSE = 5302,
    IDC_CUSTOM_BACKGROUND_PREVIEW = 5303,
    IDC_CUSTOM_BACKGROUND_UNIVERSAL_NAME = 5304,
    IDC_CUSTOM_BACKGROUND_LIST = 5305,
    IDC_CUSTOM_BACKGROUND_ADD = 5306,
    IDC_CUSTOM_BACKGROUND_EDIT = 5307,
    IDC_CUSTOM_BACKGROUND_REMOVE = 5308,
    IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW = 5309,
    IDC_CUSTOM_BACKGROUND_FOLDER_NAME = 5310,
    IDC_CUSTOM_BACKGROUND_CLEAN = 5311,

    IDC_GLOW_ENABLE = 5401,
    IDC_GLOW_CUSTOM_COLORS = 5402,
    IDC_GLOW_USE_GRADIENT = 5403,
    IDC_GLOW_PRIMARY_LABEL = 5404,
    IDC_GLOW_PRIMARY_PREVIEW = 5405,
    IDC_GLOW_PRIMARY_BUTTON = 5406,
    IDC_GLOW_SECONDARY_LABEL = 5407,
    IDC_GLOW_SECONDARY_PREVIEW = 5408,
    IDC_GLOW_SECONDARY_BUTTON = 5409,
    IDC_GLOW_SURFACE_LISTVIEW = 5410,
    IDC_GLOW_SURFACE_HEADER = 5411,
    IDC_GLOW_SURFACE_REBAR = 5412,
    IDC_GLOW_SURFACE_TOOLBAR = 5413,
    IDC_GLOW_SURFACE_EDIT = 5414,
    IDC_GLOW_SURFACE_DIRECTUI = 5415,

    IDC_GROUP_LIST = 5101,
    IDC_GROUP_NEW = 5102,
    IDC_GROUP_EDIT = 5103,
    IDC_GROUP_REMOVE = 5104,

    IDC_EDITOR_NAME = 5201,
    IDC_EDITOR_PATH_LIST = 5202,
    IDC_EDITOR_ADD_PATH = 5203,
    IDC_EDITOR_EDIT_PATH = 5204,
    IDC_EDITOR_REMOVE_PATH = 5205,
    IDC_EDITOR_COLOR_PREVIEW = 5206,
    IDC_EDITOR_COLOR_BUTTON = 5207,
    IDC_EDITOR_STYLE_LABEL = 5208,
    IDC_EDITOR_STYLE_COMBO = 5209,

    IDC_CONTEXT_TREE = 5501,
    IDC_CONTEXT_ADD_COMMAND = 5502,
    IDC_CONTEXT_ADD_SUBMENU = 5503,
    IDC_CONTEXT_ADD_SEPARATOR = 5504,
    IDC_CONTEXT_REMOVE = 5505,
    IDC_CONTEXT_MOVE_UP = 5506,
    IDC_CONTEXT_MOVE_DOWN = 5507,
    IDC_CONTEXT_INDENT = 5508,
    IDC_CONTEXT_OUTDENT = 5509,
    IDC_CONTEXT_LABEL_EDIT = 5510,
    IDC_CONTEXT_ICON_EDIT = 5511,
    IDC_CONTEXT_ICON_BROWSE = 5512,
    IDC_CONTEXT_COMMAND_PATH = 5513,
    IDC_CONTEXT_COMMAND_BROWSE = 5514,
    IDC_CONTEXT_COMMAND_ARGS = 5515,
    IDC_CONTEXT_HINTS_STATIC = 5516,
    IDC_CONTEXT_SELECTION_MIN = 5517,
    IDC_CONTEXT_SELECTION_MAX = 5518,
    IDC_CONTEXT_ANCHOR_COMBO = 5519,
    IDC_CONTEXT_SCOPE_FILES = 5520,
    IDC_CONTEXT_SCOPE_FOLDERS = 5521,
    IDC_CONTEXT_SEPARATOR_CHECK = 5522,
    IDC_CONTEXT_EXTENSION_EDIT = 5523,
    IDC_CONTEXT_EXTENSION_ADD = 5524,
    IDC_CONTEXT_EXTENSION_LIST = 5525,
    IDC_CONTEXT_EXTENSION_REMOVE = 5526,
};

struct GlowSurfaceControlMapping {
    int controlId = 0;
    GlowSurfaceOptions GlowSurfacePalette::*member = nullptr;
};

constexpr std::array<GlowSurfaceControlMapping, 6> kGlowSurfaceControlMappings = {{
    {IDC_GLOW_SURFACE_LISTVIEW, &GlowSurfacePalette::listView},
    {IDC_GLOW_SURFACE_HEADER, &GlowSurfacePalette::header},
    {IDC_GLOW_SURFACE_REBAR, &GlowSurfacePalette::rebar},
    {IDC_GLOW_SURFACE_TOOLBAR, &GlowSurfacePalette::toolbar},
    {IDC_GLOW_SURFACE_EDIT, &GlowSurfacePalette::edits},
    {IDC_GLOW_SURFACE_DIRECTUI, &GlowSurfacePalette::directUi},
}};

struct ChildPlacement {
    HWND hwnd = nullptr;
    RECT rect{};
};

struct OptionsDialogData {
    ShellTabsOptions originalOptions;
    ShellTabsOptions workingOptions;
    bool applyInvoked = false;
    bool groupsChanged = false;
    bool previewOptionsBroadcasted = false;
    int initialTab = 0;
    std::vector<SavedGroup> originalGroups;
    std::vector<SavedGroup> workingGroups;
    std::vector<std::wstring> workingGroupIds;
    std::vector<std::wstring> removedGroupIds;
    HBRUSH breadcrumbBgStartBrush = nullptr;
    HBRUSH breadcrumbBgEndBrush = nullptr;
    HBRUSH breadcrumbFontStartBrush = nullptr;
    HBRUSH breadcrumbFontEndBrush = nullptr;
    HBRUSH progressStartBrush = nullptr;
    HBRUSH progressEndBrush = nullptr;
    HBRUSH tabSelectedBrush = nullptr;
    HBRUSH tabUnselectedBrush = nullptr;
    HBRUSH glowPrimaryBrush = nullptr;
    HBRUSH glowSecondaryBrush = nullptr;
    HBITMAP universalBackgroundPreview = nullptr;
    HBITMAP folderBackgroundPreview = nullptr;
    UINT64 universalPreviewToken = 0;
    UINT64 folderPreviewToken = 0;
    std::wstring lastFolderBrowsePath;
    std::wstring lastImageBrowseDirectory;
    std::vector<std::wstring> createdCachedImagePaths;
    std::vector<std::wstring> pendingCachedImageRemovals;
    std::vector<ChildPlacement> customizationChildPlacements;
    int customizationScrollPos = 0;
    int customizationContentHeight = 0;
    int customizationScrollMax = 0;
    int customizationWheelRemainder = 0;
    std::vector<std::vector<size_t>> contextTreePaths;
    std::vector<HTREEITEM> contextTreeItems;
    std::vector<size_t> contextSelectionPath;
    bool contextSelectionValid = false;
    bool contextUpdatingControls = false;
    std::wstring contextCommandBrowseDirectory;
    std::wstring focusSavedGroupId;
    bool focusShouldEdit = false;
    bool focusHandled = false;
};

std::wstring GetWindowTextString(HWND control) {
    std::wstring text;
    if (!control) {
        return text;
    }
    const int length = GetWindowTextLengthW(control);
    if (length <= 0) {
        return text;
    }
    text.resize(static_cast<size_t>(length) + 1);
    const int copied = GetWindowTextW(control, text.data(), length + 1);
    if (copied >= 0) {
        text.resize(static_cast<size_t>(copied));
    } else {
        text.clear();
    }
    return text;
}

std::wstring DescribeContextMenuTreeItem(const ContextMenuItem& item) {
    switch (item.type) {
        case ContextMenuItemType::kCommand:
            if (!item.label.empty()) {
                return item.label;
            }
            return L"(Command)";
        case ContextMenuItemType::kSubmenu:
            if (!item.label.empty()) {
                return item.label + L" (submenu)";
            }
            return L"(Submenu)";
        case ContextMenuItemType::kSeparator:
        default:
            return L"(Separator)";
    }
}

std::vector<ContextMenuItem>* GetContextMenuContainer(std::vector<ContextMenuItem>& root,
                                                      const std::vector<size_t>& path) {
    std::vector<ContextMenuItem>* container = &root;
    if (path.empty()) {
        return container;
    }
    for (size_t i = 0; i + 1 < path.size(); ++i) {
        if (!container || path[i] >= container->size()) {
            return nullptr;
        }
        ContextMenuItem& parent = (*container)[path[i]];
        container = &parent.children;
    }
    return container;
}

const std::vector<ContextMenuItem>* GetContextMenuContainer(const std::vector<ContextMenuItem>& root,
                                                           const std::vector<size_t>& path) {
    const std::vector<ContextMenuItem>* container = &root;
    if (path.empty()) {
        return container;
    }
    for (size_t i = 0; i + 1 < path.size(); ++i) {
        if (!container || path[i] >= container->size()) {
            return nullptr;
        }
        const ContextMenuItem& parent = (*container)[path[i]];
        container = &parent.children;
    }
    return container;
}

ContextMenuItem* GetContextMenuItem(std::vector<ContextMenuItem>& root, const std::vector<size_t>& path) {
    if (path.empty()) {
        return nullptr;
    }
    auto* container = GetContextMenuContainer(root, path);
    if (!container) {
        return nullptr;
    }
    const size_t index = path.back();
    if (index >= container->size()) {
        return nullptr;
    }
    return &(*container)[index];
}

const ContextMenuItem* GetContextMenuItem(const std::vector<ContextMenuItem>& root,
                                          const std::vector<size_t>& path) {
    if (path.empty()) {
        return nullptr;
    }
    const auto* container = GetContextMenuContainer(root, path);
    if (!container) {
        return nullptr;
    }
    const size_t index = path.back();
    if (index >= container->size()) {
        return nullptr;
    }
    return &(*container)[index];
}

void InsertContextMenuTreeItems(HWND tree, HTREEITEM parent, const std::vector<ContextMenuItem>& items,
                                std::vector<size_t>* currentPath, OptionsDialogData* data) {
    if (!tree || !currentPath || !data) {
        return;
    }
    for (size_t i = 0; i < items.size(); ++i) {
        currentPath->push_back(i);
        const std::wstring label = DescribeContextMenuTreeItem(items[i]);
        data->contextTreePaths.push_back(*currentPath);
        data->contextTreeItems.push_back(nullptr);
        const size_t pathIndex = data->contextTreePaths.size() - 1;

        TVINSERTSTRUCTW insert{};
        insert.hParent = parent;
        insert.hInsertAfter = TVI_LAST;
        insert.item.mask = TVIF_TEXT | TVIF_PARAM;
        insert.item.lParam = static_cast<LPARAM>(pathIndex);
        insert.item.pszText = const_cast<wchar_t*>(label.c_str());
        HTREEITEM handle = TreeView_InsertItem(tree, &insert);
        if (pathIndex < data->contextTreeItems.size()) {
            data->contextTreeItems[pathIndex] = handle;
        }

        if (items[i].type == ContextMenuItemType::kSubmenu && !items[i].children.empty()) {
            InsertContextMenuTreeItems(tree, handle, items[i].children, currentPath, data);
            TreeView_Expand(tree, handle, TVE_EXPAND);
        }
        currentPath->pop_back();
    }
}

HTREEITEM FindContextTreeItem(const OptionsDialogData* data, const std::vector<size_t>& path) {
    if (!data) {
        return nullptr;
    }
    for (size_t i = 0; i < data->contextTreePaths.size() && i < data->contextTreeItems.size(); ++i) {
        if (data->contextTreePaths[i] == path) {
            return data->contextTreeItems[i];
        }
    }
    return nullptr;
}

void RefreshContextMenuTree(HWND page, OptionsDialogData* data, const std::vector<size_t>* selectionPath) {
    if (!data) {
        return;
    }
    HWND tree = GetDlgItem(page, IDC_CONTEXT_TREE);
    if (!tree) {
        return;
    }

    data->contextTreePaths.clear();
    data->contextTreeItems.clear();

    TreeView_DeleteAllItems(tree);

    std::vector<size_t> path;
    InsertContextMenuTreeItems(tree, TVI_ROOT, data->workingOptions.contextMenuItems, &path, data);

    if (selectionPath) {
        if (HTREEITEM item = FindContextTreeItem(data, *selectionPath)) {
            TreeView_SelectItem(tree, item);
        }
    } else if (!data->contextTreeItems.empty()) {
        TreeView_SelectItem(tree, data->contextTreeItems.front());
    }
}

bool GetContextMenuSelectedPath(HWND page, OptionsDialogData* data, std::vector<size_t>* path) {
    if (!data || !path) {
        return false;
    }
    HWND tree = GetDlgItem(page, IDC_CONTEXT_TREE);
    if (!tree) {
        return false;
    }
    HTREEITEM selection = TreeView_GetSelection(tree);
    if (!selection) {
        return false;
    }
    TVITEMW item{};
    item.mask = TVIF_PARAM;
    item.hItem = selection;
    if (!TreeView_GetItem(tree, &item)) {
        return false;
    }
    const size_t index = static_cast<size_t>(item.lParam);
    if (index >= data->contextTreePaths.size()) {
        return false;
    }
    *path = data->contextTreePaths[index];
    return true;
}

void PopulateContextMenuAnchorCombo(HWND combo) {
    if (!combo) {
        return;
    }
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    struct AnchorOption {
        ContextMenuInsertionAnchor value;
        const wchar_t* label;
    } options[] = {
        {ContextMenuInsertionAnchor::kDefault, L"Default"},
        {ContextMenuInsertionAnchor::kTop, L"Top"},
        {ContextMenuInsertionAnchor::kBottom, L"Bottom"},
        {ContextMenuInsertionAnchor::kBeforeShellItems, L"Before shell items"},
        {ContextMenuInsertionAnchor::kAfterShellItems, L"After shell items"},
    };
    for (const auto& option : options) {
        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0,
                                                        reinterpret_cast<LPARAM>(option.label)));
        if (index >= 0) {
            SendMessageW(combo, CB_SETITEMDATA, index, static_cast<LPARAM>(option.value));
        }
    }
}

void SplitCommandTemplate(const std::wstring& commandTemplate, std::wstring* path, std::wstring* args) {
    if (path) {
        path->clear();
    }
    if (args) {
        args->clear();
    }
    std::wstring trimmed = Trim(commandTemplate);
    if (trimmed.empty()) {
        return;
    }
    std::vector<wchar_t> buffer(trimmed.begin(), trimmed.end());
    buffer.push_back(L'\0');
    wchar_t* argStart = PathGetArgsW(buffer.data());
    size_t argOffset = static_cast<size_t>(argStart - buffer.data());
    std::wstring commandPath(buffer.data(), buffer.data() + argOffset);
    commandPath = Trim(commandPath);
    if (!commandPath.empty() && commandPath.front() == L'"' && commandPath.back() == L'"') {
        commandPath = commandPath.substr(1, commandPath.size() - 2);
    }
    if (path) {
        *path = commandPath;
    }
    if (args) {
        *args = Trim(std::wstring(argStart));
    }
}

std::wstring BuildCommandTemplate(const std::wstring& path, const std::wstring& args) {
    std::wstring trimmedPath = Trim(path);
    std::wstring trimmedArgs = Trim(args);
    if (trimmedPath.empty()) {
        return Trim(trimmedArgs);
    }
    bool quoted = !trimmedPath.empty() && trimmedPath.front() == L'"' && trimmedPath.back() == L'"';
    bool needsQuotes = trimmedPath.find(L' ') != std::wstring::npos || trimmedPath.find(L'\t') != std::wstring::npos;
    std::wstring command = trimmedPath;
    if (needsQuotes && !quoted) {
        command = L"\"" + trimmedPath + L"\"";
    }
    if (!trimmedArgs.empty()) {
        if (!command.empty()) {
            command.push_back(L' ');
        }
        command += trimmedArgs;
    }
    return command;
}

std::vector<std::wstring> CollectExtensionsFromList(HWND page) {
    std::vector<std::wstring> extensions;
    HWND list = GetDlgItem(page, IDC_CONTEXT_EXTENSION_LIST);
    if (!list) {
        return extensions;
    }
    const int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    for (int i = 0; i < count; ++i) {
        const int length = static_cast<int>(SendMessageW(list, LB_GETTEXTLEN, i, 0));
        if (length <= 0) {
            continue;
        }
        std::wstring buffer(static_cast<size_t>(length) + 1, L'\0');
        if (SendMessageW(list, LB_GETTEXT, i, reinterpret_cast<LPARAM>(buffer.data())) != LB_ERR) {
            buffer.resize(static_cast<size_t>(length));
            extensions.emplace_back(std::move(buffer));
        }
    }
    return extensions;
}

void RefreshContextMenuExtensionsList(HWND page, const ContextMenuItem& item) {
    HWND list = GetDlgItem(page, IDC_CONTEXT_EXTENSION_LIST);
    if (!list) {
        return;
    }
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    for (const auto& ext : item.scope.extensions) {
        SendMessageW(list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(ext.c_str()));
    }
}

void UpdateContextMenuTreeItemText(HWND page, OptionsDialogData* data, const ContextMenuItem& item) {
    if (!data || !data->contextSelectionValid) {
        return;
    }
    HWND tree = GetDlgItem(page, IDC_CONTEXT_TREE);
    if (!tree) {
        return;
    }
    HTREEITEM selection = TreeView_GetSelection(tree);
    if (!selection) {
        return;
    }
    std::wstring label = DescribeContextMenuTreeItem(item);
    TVITEMW update{};
    update.mask = TVIF_TEXT;
    update.hItem = selection;
    update.pszText = label.data();
    update.cchTextMax = static_cast<int>(label.size());
    TreeView_SetItem(tree, &update);
}

void UpdateContextMenuButtonStates(HWND page, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND addCommand = GetDlgItem(page, IDC_CONTEXT_ADD_COMMAND);
    HWND addSubmenu = GetDlgItem(page, IDC_CONTEXT_ADD_SUBMENU);
    HWND addSeparator = GetDlgItem(page, IDC_CONTEXT_ADD_SEPARATOR);
    HWND removeButton = GetDlgItem(page, IDC_CONTEXT_REMOVE);
    HWND moveUp = GetDlgItem(page, IDC_CONTEXT_MOVE_UP);
    HWND moveDown = GetDlgItem(page, IDC_CONTEXT_MOVE_DOWN);
    HWND indent = GetDlgItem(page, IDC_CONTEXT_INDENT);
    HWND outdent = GetDlgItem(page, IDC_CONTEXT_OUTDENT);
    HWND groupCheck = GetDlgItem(page, IDC_CONTEXT_SEPARATOR_CHECK);

    std::vector<size_t> path;
    const bool hasSelection = GetContextMenuSelectedPath(page, data, &path);
    const auto* container = hasSelection ?
                                 GetContextMenuContainer(data->workingOptions.contextMenuItems, path) :
                                 nullptr;
    const ContextMenuItem* item = hasSelection ?
                                      GetContextMenuItem(data->workingOptions.contextMenuItems, path) :
                                      nullptr;

    if (addCommand) {
        EnableWindow(addCommand, TRUE);
    }
    if (addSubmenu) {
        EnableWindow(addSubmenu, TRUE);
    }
    if (addSeparator) {
        EnableWindow(addSeparator, TRUE);
    }
    if (removeButton) {
        EnableWindow(removeButton, hasSelection);
    }
    if (moveUp) {
        bool canMove = false;
        if (container && !path.empty()) {
            canMove = path.back() > 0;
        }
        EnableWindow(moveUp, canMove ? TRUE : FALSE);
    }
    if (moveDown) {
        bool canMove = false;
        if (container && !path.empty()) {
            canMove = path.back() + 1 < container->size();
        }
        EnableWindow(moveDown, canMove ? TRUE : FALSE);
    }
    if (indent) {
        bool canIndent = false;
        if (container && !path.empty() && path.back() > 0) {
            const size_t siblingIndex = path.back() - 1;
            if (siblingIndex < container->size()) {
                const ContextMenuItem& sibling = (*container)[siblingIndex];
                canIndent = sibling.type == ContextMenuItemType::kSubmenu;
            }
        }
        EnableWindow(indent, canIndent ? TRUE : FALSE);
    }
    if (outdent) {
        bool canOutdent = path.size() >= 2;
        EnableWindow(outdent, canOutdent ? TRUE : FALSE);
    }
    if (groupCheck) {
        const bool enableGroup = hasSelection && item && item->type != ContextMenuItemType::kSeparator;
        EnableWindow(groupCheck, enableGroup ? TRUE : FALSE);
    }
}

bool HasSeparatorAbove(const std::vector<ContextMenuItem>& root, const std::vector<size_t>& path) {
    if (path.empty()) {
        return false;
    }
    const auto* container = GetContextMenuContainer(root, path);
    if (!container) {
        return false;
    }
    const size_t index = path.back();
    if (index == 0 || index > container->size()) {
        return false;
    }
    const ContextMenuItem& previous = (*container)[index - 1];
    return previous.type == ContextMenuItemType::kSeparator;
}

bool ToggleSeparatorAbove(HWND page, OptionsDialogData* data, bool ensure) {
    if (!data || !data->contextSelectionValid) {
        return false;
    }
    auto path = data->contextSelectionPath;
    auto* container = GetContextMenuContainer(data->workingOptions.contextMenuItems, path);
    if (!container || path.empty()) {
        return false;
    }
    const size_t index = path.back();
    if (ensure) {
        if (index > 0 && (*container)[index - 1].type == ContextMenuItemType::kSeparator) {
            return false;
        }
        ContextMenuItem separator;
        separator.type = ContextMenuItemType::kSeparator;
        container->insert(container->begin() + static_cast<std::ptrdiff_t>(index), separator);
        path.back() = index + 1;
        data->contextSelectionPath = path;
    } else {
        if (index == 0 || (*container)[index - 1].type != ContextMenuItemType::kSeparator) {
            return false;
        }
        container->erase(container->begin() + static_cast<std::ptrdiff_t>(index - 1));
        path.back() = index - 1;
        data->contextSelectionPath = path;
    }
    RefreshContextMenuTree(page, data, &data->contextSelectionPath);
    UpdateContextMenuButtonStates(page, data);
    PopulateContextMenuDetailControls(page, data);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

void PopulateContextMenuDetailControls(HWND page, OptionsDialogData* data) {
    if (!data) {
        return;
    }

    std::vector<size_t> path;
    const bool hasSelection = GetContextMenuSelectedPath(page, data, &path);
    data->contextSelectionValid = hasSelection;
    if (hasSelection) {
        data->contextSelectionPath = path;
    }

    const ContextMenuItem* item =
        hasSelection ? GetContextMenuItem(data->workingOptions.contextMenuItems, path) : nullptr;

    data->contextUpdatingControls = true;

    auto setEditText = [](HWND control, const std::wstring& text) {
        if (control) {
            SetWindowTextW(control, text.c_str());
        }
    };

    HWND labelEdit = GetDlgItem(page, IDC_CONTEXT_LABEL_EDIT);
    HWND iconEdit = GetDlgItem(page, IDC_CONTEXT_ICON_EDIT);
    HWND iconBrowse = GetDlgItem(page, IDC_CONTEXT_ICON_BROWSE);
    HWND commandPath = GetDlgItem(page, IDC_CONTEXT_COMMAND_PATH);
    HWND commandArgs = GetDlgItem(page, IDC_CONTEXT_COMMAND_ARGS);
    HWND commandBrowse = GetDlgItem(page, IDC_CONTEXT_COMMAND_BROWSE);
    HWND hintsStatic = GetDlgItem(page, IDC_CONTEXT_HINTS_STATIC);
    HWND minEdit = GetDlgItem(page, IDC_CONTEXT_SELECTION_MIN);
    HWND maxEdit = GetDlgItem(page, IDC_CONTEXT_SELECTION_MAX);
    HWND anchorCombo = GetDlgItem(page, IDC_CONTEXT_ANCHOR_COMBO);
    HWND scopeFiles = GetDlgItem(page, IDC_CONTEXT_SCOPE_FILES);
    HWND scopeFolders = GetDlgItem(page, IDC_CONTEXT_SCOPE_FOLDERS);
    HWND separatorCheck = GetDlgItem(page, IDC_CONTEXT_SEPARATOR_CHECK);
    HWND extensionEdit = GetDlgItem(page, IDC_CONTEXT_EXTENSION_EDIT);
    HWND extensionAdd = GetDlgItem(page, IDC_CONTEXT_EXTENSION_ADD);
    HWND extensionRemove = GetDlgItem(page, IDC_CONTEXT_EXTENSION_REMOVE);
    HWND extensionList = GetDlgItem(page, IDC_CONTEXT_EXTENSION_LIST);

    if (!hasSelection || !item) {
        setEditText(labelEdit, L"");
        setEditText(iconEdit, L"");
        setEditText(commandPath, L"");
        setEditText(commandArgs, L"");
        setEditText(minEdit, L"0");
        setEditText(maxEdit, L"0");
        if (scopeFiles) {
            Button_SetCheck(scopeFiles, BST_UNCHECKED);
        }
        if (scopeFolders) {
            Button_SetCheck(scopeFolders, BST_UNCHECKED);
        }
        if (separatorCheck) {
            Button_SetCheck(separatorCheck, BST_UNCHECKED);
        }
        if (extensionList) {
            SendMessageW(extensionList, LB_RESETCONTENT, 0, 0);
        }
        if (anchorCombo) {
            SendMessageW(anchorCombo, CB_SETCURSEL, 0, 0);
        }
    } else {
        setEditText(labelEdit, item->label);
        setEditText(iconEdit, item->iconSource);
        std::wstring commandExecutable;
        std::wstring arguments;
        SplitCommandTemplate(item->commandTemplate, &commandExecutable, &arguments);
        setEditText(commandPath, commandExecutable);
        setEditText(commandArgs, arguments);

        if (minEdit) {
            SetWindowTextW(minEdit, std::to_wstring(std::max(item->selection.minimumSelection, 0)).c_str());
        }
        if (maxEdit) {
            SetWindowTextW(maxEdit,
                           std::to_wstring(item->selection.maximumSelection > 0 ? item->selection.maximumSelection
                                                                               : 0)
                               .c_str());
        }
        if (scopeFiles) {
            Button_SetCheck(scopeFiles, item->scope.includeAllFiles ? BST_CHECKED : BST_UNCHECKED);
        }
        if (scopeFolders) {
            Button_SetCheck(scopeFolders, item->scope.includeAllFolders ? BST_CHECKED : BST_UNCHECKED);
        }
        if (separatorCheck) {
            Button_SetCheck(separatorCheck,
                            HasSeparatorAbove(data->workingOptions.contextMenuItems, path) ? BST_CHECKED
                                                                                        : BST_UNCHECKED);
        }
        RefreshContextMenuExtensionsList(page, *item);
        if (extensionList && !item->scope.extensions.empty()) {
            SendMessageW(extensionList, LB_SETCURSEL, 0, 0);
        }

        if (anchorCombo) {
            const int count = static_cast<int>(SendMessageW(anchorCombo, CB_GETCOUNT, 0, 0));
            for (int i = 0; i < count; ++i) {
                const LRESULT value = SendMessageW(anchorCombo, CB_GETITEMDATA, i, 0);
                if (value == static_cast<LRESULT>(item->anchor)) {
                    SendMessageW(anchorCombo, CB_SETCURSEL, i, 0);
                    break;
                }
            }
        }
    }

    const bool isCommand = item && item->type == ContextMenuItemType::kCommand;
    const bool isSeparator = item && item->type == ContextMenuItemType::kSeparator;

    EnableWindow(labelEdit, item && !isSeparator);
    EnableWindow(iconEdit, item && !isSeparator);
    EnableWindow(iconBrowse, item && !isSeparator);
    EnableWindow(commandPath, isCommand);
    EnableWindow(commandArgs, isCommand);
    EnableWindow(commandBrowse, isCommand);
    if (hintsStatic) {
        ShowWindow(hintsStatic, isCommand ? SW_SHOWNOACTIVATE : SW_HIDE);
    }
    EnableWindow(minEdit, item != nullptr);
    EnableWindow(maxEdit, item != nullptr);
    EnableWindow(anchorCombo, item != nullptr);
    EnableWindow(scopeFiles, item != nullptr);
    EnableWindow(scopeFolders, item != nullptr);
    EnableWindow(extensionEdit, item != nullptr);
    EnableWindow(extensionAdd, item != nullptr);
    EnableWindow(extensionRemove, item && !item->scope.extensions.empty());

    data->contextUpdatingControls = false;
}

bool ApplyContextMenuDetailsFromControls(HWND page, OptionsDialogData* data, bool markChanged) {
    if (!data || !data->contextSelectionValid || data->contextUpdatingControls) {
        return false;
    }
    ContextMenuItem* item =
        GetContextMenuItem(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!item) {
        return false;
    }

    bool changed = false;
    std::wstring label = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_LABEL_EDIT)));
    if (item->type != ContextMenuItemType::kSeparator && item->label != label) {
        item->label = label;
        changed = true;
    }

    std::wstring iconSource = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_ICON_EDIT)));
    if (item->type != ContextMenuItemType::kSeparator && item->iconSource != iconSource) {
        item->iconSource = std::move(iconSource);
        changed = true;
    }

    if (item->type == ContextMenuItemType::kCommand) {
        std::wstring commandPath = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_COMMAND_PATH)));
        std::wstring commandArgs = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_COMMAND_ARGS)));
        std::wstring commandTemplate = BuildCommandTemplate(commandPath, commandArgs);
        if (item->commandTemplate != commandTemplate) {
            item->commandTemplate = std::move(commandTemplate);
            changed = true;
        }
    }

    if (item->type != ContextMenuItemType::kSeparator) {
        bool includeFiles = Button_GetCheck(GetDlgItem(page, IDC_CONTEXT_SCOPE_FILES)) == BST_CHECKED;
        bool includeFolders = Button_GetCheck(GetDlgItem(page, IDC_CONTEXT_SCOPE_FOLDERS)) == BST_CHECKED;
        if (item->scope.includeAllFiles != includeFiles) {
            item->scope.includeAllFiles = includeFiles;
            changed = true;
        }
        if (item->scope.includeAllFolders != includeFolders) {
            item->scope.includeAllFolders = includeFolders;
            changed = true;
        }

        std::vector<std::wstring> extensions = CollectExtensionsFromList(page);
        std::vector<std::wstring> normalized = NormalizeContextMenuExtensions(extensions);
        if (item->scope.extensions != normalized) {
            item->scope.extensions = std::move(normalized);
            changed = true;
        }
    }

    int minSelection = GetDlgItemInt(page, IDC_CONTEXT_SELECTION_MIN, nullptr, FALSE);
    if (minSelection < 0) {
        minSelection = 0;
    }
    if (item->selection.minimumSelection != minSelection) {
        item->selection.minimumSelection = minSelection;
        changed = true;
    }

    int maxSelection = GetDlgItemInt(page, IDC_CONTEXT_SELECTION_MAX, nullptr, FALSE);
    if (maxSelection < 0) {
        maxSelection = 0;
    }
    if (item->selection.maximumSelection != maxSelection) {
        item->selection.maximumSelection = maxSelection;
        changed = true;
    }

    HWND anchorCombo = GetDlgItem(page, IDC_CONTEXT_ANCHOR_COMBO);
    if (anchorCombo) {
        const int selection = static_cast<int>(SendMessageW(anchorCombo, CB_GETCURSEL, 0, 0));
        if (selection >= 0) {
            const LRESULT value = SendMessageW(anchorCombo, CB_GETITEMDATA, selection, 0);
            ContextMenuInsertionAnchor anchor = static_cast<ContextMenuInsertionAnchor>(value);
            if (item->anchor != anchor) {
                item->anchor = anchor;
                changed = true;
            }
        }
    }

    if (changed) {
        UpdateContextMenuTreeItemText(page, data, *item);
        if (markChanged) {
            PropSheet_Changed(GetParent(page), page);
        }
    }
    return changed;
}

ContextMenuItem CreateContextMenuItem(ContextMenuItemType type) {
    ContextMenuItem item;
    item.type = type;
    if (type == ContextMenuItemType::kCommand) {
        item.label = L"New Command";
    } else if (type == ContextMenuItemType::kSubmenu) {
        item.label = L"New Submenu";
    }
    return item;
}

bool HandleContextMenuAddItem(HWND page, OptionsDialogData* data, ContextMenuItemType type) {
    if (!data) {
        return false;
    }
    std::vector<size_t> path;
    bool hasSelection = GetContextMenuSelectedPath(page, data, &path);
    ContextMenuItem newItem = CreateContextMenuItem(type);

    if (hasSelection) {
        auto* container = GetContextMenuContainer(data->workingOptions.contextMenuItems, path);
        ContextMenuItem* selectedItem =
            GetContextMenuItem(data->workingOptions.contextMenuItems, path);
        if (!container || !selectedItem) {
            return false;
        }
        if (selectedItem->type == ContextMenuItemType::kSubmenu) {
            selectedItem->children.push_back(std::move(newItem));
            data->contextSelectionPath = path;
            data->contextSelectionPath.push_back(selectedItem->children.size() - 1);
        } else {
            const size_t insertIndex = path.back() + 1;
            container->insert(container->begin() + static_cast<std::ptrdiff_t>(insertIndex), std::move(newItem));
            data->contextSelectionPath = path;
            data->contextSelectionPath.back() = insertIndex;
        }
    } else {
        data->workingOptions.contextMenuItems.push_back(std::move(newItem));
        data->contextSelectionPath = {data->workingOptions.contextMenuItems.size() - 1};
        hasSelection = true;
    }

    data->contextSelectionValid = hasSelection;
    RefreshContextMenuTree(page, data, hasSelection ? &data->contextSelectionPath : nullptr);
    UpdateContextMenuButtonStates(page, data);
    PopulateContextMenuDetailControls(page, data);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool HandleContextMenuRemoveItem(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid) {
        return false;
    }
    auto* container = GetContextMenuContainer(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!container || data->contextSelectionPath.empty()) {
        return false;
    }
    const size_t index = data->contextSelectionPath.back();
    container->erase(container->begin() + static_cast<std::ptrdiff_t>(index));
    if (index > 0) {
        data->contextSelectionPath.back() = index - 1;
    } else if (!data->contextSelectionPath.empty()) {
        if (container->empty()) {
            data->contextSelectionValid = false;
            data->contextSelectionPath.clear();
        } else {
            data->contextSelectionPath.back() = 0;
        }
    }
    RefreshContextMenuTree(page, data,
                           data->contextSelectionValid ? &data->contextSelectionPath : nullptr);
    UpdateContextMenuButtonStates(page, data);
    PopulateContextMenuDetailControls(page, data);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool MoveContextMenuItem(HWND page, OptionsDialogData* data, bool moveUp) {
    if (!data || !data->contextSelectionValid || data->contextSelectionPath.empty()) {
        return false;
    }
    auto* container = GetContextMenuContainer(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!container) {
        return false;
    }
    size_t index = data->contextSelectionPath.back();
    if ((moveUp && index == 0) || (!moveUp && index + 1 >= container->size())) {
        return false;
    }
    size_t swapIndex = moveUp ? index - 1 : index + 1;
    std::swap((*container)[index], (*container)[swapIndex]);
    data->contextSelectionPath.back() = swapIndex;
    RefreshContextMenuTree(page, data, &data->contextSelectionPath);
    UpdateContextMenuButtonStates(page, data);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool IndentContextMenuItem(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid || data->contextSelectionPath.empty()) {
        return false;
    }
    auto* container = GetContextMenuContainer(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!container) {
        return false;
    }
    const size_t index = data->contextSelectionPath.back();
    if (index == 0) {
        return false;
    }
    ContextMenuItem& previous = (*container)[index - 1];
    if (previous.type != ContextMenuItemType::kSubmenu) {
        return false;
    }
    ContextMenuItem item = std::move((*container)[index]);
    container->erase(container->begin() + static_cast<std::ptrdiff_t>(index));
    previous.children.push_back(std::move(item));
    data->contextSelectionPath.back() = index - 1;
    data->contextSelectionPath.push_back(previous.children.size() - 1);
    RefreshContextMenuTree(page, data, &data->contextSelectionPath);
    UpdateContextMenuButtonStates(page, data);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool OutdentContextMenuItem(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid || data->contextSelectionPath.size() < 2) {
        return false;
    }
    std::vector<size_t> path = data->contextSelectionPath;
    auto* container = GetContextMenuContainer(data->workingOptions.contextMenuItems, path);
    if (!container) {
        return false;
    }
    const size_t index = path.back();
    ContextMenuItem item = std::move((*container)[index]);
    container->erase(container->begin() + static_cast<std::ptrdiff_t>(index));

    path.pop_back();
    const size_t parentIndex = path.back();
    std::vector<size_t> destinationPath = path;
    destinationPath.pop_back();

    auto* destinationContainer = destinationPath.empty()
                                     ? &data->workingOptions.contextMenuItems
                                     : GetContextMenuContainer(data->workingOptions.contextMenuItems,
                                                               destinationPath);
    if (!destinationContainer) {
        return false;
    }
    const size_t insertIndex = parentIndex + 1;
    destinationContainer->insert(destinationContainer->begin() + static_cast<std::ptrdiff_t>(insertIndex),
                                 std::move(item));
    data->contextSelectionPath = destinationPath;
    data->contextSelectionPath.push_back(insertIndex);
    data->contextSelectionValid = true;
    RefreshContextMenuTree(page, data, &data->contextSelectionPath);
    UpdateContextMenuButtonStates(page, data);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool BrowseForCommandExecutable(HWND owner, std::wstring* path, std::wstring* directory) {
    if (!path) {
        return false;
    }
    wchar_t buffer[MAX_PATH] = {};
    if (!path->empty()) {
        wcsncpy_s(buffer, path->c_str(), _TRUNCATE);
    }
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFile = buffer;
    ofn.nMaxFile = ARRAYSIZE(buffer);
    ofn.lpstrFilter = L"Executable Files (*.exe;*.bat;*.cmd;*.com)\0*.exe;*.bat;*.cmd;*.com\0All Files (*.*)\0*.*\0\0";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR | OFN_HIDEREADONLY;
    std::wstring initialDir;
    if (directory && !directory->empty()) {
        initialDir = *directory;
    } else if (!path->empty()) {
        wchar_t temp[MAX_PATH] = {};
        wcsncpy_s(temp, path->c_str(), _TRUNCATE);
        PathRemoveFileSpecW(temp);
        initialDir = temp;
    }
    if (!initialDir.empty()) {
        ofn.lpstrInitialDir = initialDir.c_str();
    }
    if (!GetOpenFileNameW(&ofn)) {
        return false;
    }
    *path = buffer;
    if (directory) {
        wchar_t dirBuffer[MAX_PATH] = {};
        wcsncpy_s(dirBuffer, buffer, _TRUNCATE);
        PathRemoveFileSpecW(dirBuffer);
        *directory = dirBuffer;
    }
    return true;
}

bool HandleContextMenuBrowseIcon(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid) {
        return false;
    }
    ContextMenuItem* item =
        GetContextMenuItem(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!item || item->type == ContextMenuItemType::kSeparator) {
        return false;
    }
    std::wstring iconSource = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_ICON_EDIT)));
    std::vector<wchar_t> buffer(iconSource.begin(), iconSource.end());
    buffer.resize(std::max<size_t>(buffer.size(), MAX_PATH), L'\0');
    int iconIndex = 0;
    if (!buffer.empty()) {
        iconIndex = PathParseIconLocationW(buffer.data());
    }
    if (!PickIconDlg(page, buffer.data(), static_cast<UINT>(buffer.size()), &iconIndex)) {
        return false;
    }
    std::wstring result = buffer.data();
    if (iconIndex != 0) {
        result += L",";
        result += std::to_wstring(iconIndex);
    }
    SetWindowTextW(GetDlgItem(page, IDC_CONTEXT_ICON_EDIT), result.c_str());
    item->iconSource = result;
    UpdateContextMenuTreeItemText(page, data, *item);
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool HandleContextMenuBrowseCommand(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid) {
        return false;
    }
    ContextMenuItem* item =
        GetContextMenuItem(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!item || item->type != ContextMenuItemType::kCommand) {
        return false;
    }
    std::wstring executable = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_COMMAND_PATH)));
    if (!BrowseForCommandExecutable(page, &executable, &data->contextCommandBrowseDirectory)) {
        return false;
    }
    SetWindowTextW(GetDlgItem(page, IDC_CONTEXT_COMMAND_PATH), executable.c_str());
    ApplyContextMenuDetailsFromControls(page, data, true);
    return true;
}

bool HandleContextMenuExtensionAdd(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid) {
        return false;
    }
    ContextMenuItem* item =
        GetContextMenuItem(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!item) {
        return false;
    }
    std::wstring extension = Trim(GetWindowTextString(GetDlgItem(page, IDC_CONTEXT_EXTENSION_EDIT)));
    if (extension.empty()) {
        return false;
    }
    std::vector<std::wstring> normalized = NormalizeContextMenuExtensions({extension});
    if (normalized.empty()) {
        return false;
    }
    const std::wstring& normalizedExtension = normalized.front();
    if (std::find(item->scope.extensions.begin(), item->scope.extensions.end(), normalizedExtension) !=
        item->scope.extensions.end()) {
        return false;
    }
    item->scope.extensions.push_back(normalizedExtension);
    std::sort(item->scope.extensions.begin(), item->scope.extensions.end());
    item->scope.extensions.erase(std::unique(item->scope.extensions.begin(), item->scope.extensions.end()),
                                 item->scope.extensions.end());
    SetWindowTextW(GetDlgItem(page, IDC_CONTEXT_EXTENSION_EDIT), L"");
    RefreshContextMenuExtensionsList(page, *item);
    if (HWND list = GetDlgItem(page, IDC_CONTEXT_EXTENSION_LIST)) {
        if (!item->scope.extensions.empty()) {
            SendMessageW(list, LB_SETCURSEL, 0, 0);
        }
    }
    if (HWND removeButton = GetDlgItem(page, IDC_CONTEXT_EXTENSION_REMOVE)) {
        EnableWindow(removeButton, !item->scope.extensions.empty());
    }
    PropSheet_Changed(GetParent(page), page);
    return true;
}

bool HandleContextMenuExtensionRemove(HWND page, OptionsDialogData* data) {
    if (!data || !data->contextSelectionValid) {
        return false;
    }
    ContextMenuItem* item =
        GetContextMenuItem(data->workingOptions.contextMenuItems, data->contextSelectionPath);
    if (!item) {
        return false;
    }
    HWND list = GetDlgItem(page, IDC_CONTEXT_EXTENSION_LIST);
    if (!list) {
        return false;
    }
    const int selection = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
    if (selection < 0 || selection >= static_cast<int>(item->scope.extensions.size())) {
        return false;
    }
    item->scope.extensions.erase(item->scope.extensions.begin() + selection);
    RefreshContextMenuExtensionsList(page, *item);
    if (HWND removeButton = GetDlgItem(page, IDC_CONTEXT_EXTENSION_REMOVE)) {
        EnableWindow(removeButton, !item->scope.extensions.empty());
    }
    PropSheet_Changed(GetParent(page), page);
    return true;
}

struct ContextMenuValidationError {
    std::wstring message;
    std::vector<size_t> path;
};

bool CommandExecutableExists(const std::wstring& executablePath) {
    std::wstring trimmed = Trim(executablePath);
    if (trimmed.empty()) {
        return false;
    }
    if (!trimmed.empty() && trimmed.front() == L'"' && trimmed.back() == L'"') {
        trimmed = trimmed.substr(1, trimmed.size() - 2);
    }
    std::wstring expanded;
    expanded.resize(MAX_PATH);
    DWORD copied = ExpandEnvironmentStringsW(trimmed.c_str(), expanded.data(), static_cast<DWORD>(expanded.size()));
    if (copied > 0 && copied < expanded.size()) {
        expanded.resize(copied - 1);
    } else {
        expanded = trimmed;
    }
    if (PathFileExistsW(expanded.c_str())) {
        return true;
    }
    wchar_t buffer[MAX_PATH];
    if (SearchPathW(nullptr, expanded.c_str(), nullptr, ARRAYSIZE(buffer), buffer, nullptr)) {
        return true;
    }
    if (SearchPathW(nullptr, expanded.c_str(), L".exe", ARRAYSIZE(buffer), buffer, nullptr)) {
        return true;
    }
    return false;
}

bool ValidateContextMenuItems(const std::vector<ContextMenuItem>& items, std::vector<size_t>* path,
                              ContextMenuValidationError* error) {
    if (!path || !error) {
        return true;
    }
    for (size_t i = 0; i < items.size(); ++i) {
        path->push_back(i);
        const ContextMenuItem& item = items[i];
        if (item.type != ContextMenuItemType::kSeparator && Trim(item.label).empty()) {
            error->message = L"Context menu items must have a label.";
            error->path = *path;
            return false;
        }
        if (item.type == ContextMenuItemType::kCommand) {
            std::wstring commandPath;
            std::wstring commandArgs;
            SplitCommandTemplate(item.commandTemplate, &commandPath, &commandArgs);
            if (Trim(commandPath).empty()) {
                error->message = L"Command menu items must specify an executable.";
                error->path = *path;
                return false;
            }
            if (!CommandExecutableExists(commandPath)) {
                error->message = L"The specified command could not be located.";
                error->path = *path;
                return false;
            }
        }
        if (item.selection.maximumSelection > 0 &&
            item.selection.maximumSelection < item.selection.minimumSelection) {
            error->message = L"Maximum selection must be zero or greater than minimum selection.";
            error->path = *path;
            return false;
        }
        if (!ValidateContextMenuItems(item.children, path, error)) {
            return false;
        }
        path->pop_back();
    }
    return true;
}

NewTabTemplate GetSelectedNewTabTemplate(HWND hwnd, OptionsDialogData* data) {
    HWND combo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_COMBO);
    if (!combo) {
        return data ? data->workingOptions.newTabTemplate : NewTabTemplate::kDuplicateCurrent;
    }
    const LRESULT selection = SendMessageW(combo, CB_GETCURSEL, 0, 0);
    if (selection >= 0) {
        const LRESULT value = SendMessageW(combo, CB_GETITEMDATA, selection, 0);
        if (value != CB_ERR) {
            return static_cast<NewTabTemplate>(value);
        }
    }
    return data ? data->workingOptions.newTabTemplate : NewTabTemplate::kDuplicateCurrent;
}

void PopulateNewTabTemplateCombo(HWND hwnd, OptionsDialogData* data) {
    HWND combo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_COMBO);
    if (!combo) {
        return;
    }
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    struct TemplateEntry {
        NewTabTemplate value;
        const wchar_t* label;
    } entries[] = {
        {NewTabTemplate::kDuplicateCurrent, L"Duplicate current tab"},
        {NewTabTemplate::kThisPc, L"This PC"},
        {NewTabTemplate::kCustomPath, L"Custom path"},
        {NewTabTemplate::kSavedGroup, L"Saved group"},
    };

    int selectionIndex = -1;
    for (const auto& entry : entries) {
        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.label)));
        if (index >= 0) {
            SendMessageW(combo, CB_SETITEMDATA, index, static_cast<LPARAM>(entry.value));
            if (data && data->workingOptions.newTabTemplate == entry.value && selectionIndex < 0) {
                selectionIndex = index;
            }
        }
    }

    if (selectionIndex < 0) {
        selectionIndex = 0;
    }
    SendMessageW(combo, CB_SETCURSEL, selectionIndex, 0);
    if (data) {
        const LRESULT value = SendMessageW(combo, CB_GETITEMDATA, selectionIndex, 0);
        if (value != CB_ERR) {
            data->workingOptions.newTabTemplate = static_cast<NewTabTemplate>(value);
        }
    }
}

void PopulateNewTabGroupCombo(HWND hwnd, OptionsDialogData* data) {
    HWND combo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_GROUP_COMBO);
    if (!combo) {
        return;
    }

    const std::wstring previousSelection = data ? data->workingOptions.newTabSavedGroup : std::wstring();
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);

    if (!data || data->workingGroups.empty()) {
        const wchar_t placeholder[] = L"No saved groups available";
        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(placeholder)));
        if (index >= 0) {
            SendMessageW(combo, CB_SETCURSEL, index, 0);
        }
        EnableWindow(combo, FALSE);
        if (data) {
            data->workingOptions.newTabSavedGroup.clear();
        }
        return;
    }

    EnableWindow(combo, TRUE);
    int selectionIndex = -1;
    for (const auto& group : data->workingGroups) {
        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(group.name.c_str())));
        if (index >= 0) {
            if (!previousSelection.empty() && _wcsicmp(group.name.c_str(), previousSelection.c_str()) == 0) {
                selectionIndex = index;
            }
        }
    }

    if (selectionIndex >= 0) {
        SendMessageW(combo, CB_SETCURSEL, selectionIndex, 0);
    } else {
        SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(-1), 0);
        data->workingOptions.newTabSavedGroup.clear();
    }
}

void UpdateNewTabTemplateControls(HWND hwnd, OptionsDialogData* data) {
    const NewTabTemplate selected = GetSelectedNewTabTemplate(hwnd, data);
    if (data) {
        data->workingOptions.newTabTemplate = selected;
    }

    const bool showPath = (selected == NewTabTemplate::kCustomPath);
    const bool showGroup = (selected == NewTabTemplate::kSavedGroup);

    auto updateControl = [](HWND control, bool visible, bool enable) {
        if (!control) {
            return;
        }
        ShowWindow(control, visible ? SW_SHOWNOACTIVATE : SW_HIDE);
        EnableWindow(control, visible && enable);
    };

    HWND pathLabel = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_PATH_LABEL);
    HWND pathEdit = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT);
    HWND browseButton = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_BROWSE);
    updateControl(pathLabel, showPath, true);
    updateControl(pathEdit, showPath, true);
    updateControl(browseButton, showPath, true);

    HWND groupLabel = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_GROUP_LABEL);
    HWND groupCombo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_GROUP_COMBO);
    const bool hasGroups = data && !data->workingGroups.empty();
    if (groupLabel) {
        ShowWindow(groupLabel, showGroup ? SW_SHOWNOACTIVATE : SW_HIDE);
        EnableWindow(groupLabel, showGroup);
    }
    if (groupCombo) {
        ShowWindow(groupCombo, showGroup ? SW_SHOWNOACTIVATE : SW_HIDE);
        EnableWindow(groupCombo, showGroup && hasGroups);
    }
}

void AlignDialogBuffer(std::vector<BYTE>& buffer) {
    while (buffer.size() % 4 != 0) {
        buffer.push_back(0);
    }
}

void AppendWord(std::vector<BYTE>& buffer, WORD value) {
    buffer.push_back(static_cast<BYTE>(value & 0xFF));
    buffer.push_back(static_cast<BYTE>((value >> 8) & 0xFF));
}

void AppendString(std::vector<BYTE>& buffer, const wchar_t* text) {
    if (!text) {
        AppendWord(buffer, 0);
        return;
    }
    while (*text) {
        AppendWord(buffer, static_cast<WORD>(*text));
        ++text;
    }
    AppendWord(buffer, 0);
}

using DialogTemplatePtr = std::unique_ptr<DLGTEMPLATE, void (*)(void*)>;

DialogTemplatePtr AllocateAlignedTemplate(const std::vector<BYTE>& source) {
    if (source.empty()) {
        return DialogTemplatePtr(nullptr, &_aligned_free);
    }

    void* memory = _aligned_malloc(source.size(), alignof(DLGTEMPLATE));
    if (!memory) {
        return DialogTemplatePtr(nullptr, &_aligned_free);
    }

    std::memcpy(memory, source.data(), source.size());
    return DialogTemplatePtr(reinterpret_cast<DLGTEMPLATE*>(memory), &_aligned_free);
}

BOOL CALLBACK ForwardOptionsChangedToChild(HWND hwnd, LPARAM param) {
    const UINT message = static_cast<UINT>(param);
    if (message == 0 || !IsWindow(hwnd)) {
        return TRUE;
    }

    SendMessageTimeoutW(hwnd, message, 0, 0, SMTO_ABORTIFHUNG | SMTO_NOTIMEOUTIFNOTHUNG, 200, nullptr);
    return TRUE;
}

void ForceExplorerUIRefresh(HWND parentWindow) {
    const UINT optionsChangedMessage = GetOptionsChangedMessage();
    if (optionsChangedMessage == 0) {
        return;
    }

    SendMessageTimeoutW(HWND_BROADCAST, optionsChangedMessage, 0, 0,
                        SMTO_ABORTIFHUNG | SMTO_NOTIMEOUTIFNOTHUNG, 200, nullptr);

    if (!parentWindow || !IsWindow(parentWindow)) {
        return;
    }

    SendMessageTimeoutW(parentWindow, optionsChangedMessage, 0, 0,
                        SMTO_ABORTIFHUNG | SMTO_NOTIMEOUTIFNOTHUNG, 200, nullptr);
    EnumChildWindows(parentWindow, &ForwardOptionsChangedToChild,
                     static_cast<LPARAM>(optionsChangedMessage));
    RedrawWindow(parentWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_ERASE);
}

void ApplyCustomizationPreview(HWND pageWindow, OptionsDialogData* data) {
    if (!data) {
        return;
    }

    UpdateGlowPaletteFromLegacySettings(data->workingOptions);
    OptionsStore::Instance().Set(data->workingOptions);
    data->previewOptionsBroadcasted = true;
    ForceExplorerUIRefresh(GetParent(pageWindow));
}

std::vector<BYTE> BuildMainPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 13;
    dlg->x = 0;
    dlg->y = 0;
    dlg->cx = kMainDialogWidth;
    dlg->cy = kMainDialogHeight;

    AppendWord(data, 0);  // menu
    AppendWord(data, 0);  // class
    AppendWord(data, 0);  // title
    AppendWord(data, 9);  // font size
    AppendString(data, L"Segoe UI");

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* reopenCheck = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    reopenCheck->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    reopenCheck->dwExtendedStyle = 0;
    reopenCheck->x = 10;
    reopenCheck->y = 12;
    reopenCheck->cx = kMainCheckboxWidth;
    reopenCheck->cy = 12;
    reopenCheck->id = static_cast<WORD>(IDC_MAIN_REOPEN);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Always reopen last session after crash");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* persistCheck = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    persistCheck->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    persistCheck->dwExtendedStyle = 0;
    persistCheck->x = 10;
    persistCheck->y = 32;
    persistCheck->cx = kMainCheckboxWidth;
    persistCheck->cy = 12;
    persistCheck->id = static_cast<WORD>(IDC_MAIN_PERSIST);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Remember saved group paths on close");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* explorerAccentCheck = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    explorerAccentCheck->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    explorerAccentCheck->dwExtendedStyle = 0;
    explorerAccentCheck->x = 10;
    explorerAccentCheck->y = 52;
    explorerAccentCheck->cx = kMainCheckboxWidth;
    explorerAccentCheck->cy = 12;
    explorerAccentCheck->id = static_cast<WORD>(IDC_MAIN_LISTVIEW_ACCENT);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Apply tab group accents to Explorer list view");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* exampleStatic = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    exampleStatic->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    exampleStatic->dwExtendedStyle = 0;
    exampleStatic->x = 10;
    exampleStatic->y = 76;
    exampleStatic->cx = kMainDialogWidth - 20;
    exampleStatic->cy = 60;
    exampleStatic->id = static_cast<WORD>(IDC_MAIN_EXAMPLE);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    newTabLabel->dwExtendedStyle = 0;
    newTabLabel->x = 10;
    newTabLabel->y = 142;
    newTabLabel->cx = kMainDialogWidth - 20;
    newTabLabel->cy = 12;
    newTabLabel->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Default new tab content:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabCombo = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabCombo->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL;
    newTabCombo->dwExtendedStyle = WS_EX_CLIENTEDGE;
    newTabCombo->x = 10;
    newTabCombo->y = 156;
    newTabCombo->cx = kMainDialogWidth - 20;
    newTabCombo->cy = 70;
    newTabCombo->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_COMBO);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0085);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabPathLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabPathLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    newTabPathLabel->dwExtendedStyle = 0;
    newTabPathLabel->x = 10;
    newTabPathLabel->y = 178;
    newTabPathLabel->cx = kMainDialogWidth - 80;
    newTabPathLabel->cy = 12;
    newTabPathLabel->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_PATH_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Custom path:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabPathEdit = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabPathEdit->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
    newTabPathEdit->dwExtendedStyle = WS_EX_CLIENTEDGE;
    newTabPathEdit->x = 10;
    newTabPathEdit->y = 192;
    newTabPathEdit->cx = kMainDialogWidth - 100;
    newTabPathEdit->cy = 14;
    newTabPathEdit->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_PATH_EDIT);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0081);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabBrowse = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabBrowse->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    newTabBrowse->dwExtendedStyle = 0;
    newTabBrowse->x = static_cast<short>(kMainDialogWidth - 84);
    newTabBrowse->y = 191;
    newTabBrowse->cx = 74;
    newTabBrowse->cy = 16;
    newTabBrowse->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_BROWSE);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Browse...");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabGroupLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabGroupLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    newTabGroupLabel->dwExtendedStyle = 0;
    newTabGroupLabel->x = 10;
    newTabGroupLabel->y = 218;
    newTabGroupLabel->cx = kMainDialogWidth - 20;
    newTabGroupLabel->cy = 12;
    newTabGroupLabel->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_GROUP_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Saved group:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newTabGroupCombo = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newTabGroupCombo->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL;
    newTabGroupCombo->dwExtendedStyle = WS_EX_CLIENTEDGE;
    newTabGroupCombo->x = 10;
    newTabGroupCombo->y = 232;
    newTabGroupCombo->cx = kMainDialogWidth - 20;
    newTabGroupCombo->cy = 70;
    newTabGroupCombo->id = static_cast<WORD>(IDC_MAIN_NEW_TAB_GROUP_COMBO);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0085);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* dockLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    dockLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    dockLabel->dwExtendedStyle = 0;
    dockLabel->x = 10;
    dockLabel->y = 258;
    dockLabel->cx = kMainDialogWidth - 20;
    dockLabel->cy = 12;
    dockLabel->id = static_cast<WORD>(IDC_MAIN_DOCK_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Tab bar docking location:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* dockCombo = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    dockCombo->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL;
    dockCombo->dwExtendedStyle = WS_EX_CLIENTEDGE;
    dockCombo->x = 10;
    dockCombo->y = 272;
    dockCombo->cx = kMainDialogWidth - 20;
    dockCombo->cy = 70;
    dockCombo->id = static_cast<WORD>(IDC_MAIN_DOCK_COMBO);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0085);
    AppendString(data, L"");
    AppendWord(data, 0);

    return data;
}

std::vector<BYTE> BuildCustomizationPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style =
        DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VSCROLL;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 50;
    dlg->x = 0;
    dlg->y = 0;
    dlg->cx = kMainDialogWidth;
    dlg->cy = kMainDialogHeight;

    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendWord(data, 9);
    AppendString(data, L"Segoe UI");

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* breadcrumbGroup = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    breadcrumbGroup->style = WS_CHILD | WS_VISIBLE | BS_GROUPBOX;
    breadcrumbGroup->dwExtendedStyle = 0;
    breadcrumbGroup->x = 6;
    breadcrumbGroup->y = 6;
    breadcrumbGroup->cx = kMainDialogWidth - 12;
    breadcrumbGroup->cy = 310;
    breadcrumbGroup->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"BreadCrumb Bar");
    AppendWord(data, 0);

    auto addCheckbox = [&](int controlId, int x, int y, const wchar_t* text) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = kMainDialogWidth - 24;
        item->cy = 12;
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0080);
        AppendString(data, text);
        AppendWord(data, 0);
    };

    auto addStatic = [&](int controlId, int x, int y, int cx, int cy, const wchar_t* text,
                         DWORD style = SS_LEFT) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | style;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0082);
        AppendString(data, text);
        AppendWord(data, 0);
    };

    auto addPreview = [&](int controlId, int x, int y) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | SS_SUNKEN;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = 32;
        item->cy = 16;
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0082);
        AppendString(data, L"");
        AppendWord(data, 0);
    };

    auto addButton = [&](int controlId, int x, int y, const wchar_t* text) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = 50;
        item->cy = 16;
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0080);
        AppendString(data, text);
        AppendWord(data, 0);
    };

    auto addSizedButton = [&](int controlId, int x, int y, int cx, int cy, const wchar_t* text) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0080);
        AppendString(data, text);
        AppendWord(data, 0);
    };

    auto addSlider = [&](int controlId, int x, int y) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | TBS_AUTOTICKS;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = 170;
        item->cy = 16;
        item->id = static_cast<WORD>(controlId);
        AppendString(data, TRACKBAR_CLASSW);
        AppendWord(data, 0);
        AppendWord(data, 0);
    };

    addCheckbox(IDC_MAIN_BREADCRUMB, 16, 24, L"Enable breadcrumb background gradient");
    addCheckbox(IDC_MAIN_BREADCRUMB_FONT, 16, 44, L"Enable breadcrumb font color gradient");
    addStatic(IDC_MAIN_BREADCRUMB_BG_LABEL, 24, 64, kMainDialogWidth - 32, 10, L"Background transparency:");
    addSlider(IDC_MAIN_BREADCRUMB_BG_SLIDER, 24, 78);
    addStatic(IDC_MAIN_BREADCRUMB_BG_VALUE, 200, 80, 40, 12, L"", SS_RIGHT);
    addStatic(IDC_MAIN_BREADCRUMB_FONT_LABEL, 24, 102, kMainDialogWidth - 32, 10, L"Font brightness:");
    addSlider(IDC_MAIN_BREADCRUMB_FONT_SLIDER, 24, 116);
    addStatic(IDC_MAIN_BREADCRUMB_FONT_VALUE, 200, 118, 40, 12, L"", SS_RIGHT);
    addStatic(IDC_MAIN_BREADCRUMB_HIGHLIGHT_LABEL, 24, 140, kMainDialogWidth - 32, 10, L"Highlight intensity:");
    addSlider(IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER, 24, 154);
    addStatic(IDC_MAIN_BREADCRUMB_HIGHLIGHT_VALUE, 200, 156, 40, 12, L"", SS_RIGHT);
    addStatic(IDC_MAIN_BREADCRUMB_DROPDOWN_LABEL, 24, 178, kMainDialogWidth - 32, 10,
              L"Dropdown arrow intensity:");
    addSlider(IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER, 24, 192);
    addStatic(IDC_MAIN_BREADCRUMB_DROPDOWN_VALUE, 200, 194, 40, 12, L"", SS_RIGHT);
    addCheckbox(IDC_MAIN_BREADCRUMB_BG_CUSTOM, 16, 218, L"Use custom background gradient colors");
    addStatic(IDC_MAIN_BREADCRUMB_BG_START_LABEL, 24, 236, 60, 10, L"Start:");
    addPreview(IDC_MAIN_BREADCRUMB_BG_START_PREVIEW, 86, 234);
    addButton(IDC_MAIN_BREADCRUMB_BG_START_BUTTON, 124, 233, L"Choose");
    addStatic(IDC_MAIN_BREADCRUMB_BG_END_LABEL, 24, 256, 60, 10, L"End:");
    addPreview(IDC_MAIN_BREADCRUMB_BG_END_PREVIEW, 86, 254);
    addButton(IDC_MAIN_BREADCRUMB_BG_END_BUTTON, 124, 253, L"Choose");
    addCheckbox(IDC_MAIN_BREADCRUMB_FONT_CUSTOM, 16, 278, L"Use custom breadcrumb text colors");
    addStatic(IDC_MAIN_BREADCRUMB_FONT_START_LABEL, 24, 296, 60, 10, L"Start:");
    addPreview(IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW, 86, 294);
    addButton(IDC_MAIN_BREADCRUMB_FONT_START_BUTTON, 124, 293, L"Choose");
    addStatic(IDC_MAIN_BREADCRUMB_FONT_END_LABEL, 24, 316, 60, 10, L"End:");
    addPreview(IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW, 86, 314);
    addButton(IDC_MAIN_BREADCRUMB_FONT_END_BUTTON, 124, 313, L"Choose");
    addCheckbox(IDC_MAIN_PROGRESS_CUSTOM, 16, 338, L"Use custom progress bar gradient colors");
    addStatic(IDC_MAIN_PROGRESS_START_LABEL, 24, 356, 60, 10, L"Start:");
    addPreview(IDC_MAIN_PROGRESS_START_PREVIEW, 86, 354);
    addButton(IDC_MAIN_PROGRESS_START_BUTTON, 124, 353, L"Choose");
    addStatic(IDC_MAIN_PROGRESS_END_LABEL, 24, 376, 60, 10, L"End:");
    addPreview(IDC_MAIN_PROGRESS_END_PREVIEW, 86, 374);
    addButton(IDC_MAIN_PROGRESS_END_BUTTON, 124, 373, L"Choose");

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* tabsGroup = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    tabsGroup->style = WS_CHILD | WS_VISIBLE | BS_GROUPBOX;
    tabsGroup->dwExtendedStyle = 0;
    tabsGroup->x = 6;
    tabsGroup->y = 412;
    tabsGroup->cx = kMainDialogWidth - 12;
    tabsGroup->cy = 88;
    tabsGroup->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Tabs");
    AppendWord(data, 0);

    addCheckbox(IDC_MAIN_TAB_SELECTED_CHECK, 16, 428, L"Use custom selected tab color");
    addPreview(IDC_MAIN_TAB_SELECTED_PREVIEW, 24, 446);
    addButton(IDC_MAIN_TAB_SELECTED_BUTTON, 62, 445, L"Choose");
    addCheckbox(IDC_MAIN_TAB_UNSELECTED_CHECK, 16, 464, L"Use custom unselected tab color");
    addPreview(IDC_MAIN_TAB_UNSELECTED_PREVIEW, 24, 482);
    addButton(IDC_MAIN_TAB_UNSELECTED_BUTTON, 62, 481, L"Choose");

    auto addSizedPreview = [&](int controlId, int x, int y, int cx, int cy) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | SS_SUNKEN | SS_BITMAP | SS_CENTERIMAGE;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0082);
        AppendString(data, L"");
        AppendWord(data, 0);
    };

    auto addListView = [&](int controlId, int x, int y, int cx, int cy) {
        AlignDialogBuffer(data);
        size_t innerOffset = data.size();
        data.resize(innerOffset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + innerOffset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | LVS_REPORT | LVS_SINGLESEL |
                      LVS_SHOWSELALWAYS;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendString(data, WC_LISTVIEWW);
        AppendWord(data, 0);
        AppendWord(data, 0);
    };

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* backgroundsGroup = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    backgroundsGroup->style = WS_CHILD | WS_VISIBLE | BS_GROUPBOX;
    backgroundsGroup->dwExtendedStyle = 0;
    backgroundsGroup->x = 6;
    backgroundsGroup->y = 510;
    backgroundsGroup->cx = kMainDialogWidth - 12;
    backgroundsGroup->cy = 310;
    backgroundsGroup->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Folder Backgrounds");
    AppendWord(data, 0);

    addCheckbox(IDC_CUSTOM_BACKGROUND_ENABLE, 16, 526, L"Enable custom folder backgrounds");
    addStatic(0, 24, 546, kMainDialogWidth - 32, 10, L"Universal background image:");
    addSizedPreview(IDC_CUSTOM_BACKGROUND_PREVIEW, 24, 562, kUniversalPreviewSize.cx, kUniversalPreviewSize.cy);
    addSizedButton(IDC_CUSTOM_BACKGROUND_BROWSE, 130, 562, 90, 16, L"Browse...");
    addStatic(IDC_CUSTOM_BACKGROUND_UNIVERSAL_NAME, 130, 638, kMainDialogWidth - 146, 12, L"", SS_LEFT);
    addStatic(0, 24, 646, kMainDialogWidth - 32, 10, L"Folder overrides:");
    addListView(IDC_CUSTOM_BACKGROUND_LIST, 24, 660, 140, 96);
    addStatic(0, 176, 660, 64, 10, L"Preview:");
    addSizedPreview(IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW, 176, 674, kFolderPreviewSize.cx, kFolderPreviewSize.cy);
    addStatic(IDC_CUSTOM_BACKGROUND_FOLDER_NAME, 176, 742, kMainDialogWidth - 200, 12, L"", SS_LEFT);
    addSizedButton(IDC_CUSTOM_BACKGROUND_ADD, 24, 764, 60, 16, L"Add");
    addSizedButton(IDC_CUSTOM_BACKGROUND_EDIT, 92, 764, 60, 16, L"Edit");
    addSizedButton(IDC_CUSTOM_BACKGROUND_REMOVE, 160, 764, 60, 16, L"Remove");
    addSizedButton(IDC_CUSTOM_BACKGROUND_CLEAN, 228, 764, 90, 16, L"Clean Up...");

    AlignDialogBuffer(data);
    return data;
}

std::vector<BYTE> BuildGlowPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 16;
    dlg->x = 0;
    dlg->y = 0;
    dlg->cx = kGlowDialogWidth;
    dlg->cy = kGlowDialogHeight;

    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendWord(data, 9);
    AppendString(data, L"Segoe UI");

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* glowGroup = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    glowGroup->style = WS_CHILD | WS_VISIBLE | BS_GROUPBOX;
    glowGroup->dwExtendedStyle = 0;
    glowGroup->x = 6;
    glowGroup->y = 6;
    glowGroup->cx = kGlowDialogWidth - 12;
    glowGroup->cy = kGlowDialogHeight - 12;
    glowGroup->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Neon glow");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* enableGlow = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    enableGlow->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    enableGlow->dwExtendedStyle = 0;
    enableGlow->x = 16;
    enableGlow->y = 24;
    enableGlow->cx = kGlowCheckboxWidth;
    enableGlow->cy = 12;
    enableGlow->id = static_cast<WORD>(IDC_GLOW_ENABLE);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Enable neon glow effects");
    AppendWord(data, 0);

    auto appendSurfaceCheckbox = [&](int controlId, SHORT y, const wchar_t* label) {
        AlignDialogBuffer(data);
        size_t localOffset = data.size();
        data.resize(localOffset + sizeof(DLGITEMTEMPLATE));
        auto* checkbox = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + localOffset);
        checkbox->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
        checkbox->dwExtendedStyle = 0;
        checkbox->x = 16;
        checkbox->y = y;
        checkbox->cx = kGlowCheckboxWidth;
        checkbox->cy = 12;
        checkbox->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0080);
        AppendString(data, label);
        AppendWord(data, 0);
    };

    appendSurfaceCheckbox(IDC_GLOW_SURFACE_LISTVIEW, 44, L"Enable list view glow");
    appendSurfaceCheckbox(IDC_GLOW_SURFACE_HEADER, 60, L"Enable column header glow");
    appendSurfaceCheckbox(IDC_GLOW_SURFACE_REBAR, 76, L"Enable rebar glow");
    appendSurfaceCheckbox(IDC_GLOW_SURFACE_TOOLBAR, 92, L"Enable toolbar glow");
    appendSurfaceCheckbox(IDC_GLOW_SURFACE_EDIT, 108, L"Enable address bar glow");
    appendSurfaceCheckbox(IDC_GLOW_SURFACE_DIRECTUI, 124, L"Enable DirectUI glow");

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* customColors = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    customColors->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    customColors->dwExtendedStyle = 0;
    customColors->x = 16;
    customColors->y = 144;
    customColors->cx = kGlowCheckboxWidth;
    customColors->cy = 12;
    customColors->id = static_cast<WORD>(IDC_GLOW_CUSTOM_COLORS);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Use custom glow colors");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* useGradient = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    useGradient->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX;
    useGradient->dwExtendedStyle = 0;
    useGradient->x = 16;
    useGradient->y = 164;
    useGradient->cx = kGlowCheckboxWidth;
    useGradient->cy = 12;
    useGradient->id = static_cast<WORD>(IDC_GLOW_USE_GRADIENT);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Blend glow with gradient");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* primaryLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    primaryLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    primaryLabel->dwExtendedStyle = 0;
    primaryLabel->x = 16;
    primaryLabel->y = 192;
    primaryLabel->cx = 68;
    primaryLabel->cy = 12;
    primaryLabel->id = static_cast<WORD>(IDC_GLOW_PRIMARY_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Primary color:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* primaryPreview = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    primaryPreview->style = WS_CHILD | WS_VISIBLE | SS_SUNKEN;
    primaryPreview->dwExtendedStyle = 0;
    primaryPreview->x = 86;
    primaryPreview->y = 190;
    primaryPreview->cx = 40;
    primaryPreview->cy = 16;
    primaryPreview->id = static_cast<WORD>(IDC_GLOW_PRIMARY_PREVIEW);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* primaryButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    primaryButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    primaryButton->dwExtendedStyle = 0;
    primaryButton->x = 134;
    primaryButton->y = 188;
    primaryButton->cx = 72;
    primaryButton->cy = 14;
    primaryButton->id = static_cast<WORD>(IDC_GLOW_PRIMARY_BUTTON);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Choose...");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* secondaryLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    secondaryLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    secondaryLabel->dwExtendedStyle = 0;
    secondaryLabel->x = 16;
    secondaryLabel->y = 220;
    secondaryLabel->cx = 68;
    secondaryLabel->cy = 12;
    secondaryLabel->id = static_cast<WORD>(IDC_GLOW_SECONDARY_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Secondary color:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* secondaryPreview = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    secondaryPreview->style = WS_CHILD | WS_VISIBLE | SS_SUNKEN;
    secondaryPreview->dwExtendedStyle = 0;
    secondaryPreview->x = 86;
    secondaryPreview->y = 218;
    secondaryPreview->cx = 40;
    secondaryPreview->cy = 16;
    secondaryPreview->id = static_cast<WORD>(IDC_GLOW_SECONDARY_PREVIEW);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* secondaryButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    secondaryButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    secondaryButton->dwExtendedStyle = 0;
    secondaryButton->x = 134;
    secondaryButton->y = 216;
    secondaryButton->cx = 72;
    secondaryButton->cy = 14;
    secondaryButton->id = static_cast<WORD>(IDC_GLOW_SECONDARY_BUTTON);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Choose...");
    AppendWord(data, 0);

    return data;
}

std::vector<BYTE> BuildContextMenuPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 36;
    dlg->x = 0;
    dlg->y = 0;
    dlg->cx = kContextDialogWidth;
    dlg->cy = kContextDialogHeight;

    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendWord(data, 9);
    AppendString(data, L"Segoe UI");

    auto addButton = [&](int controlId, int x, int y, int cx, int cy, const wchar_t* text, DWORD style) {
        AlignDialogBuffer(data);
        size_t offset = data.size();
        data.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
        item->style = style;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0080);
        AppendString(data, text);
        AppendWord(data, 0);
    };

    auto addStatic = [&](int controlId, int x, int y, int cx, int cy, const wchar_t* text, DWORD style) {
        AlignDialogBuffer(data);
        size_t offset = data.size();
        data.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
        item->style = style;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0082);
        AppendString(data, text);
        AppendWord(data, 0);
    };

    auto addEdit = [&](int controlId, int x, int y, int cx, int cy, DWORD style) {
        AlignDialogBuffer(data);
        size_t offset = data.size();
        data.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
        item->style = style;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0081);
        AppendString(data, L"");
        AppendWord(data, 0);
    };

    auto addCombo = [&](int controlId, int x, int y, int cx, int cy) {
        AlignDialogBuffer(data);
        size_t offset = data.size();
        data.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(cx);
        item->cy = static_cast<short>(cy);
        item->id = static_cast<WORD>(controlId);
        AppendWord(data, 0xFFFF);
        AppendWord(data, 0x0085);
        AppendString(data, L"");
        AppendWord(data, 0);
    };

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* tree = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    tree->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | TVS_HASBUTTONS | TVS_LINESATROOT |
                  TVS_SHOWSELALWAYS;
    tree->dwExtendedStyle = WS_EX_CLIENTEDGE;
    tree->x = 8;
    tree->y = 8;
    tree->cx = 150;
    tree->cy = 220;
    tree->id = static_cast<WORD>(IDC_CONTEXT_TREE);
    AppendString(data, WC_TREEVIEW);
    AppendWord(data, 0);

    addButton(IDC_CONTEXT_ADD_COMMAND, 8, 236, 70, 16, L"Add command",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_ADD_SUBMENU, 8, 256, 70, 16, L"Add submenu",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_ADD_SEPARATOR, 8, 276, 70, 16, L"Add separator",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_REMOVE, 8, 296, 70, 16, L"Remove",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_MOVE_UP, 88, 236, 70, 16, L"Move up",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_MOVE_DOWN, 88, 256, 70, 16, L"Move down",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_INDENT, 88, 276, 70, 16, L"Indent",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    addButton(IDC_CONTEXT_OUTDENT, 88, 296, 70, 16, L"Outdent",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);

    const int detailX = 180;
    const int detailWidth = kContextDialogWidth - detailX - 10;

    addStatic(0, detailX, 8, detailWidth, 10, L"Display name:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_LABEL_EDIT, detailX, 20, detailWidth, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL);

    addStatic(0, detailX, 44, detailWidth, 10, L"Icon:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_ICON_EDIT, detailX, 56, detailWidth - 60, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL);
    addButton(IDC_CONTEXT_ICON_BROWSE, detailX + detailWidth - 58, 56, 58, 14, L"Browse...",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);

    addStatic(0, detailX, 84, detailWidth, 10, L"Command path:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_COMMAND_PATH, detailX, 96, detailWidth - 60, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL);
    addButton(IDC_CONTEXT_COMMAND_BROWSE, detailX + detailWidth - 58, 96, 58, 14, L"Browse...",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);

    addStatic(0, detailX, 124, detailWidth, 10, L"Command arguments:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_COMMAND_ARGS, detailX, 136, detailWidth, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL);

    addStatic(IDC_CONTEXT_HINTS_STATIC, detailX, 160, detailWidth, 20, L"",
              WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX);

    addStatic(0, detailX, 188, detailWidth, 10, L"Selection count:", WS_CHILD | WS_VISIBLE);
    addStatic(0, detailX, 202, 30, 10, L"Min:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_SELECTION_MIN, detailX + 32, 200, 40, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_NUMBER | ES_AUTOHSCROLL);
    addStatic(0, detailX + 78, 202, 30, 10, L"Max:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_SELECTION_MAX, detailX + 110, 200, 40, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_NUMBER | ES_AUTOHSCROLL);

    addStatic(0, detailX, 224, detailWidth, 10, L"Insertion anchor:", WS_CHILD | WS_VISIBLE);
    addCombo(IDC_CONTEXT_ANCHOR_COMBO, detailX, 236, detailWidth, 70);

    addStatic(0, detailX, 264, detailWidth, 10, L"Scope:", WS_CHILD | WS_VISIBLE);
    addButton(IDC_CONTEXT_SCOPE_FILES, detailX, 278, detailWidth, 14, L"Apply to all files",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);
    addButton(IDC_CONTEXT_SCOPE_FOLDERS, detailX, 296, detailWidth, 14, L"Apply to all folders",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);
    addButton(IDC_CONTEXT_SEPARATOR_CHECK, detailX, 314, detailWidth, 14, L"Group with separator above",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);

    addStatic(0, detailX, 334, detailWidth, 10, L"Extensions:", WS_CHILD | WS_VISIBLE);
    addEdit(IDC_CONTEXT_EXTENSION_EDIT, detailX, 346, detailWidth - 60, 14,
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL);
    addButton(IDC_CONTEXT_EXTENSION_ADD, detailX + detailWidth - 58, 346, 58, 14, L"Add",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* list = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    list->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | LBS_NOTIFY | LBS_HASSTRINGS |
                  LBS_NOINTEGRALHEIGHT | WS_VSCROLL;
    list->dwExtendedStyle = WS_EX_CLIENTEDGE;
    list->x = static_cast<short>(detailX);
    list->y = 366;
    list->cx = static_cast<short>(detailWidth - 60);
    list->cy = 60;
    list->id = static_cast<WORD>(IDC_CONTEXT_EXTENSION_LIST);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0083);
    AppendString(data, L"");
    AppendWord(data, 0);

    addButton(IDC_CONTEXT_EXTENSION_REMOVE, detailX + detailWidth - 58, 366, 58, 14, L"Remove",
              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);

    return data;
}

std::vector<BYTE> BuildGroupPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 4;
    dlg->x = 0;
    dlg->y = 0;
    dlg->cx = kGroupDialogWidth;
    dlg->cy = kGroupDialogHeight;

    AppendWord(data, 0);  // menu
    AppendWord(data, 0);  // class
    AppendWord(data, 0);  // title
    AppendWord(data, 9);
    AppendString(data, L"Segoe UI");

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* list = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    list->style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LBS_NOTIFY |
                  LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_HSCROLL;
    list->dwExtendedStyle = WS_EX_CLIENTEDGE;
    list->x = 10;
    list->y = 12;
    list->cx = 200;
    list->cy = 140;
    list->id = static_cast<WORD>(IDC_GROUP_LIST);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0083);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* newButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    newButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    newButton->dwExtendedStyle = 0;
    newButton->x = 220;
    newButton->y = 12;
    newButton->cx = 80;
    newButton->cy = 14;
    newButton->id = static_cast<WORD>(IDC_GROUP_NEW);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"New Group");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* editButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    editButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    editButton->dwExtendedStyle = 0;
    editButton->x = 220;
    editButton->y = 32;
    editButton->cx = 80;
    editButton->cy = 14;
    editButton->id = static_cast<WORD>(IDC_GROUP_EDIT);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Edit Group");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* removeButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    removeButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    removeButton->dwExtendedStyle = 0;
    removeButton->x = 220;
    removeButton->y = 52;
    removeButton->cx = 80;
    removeButton->cy = 14;
    removeButton->id = static_cast<WORD>(IDC_GROUP_REMOVE);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Remove");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    return data;
}

std::vector<BYTE> BuildGroupEditorTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_MODALFRAME | WS_POPUP | WS_CAPTION | WS_SYSMENU;
    dlg->dwExtendedStyle = 0;
    dlg->cdit = 14;
    dlg->x = 0;
    dlg->y = 0;
    dlg->cx = kEditorWidth;
    dlg->cy = kEditorHeight;

    AppendWord(data, 0);
    AppendWord(data, 0);
    AppendString(data, L"Edit Group");
    AppendWord(data, 9);
    AppendString(data, L"Segoe UI");

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* nameLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    nameLabel->style = WS_CHILD | WS_VISIBLE;
    nameLabel->dwExtendedStyle = 0;
    nameLabel->x = 10;
    nameLabel->y = 10;
    nameLabel->cx = 60;
    nameLabel->cy = 10;
    nameLabel->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Group name:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* nameEdit = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    nameEdit->style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL;
    nameEdit->dwExtendedStyle = WS_EX_CLIENTEDGE;
    nameEdit->x = 10;
    nameEdit->y = 22;
    nameEdit->cx = 200;
    nameEdit->cy = 14;
    nameEdit->id = static_cast<WORD>(IDC_EDITOR_NAME);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0081);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* colorLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    colorLabel->style = WS_CHILD | WS_VISIBLE;
    colorLabel->dwExtendedStyle = 0;
    colorLabel->x = 10;
    colorLabel->y = 42;
    colorLabel->cx = 40;
    colorLabel->cy = 10;
    colorLabel->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Color:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* colorPreview = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    colorPreview->style = WS_CHILD | WS_VISIBLE | SS_SUNKEN;
    colorPreview->dwExtendedStyle = 0;
    colorPreview->x = 55;
    colorPreview->y = 40;
    colorPreview->cx = 40;
    colorPreview->cy = 16;
    colorPreview->id = static_cast<WORD>(IDC_EDITOR_COLOR_PREVIEW);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* colorButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    colorButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    colorButton->dwExtendedStyle = 0;
    colorButton->x = 102;
    colorButton->y = 40;
    colorButton->cx = 80;
    colorButton->cy = 14;
    colorButton->id = static_cast<WORD>(IDC_EDITOR_COLOR_BUTTON);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Choose Color...");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* styleLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    styleLabel->style = WS_CHILD | WS_VISIBLE;
    styleLabel->dwExtendedStyle = 0;
    styleLabel->x = 10;
    styleLabel->y = 62;
    styleLabel->cx = 60;
    styleLabel->cy = 10;
    styleLabel->id = static_cast<WORD>(IDC_EDITOR_STYLE_LABEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Style:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* styleCombo = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    styleCombo->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL;
    styleCombo->dwExtendedStyle = WS_EX_CLIENTEDGE;
    styleCombo->x = 55;
    styleCombo->y = 74;
    styleCombo->cx = 127;
    styleCombo->cy = 110;
    styleCombo->id = static_cast<WORD>(IDC_EDITOR_STYLE_COMBO);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0085);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* pathsLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    pathsLabel->style = WS_CHILD | WS_VISIBLE;
    pathsLabel->dwExtendedStyle = 0;
    pathsLabel->x = 10;
    pathsLabel->y = 96;
    pathsLabel->cx = 60;
    pathsLabel->cy = 10;
    pathsLabel->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Paths:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* pathList = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    pathList->style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LBS_NOTIFY |
                      LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_HSCROLL;
    pathList->dwExtendedStyle = WS_EX_CLIENTEDGE;
    pathList->x = 10;
    pathList->y = 108;
    pathList->cx = 220;
    pathList->cy = 96;
    pathList->id = static_cast<WORD>(IDC_EDITOR_PATH_LIST);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0083);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* addButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    addButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    addButton->dwExtendedStyle = 0;
    addButton->x = 240;
    addButton->y = 108;
    addButton->cx = 80;
    addButton->cy = 14;
    addButton->id = static_cast<WORD>(IDC_EDITOR_ADD_PATH);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Add Path...");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* editButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    editButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    editButton->dwExtendedStyle = 0;
    editButton->x = 240;
    editButton->y = 128;
    editButton->cx = 80;
    editButton->cy = 14;
    editButton->id = static_cast<WORD>(IDC_EDITOR_EDIT_PATH);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Edit Path...");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* removeButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    removeButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    removeButton->dwExtendedStyle = 0;
    removeButton->x = 240;
    removeButton->y = 148;
    removeButton->cx = 80;
    removeButton->cy = 14;
    removeButton->id = static_cast<WORD>(IDC_EDITOR_REMOVE_PATH);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Remove");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* okButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    okButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON;
    okButton->dwExtendedStyle = 0;
    okButton->x = kEditorWidth - 120;
    okButton->y = kEditorHeight - 28;
    okButton->cx = 50;
    okButton->cy = 14;
    okButton->id = static_cast<WORD>(IDOK);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Save");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* cancelButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    cancelButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    cancelButton->dwExtendedStyle = 0;
    cancelButton->x = kEditorWidth - 64;
    cancelButton->y = kEditorHeight - 28;
    cancelButton->cx = 50;
    cancelButton->cy = 14;
    cancelButton->id = static_cast<WORD>(IDCANCEL);
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Cancel");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    return data;
}

bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
    return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

std::wstring ExtractDirectoryFromPath(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }
    size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring::npos) {
        return {};
    }
    return path.substr(0, separator);
}

bool CopyImageToCache(const std::wstring& sourcePath,
                      const std::wstring& displayName,
                      CachedImageMetadata* metadata,
                      std::wstring* createdPath,
                      std::wstring* errorMessage) {
    return CopyImageToBackgroundCache(sourcePath, displayName, metadata, createdPath, errorMessage);
}

bool BrowseForImage(HWND parent, std::wstring* path, std::wstring* displayName, const std::wstring& initialDirectory) {
    if (!path) {
        return false;
    }

    Microsoft::WRL::ComPtr<IFileDialog> dialog;
    if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog))) && dialog) {
        DWORD options = 0;
        if (SUCCEEDED(dialog->GetOptions(&options))) {
            dialog->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST);
        }
        COMDLG_FILTERSPEC filters[] = {
            {L"Image Files", L"*.png;*.jpg;*.jpeg;*.jfif;*.bmp;*.dib;*.gif"},
            {L"All Files", L"*.*"},
        };
        dialog->SetFileTypes(static_cast<UINT>(ARRAYSIZE(filters)), filters);
        dialog->SetFileTypeIndex(1);

        std::wstring initial = !initialDirectory.empty() ? initialDirectory : ExtractDirectoryFromPath(*path);
        if (!initial.empty()) {
            Microsoft::WRL::ComPtr<IShellItem> folder;
            if (SUCCEEDED(SHCreateItemFromParsingName(initial.c_str(), nullptr, IID_PPV_ARGS(&folder))) && folder) {
                dialog->SetFolder(folder.Get());
            }
        }

        if (SUCCEEDED(dialog->Show(parent))) {
            Microsoft::WRL::ComPtr<IShellItem> result;
            if (SUCCEEDED(dialog->GetResult(&result)) && result) {
                std::wstring filePath;
                if (TryGetFileSystemPath(result.Get(), &filePath)) {
                    *path = NormalizeFileSystemPath(filePath);
                    if (displayName) {
                        PWSTR buffer = nullptr;
                        if (SUCCEEDED(result->GetDisplayName(SIGDN_NORMALDISPLAY, &buffer)) && buffer) {
                            *displayName = buffer;
                            CoTaskMemFree(buffer);
                        } else {
                            const wchar_t* nameOnly = PathFindFileNameW(path->c_str());
                            *displayName = nameOnly ? nameOnly : *path;
                        }
                    }
                    return true;
                }
            }
            return false;
        }
    }

    wchar_t buffer[MAX_PATH] = {};
    if (!path->empty() && path->size() < ARRAYSIZE(buffer)) {
        wcsncpy_s(buffer, ARRAYSIZE(buffer), path->c_str(), _TRUNCATE);
    }
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = parent;
    ofn.lpstrFilter =
        L"Image Files (*.png;*.jpg;*.jpeg;*.jfif;*.bmp;*.dib;*.gif)\0*.png;*.jpg;*.jpeg;*.jfif;*.bmp;*.dib;*.gif\0"
        L"All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFile = buffer;
    ofn.nMaxFile = ARRAYSIZE(buffer);
    std::wstring initial = !initialDirectory.empty() ? initialDirectory : ExtractDirectoryFromPath(*path);
    ofn.lpstrInitialDir = initial.empty() ? nullptr : initial.c_str();
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
    if (GetOpenFileNameW(&ofn)) {
        *path = NormalizeFileSystemPath(buffer);
        if (displayName) {
            const wchar_t* nameOnly = PathFindFileNameW(buffer);
            *displayName = nameOnly ? nameOnly : *path;
        }
        return true;
    }
    return false;
}

void TrackCreatedCachedImage(OptionsDialogData* data, const std::wstring& path) {
    if (!data || path.empty()) {
        return;
    }
    auto it = std::find_if(data->createdCachedImagePaths.begin(), data->createdCachedImagePaths.end(),
                           [&](const std::wstring& existing) { return EqualsInsensitive(existing, path); });
    if (it == data->createdCachedImagePaths.end()) {
        data->createdCachedImagePaths.push_back(path);
    }
}

bool IsCachedImageInUse(const OptionsDialogData* data, const std::wstring& path) {
    if (!data || path.empty()) {
        return false;
    }
    if (EqualsInsensitive(data->workingOptions.universalFolderBackgroundImage.cachedImagePath, path)) {
        return true;
    }
    for (const auto& entry : data->workingOptions.folderBackgroundEntries) {
        if (EqualsInsensitive(entry.image.cachedImagePath, path)) {
            return true;
        }
    }
    return false;
}

void ScheduleCachedImageRemoval(OptionsDialogData* data, const std::wstring& path) {
    if (!data || path.empty()) {
        return;
    }
    if (IsCachedImageInUse(data, path)) {
        return;
    }
    auto createdIt = std::find_if(data->createdCachedImagePaths.begin(), data->createdCachedImagePaths.end(),
                                  [&](const std::wstring& created) { return EqualsInsensitive(created, path); });
    if (createdIt != data->createdCachedImagePaths.end()) {
        DeleteFileW(createdIt->c_str());
        data->createdCachedImagePaths.erase(createdIt);
        return;
    }
    auto pendingIt = std::find_if(data->pendingCachedImageRemovals.begin(), data->pendingCachedImageRemovals.end(),
                                  [&](const std::wstring& pending) { return EqualsInsensitive(pending, path); });
    if (pendingIt == data->pendingCachedImageRemovals.end()) {
        data->pendingCachedImageRemovals.push_back(path);
    }
}

void InitializeFolderBackgroundList(HWND list) {
    if (!list) {
        return;
    }
    ListView_SetExtendedListViewStyleEx(list, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER,
                                        LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
    LVCOLUMNW column{};
    column.mask = LVCF_TEXT | LVCF_WIDTH;
    column.pszText = const_cast<wchar_t*>(L"Folder");
    column.cx = 140;
    ListView_InsertColumn(list, 0, &column);
    column.pszText = const_cast<wchar_t*>(L"Image");
    column.cx = 100;
    ListView_InsertColumn(list, 1, &column);
}

void RefreshFolderBackgroundListView(HWND list, const OptionsDialogData* data) {
    if (!list) {
        return;
    }
    ListView_DeleteAllItems(list);
    if (!data) {
        return;
    }
    for (size_t i = 0; i < data->workingOptions.folderBackgroundEntries.size(); ++i) {
        const auto& entry = data->workingOptions.folderBackgroundEntries[i];
        LVITEMW item{};
        item.mask = LVIF_TEXT | LVIF_PARAM;
        item.iItem = static_cast<int>(i);
        item.pszText = const_cast<wchar_t*>(entry.folderPath.c_str());
        item.lParam = static_cast<LPARAM>(i);
        const int index = ListView_InsertItem(list, &item);
        if (index >= 0) {
            ListView_SetItemText(list, index, 1, const_cast<wchar_t*>(entry.image.displayName.c_str()));
        }
    }
}

int GetSelectedFolderBackgroundIndex(HWND list) {
    if (!list) {
        return -1;
    }
    return ListView_GetNextItem(list, -1, LVNI_SELECTED);
}

void UpdateFolderBackgroundControlsEnabled(HWND hwnd, bool enabled) {
    const int controls[] = {IDC_CUSTOM_BACKGROUND_BROWSE,        IDC_CUSTOM_BACKGROUND_PREVIEW,
                            IDC_CUSTOM_BACKGROUND_UNIVERSAL_NAME, IDC_CUSTOM_BACKGROUND_LIST,
                            IDC_CUSTOM_BACKGROUND_ADD,           IDC_CUSTOM_BACKGROUND_EDIT,
                            IDC_CUSTOM_BACKGROUND_REMOVE,        IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW,
                            IDC_CUSTOM_BACKGROUND_FOLDER_NAME,   IDC_CUSTOM_BACKGROUND_CLEAN};
    for (int id : controls) {
        if (HWND control = GetDlgItem(hwnd, id)) {
            EnableWindow(control, enabled);
        }
    }
}

void UpdateFolderBackgroundButtons(HWND hwnd) {
    const bool enabled = IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED;
    HWND list = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
    const int selection = GetSelectedFolderBackgroundIndex(list);
    const bool hasSelection = enabled && selection >= 0;
    if (HWND addButton = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_ADD)) {
        EnableWindow(addButton, enabled);
    }
    if (HWND editButton = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_EDIT)) {
        EnableWindow(editButton, hasSelection);
    }
    if (HWND removeButton = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_REMOVE)) {
        EnableWindow(removeButton, hasSelection);
    }
    if (HWND cleanButton = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_CLEAN)) {
        EnableWindow(cleanButton, enabled);
    }
}

std::wstring FormatCacheMaintenanceSummary(const CacheMaintenanceResult& result) {
    if (result.removedPaths.empty() && result.failures.empty()) {
        return L"No orphaned cache entries were found.";
    }

    std::wstring message;
    if (!result.removedPaths.empty()) {
        message += L"Removed " + std::to_wstring(result.removedPaths.size()) + L" orphaned cached image";
        message += result.removedPaths.size() == 1 ? L"." : L"s.";
        const size_t listCount = std::min<size_t>(result.removedPaths.size(), 5);
        for (size_t i = 0; i < listCount; ++i) {
            message += L"\n  - " + result.removedPaths[i];
        }
        if (result.removedPaths.size() > listCount) {
            message += L"\n  - ...";
        }
    }

    if (!result.failures.empty()) {
        if (!message.empty()) {
            message += L"\n\n";
        }
        message += L"Unable to remove " + std::to_wstring(result.failures.size()) + L" cache item";
        message += result.failures.size() == 1 ? L":" : L"s:";
        for (const auto& failure : result.failures) {
            message += L"\n  - " + failure.path;
            if (!failure.message.empty()) {
                message += L" (" + failure.message + L")";
            }
        }
    }

    return message;
}

void HandleBackgroundCacheMaintenance(HWND hwnd, OptionsDialogData* data) {
    auto& store = OptionsStore::Instance();
    std::wstring optionsLoadError;
    if (!store.Load(&optionsLoadError)) {
        if (!optionsLoadError.empty()) {
            LogMessage(LogLevel::Warning, L"HandleBackgroundCacheMaintenance failed to load options: %ls",
                       optionsLoadError.c_str());
        } else {
            LogMessage(LogLevel::Warning, L"HandleBackgroundCacheMaintenance failed to load options");
        }
    }
    ShellTabsOptions persisted = store.Get();

    std::vector<std::wstring> protectedPaths;
    if (data) {
        std::vector<std::wstring> workingReferences = CollectCachedImageReferences(data->workingOptions);
        protectedPaths.insert(protectedPaths.end(), workingReferences.begin(), workingReferences.end());
        for (const auto& created : data->createdCachedImagePaths) {
            if (!created.empty()) {
                protectedPaths.push_back(created);
            }
        }
    }

    CacheMaintenanceResult maintenance = RemoveOrphanedCacheEntries(persisted, protectedPaths);
    std::wstring summary = FormatCacheMaintenanceSummary(maintenance);
    if (summary.empty()) {
        summary = L"Cache maintenance completed.";
    }

    const UINT icon = maintenance.failures.empty() ? MB_ICONINFORMATION : MB_ICONWARNING;
    MessageBoxW(hwnd, summary.c_str(), L"ShellTabs", MB_OK | icon);
}

HBITMAP LoadPreviewBitmapSync(const std::wstring& path, const SIZE& size) {
    if (path.empty()) {
        return nullptr;
    }
    Microsoft::WRL::ComPtr<IShellItem> item;
    if (FAILED(SHCreateItemFromParsingName(path.c_str(), nullptr, IID_PPV_ARGS(&item))) || !item) {
        return nullptr;
    }
    Microsoft::WRL::ComPtr<IShellItemImageFactory> factory;
    if (FAILED(item.As(&factory)) || !factory) {
        return nullptr;
    }
    SIZE desired = size;
    HBITMAP bitmap = nullptr;
    HRESULT hr = factory->GetImage(desired, SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK | SIIGBF_THUMBNAILONLY, &bitmap);
    if (FAILED(hr)) {
        hr = factory->GetImage(desired, SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK, &bitmap);
    }
    if (FAILED(hr)) {
        hr = factory->GetImage(desired, SIIGBF_ICONONLY, &bitmap);
    }
    if (FAILED(hr)) {
        return nullptr;
    }
    TouchCachedImage(path);
    return bitmap;
}

HBITMAP CreatePlaceholderBitmap(const SIZE& size) {
    if (size.cx <= 0 || size.cy <= 0) {
        return nullptr;
    }

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = size.cx;
    info.bmiHeader.biHeight = -size.cy;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap || !bits) {
        if (bitmap && !bits) {
            DeleteObject(bitmap);
            bitmap = nullptr;
        }
        return bitmap;
    }

    DWORD* pixels = static_cast<DWORD*>(bits);
    for (int y = 0; y < size.cy; ++y) {
        for (int x = 0; x < size.cx; ++x) {
            const bool dark = (((x / 4) + (y / 4)) & 1) == 0;
            const BYTE value = dark ? 0xC0 : 0xE0;
            pixels[y * size.cx + x] = 0xFF000000 | (static_cast<DWORD>(value) << 16) |
                                      (static_cast<DWORD>(value) << 8) | static_cast<DWORD>(value);
        }
    }
    return bitmap;
}

void SetPreviewBitmap(HWND hwnd, int controlId, HBITMAP* stored, HBITMAP bitmap) {
    if (!stored) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return;
    }
    HBITMAP previousStored = *stored;
    HBITMAP oldControl = reinterpret_cast<HBITMAP>(
        SendDlgItemMessageW(hwnd, controlId, STM_SETIMAGE, IMAGE_BITMAP, reinterpret_cast<LPARAM>(bitmap)));
    if (oldControl && oldControl != bitmap && oldControl != previousStored) {
        DeleteObject(oldControl);
    }
    if (previousStored && previousStored != bitmap) {
        DeleteObject(previousStored);
    }
    *stored = bitmap;
}

void RequestPreviewBitmap(HWND hwnd,
                          int controlId,
                          const std::wstring& path,
                          const SIZE& size,
                          UINT64* tokenStorage,
                          HBITMAP* storedBitmap) {
    if (!tokenStorage || !storedBitmap) {
        return;
    }

    const UINT64 token = ++(*tokenStorage);

    if (path.empty()) {
        SetPreviewBitmap(hwnd, controlId, storedBitmap, nullptr);
        return;
    }

    HBITMAP placeholder = CreatePlaceholderBitmap(size);
    SetPreviewBitmap(hwnd, controlId, storedBitmap, placeholder);

    std::wstring pathCopy = path;
    std::thread([hwnd, controlId, size, token, pathCopy]() {
        HRESULT initHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(initHr) && initHr != RPC_E_CHANGED_MODE) {
            auto* result = new PreviewBitmapResult{token, nullptr};
            if (!PostMessageW(hwnd, WM_PREVIEW_BITMAP_READY, static_cast<WPARAM>(controlId),
                               reinterpret_cast<LPARAM>(result))) {
                delete result;
            }
            return;
        }

        HBITMAP bitmap = LoadPreviewBitmapSync(pathCopy, size);

        if (SUCCEEDED(initHr)) {
            CoUninitialize();
        }

        auto* result = new PreviewBitmapResult{token, bitmap};
        if (!PostMessageW(hwnd, WM_PREVIEW_BITMAP_READY, static_cast<WPARAM>(controlId),
                          reinterpret_cast<LPARAM>(result))) {
            if (result->bitmap) {
                DeleteObject(result->bitmap);
            }
            delete result;
        }
    }).detach();
}

void UpdateUniversalBackgroundPreview(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    RequestPreviewBitmap(hwnd, IDC_CUSTOM_BACKGROUND_PREVIEW,
                         data->workingOptions.universalFolderBackgroundImage.cachedImagePath,
                         kUniversalPreviewSize, &data->universalPreviewToken, &data->universalBackgroundPreview);
    const std::wstring& name = data->workingOptions.universalFolderBackgroundImage.displayName;
    SetDlgItemTextW(hwnd, IDC_CUSTOM_BACKGROUND_UNIVERSAL_NAME, name.empty() ? L"(None)" : name.c_str());
}

void UpdateSelectedFolderBackgroundPreview(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND list = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
    const int selection = GetSelectedFolderBackgroundIndex(list);
    std::wstring name;
    std::wstring path;
    if (selection >= 0 && static_cast<size_t>(selection) < data->workingOptions.folderBackgroundEntries.size()) {
        const auto& entry = data->workingOptions.folderBackgroundEntries[static_cast<size_t>(selection)];
        name = entry.image.displayName;
        path = entry.image.cachedImagePath;
    }
    RequestPreviewBitmap(hwnd, IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW, path, kFolderPreviewSize,
                         &data->folderPreviewToken, &data->folderBackgroundPreview);
    SetDlgItemTextW(hwnd, IDC_CUSTOM_BACKGROUND_FOLDER_NAME, name.c_str());
}

struct PlacementCaptureContext {
    HWND parent = nullptr;
    OptionsDialogData* data = nullptr;
};

BOOL CALLBACK CaptureChildPlacementProc(HWND child, LPARAM param) {
    auto* context = reinterpret_cast<PlacementCaptureContext*>(param);
    if (!context || !context->data) {
        return TRUE;
    }
    if (!IsWindow(child)) {
        return TRUE;
    }
    wchar_t className[32];
    if (GetClassNameW(child, className, ARRAYSIZE(className))) {
        if (_wcsicmp(className, L"ScrollBar") == 0) {
            return TRUE;
        }
    }
    RECT windowRect{};
    if (!GetWindowRect(child, &windowRect)) {
        return TRUE;
    }
    POINT topLeft{windowRect.left, windowRect.top};
    POINT bottomRight{windowRect.right, windowRect.bottom};
    ScreenToClient(context->parent, &topLeft);
    ScreenToClient(context->parent, &bottomRight);
    ChildPlacement placement;
    placement.hwnd = child;
    placement.rect.left = topLeft.x;
    placement.rect.top = topLeft.y;
    placement.rect.right = bottomRight.x;
    placement.rect.bottom = bottomRight.y;
    context->data->customizationChildPlacements.push_back(placement);
    context->data->customizationContentHeight =
        std::max(context->data->customizationContentHeight, static_cast<int>(placement.rect.bottom));
    return TRUE;
}

void CaptureCustomizationChildPlacements(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    data->customizationChildPlacements.clear();
    data->customizationContentHeight = 0;
    PlacementCaptureContext context{hwnd, data};
    EnumChildWindows(hwnd, &CaptureChildPlacementProc, reinterpret_cast<LPARAM>(&context));
}

void RepositionCustomizationChildren(HWND hwnd, OptionsDialogData* data);

bool UpdateCustomizationScrollPosition(HWND hwnd, OptionsDialogData* data, int newPos) {
    if (!data) {
        return false;
    }
    const int clamped = std::clamp(newPos, 0, data->customizationScrollMax);
    if (clamped == data->customizationScrollPos) {
        return false;
    }
    data->customizationScrollPos = clamped;
    SetScrollPos(hwnd, SB_VERT, clamped, TRUE);
    RepositionCustomizationChildren(hwnd, data);
    return true;
}

bool ApplyCustomizationScrollDelta(HWND hwnd, OptionsDialogData* data, int delta) {
    if (!data || delta == 0) {
        return false;
    }
    const int newPos = data->customizationScrollPos + delta;
    return UpdateCustomizationScrollPosition(hwnd, data, newPos);
}

void RepositionCustomizationChildren(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    const size_t childCount = data->customizationChildPlacements.size();
    HDWP deferHandle = childCount > 0
                            ? BeginDeferWindowPos(static_cast<int>(childCount))
                            : nullptr;
    const bool attemptDefer = deferHandle != nullptr;
    for (const auto& placement : data->customizationChildPlacements) {
        if (!IsWindow(placement.hwnd)) {
            continue;
        }
        const int width = placement.rect.right - placement.rect.left;
        const int height = placement.rect.bottom - placement.rect.top;
        const int targetY = placement.rect.top - data->customizationScrollPos;
        if (deferHandle) {
            HDWP nextHandle = DeferWindowPos(deferHandle, placement.hwnd, nullptr, placement.rect.left,
                                            targetY, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
            if (!nextHandle) {
                deferHandle = nullptr;
                SetWindowPos(placement.hwnd, nullptr, placement.rect.left, targetY, width, height,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            } else {
                deferHandle = nextHandle;
            }
        } else {
            SetWindowPos(placement.hwnd, nullptr, placement.rect.left, targetY, width, height,
                         SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }
    if (deferHandle) {
        EndDeferWindowPos(deferHandle);
    }
    if (attemptDefer && !deferHandle) {
        // Ensure any partially moved children receive a layout update if defer failed midway.
        RedrawWindow(hwnd, nullptr, nullptr,
                     RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    }
    RedrawWindow(hwnd, nullptr, nullptr,
                 RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}

void UpdateCustomizationScrollInfo(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    RECT client{};
    if (!GetClientRect(hwnd, &client)) {
        return;
    }
    const int clientHeight = client.bottom - client.top;
    const int contentHeight = std::max(data->customizationContentHeight, clientHeight);
    data->customizationScrollMax = std::max(0, contentHeight - clientHeight);
    if (data->customizationScrollPos > data->customizationScrollMax) {
        data->customizationScrollPos = data->customizationScrollMax;
    }
    SCROLLINFO info{};
    info.cbSize = sizeof(info);
    info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    info.nMin = 0;
    info.nMax = contentHeight;
    info.nPage = std::max(0, clientHeight);
    info.nPos = data->customizationScrollPos;
    SetScrollInfo(hwnd, SB_VERT, &info, TRUE);
    RepositionCustomizationChildren(hwnd, data);
}

void HandleUniversalBackgroundBrowse(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    std::wstring imagePath = data->lastImageBrowseDirectory;
    if (imagePath.empty()) {
        imagePath = data->workingOptions.universalFolderBackgroundImage.cachedImagePath;
    }
    std::wstring displayName;
    const std::wstring initialDirectory = !data->lastImageBrowseDirectory.empty()
                                              ? data->lastImageBrowseDirectory
                                              : ExtractDirectoryFromPath(imagePath);
    if (!BrowseForImage(hwnd, &imagePath, &displayName, initialDirectory)) {
        return;
    }
    CachedImageMetadata metadata = data->workingOptions.universalFolderBackgroundImage;
    std::wstring createdPath;
    std::wstring errorMessage;
    const std::wstring previousPath = metadata.cachedImagePath;
    if (!CopyImageToCache(imagePath, displayName, &metadata, &createdPath, &errorMessage)) {
        std::wstring message = L"Unable to copy the selected image.";
        if (!errorMessage.empty()) {
            message += L"\n\n";
            message += errorMessage;
        }
        MessageBoxW(hwnd, message.c_str(), L"ShellTabs", MB_OK | MB_ICONERROR);
        return;
    }
    data->workingOptions.universalFolderBackgroundImage = metadata;
    if (!createdPath.empty()) {
        TrackCreatedCachedImage(data, createdPath);
    }
    if (!previousPath.empty() && !EqualsInsensitive(previousPath, metadata.cachedImagePath)) {
        ScheduleCachedImageRemoval(data, previousPath);
    }
    data->lastImageBrowseDirectory = ExtractDirectoryFromPath(imagePath);
    UpdateUniversalBackgroundPreview(hwnd, data);
    SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
}

void HandleAddFolderBackgroundEntry(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    std::wstring folder = data->lastFolderBrowsePath;
    if (!BrowseForFolder(hwnd, &folder)) {
        return;
    }
    folder = NormalizeFileSystemPath(folder);
    if (folder.empty()) {
        return;
    }
    for (const auto& entry : data->workingOptions.folderBackgroundEntries) {
        if (EqualsInsensitive(entry.folderPath, folder)) {
            MessageBoxW(hwnd, L"A background for that folder already exists.", L"ShellTabs",
                        MB_OK | MB_ICONWARNING);
            return;
        }
    }

    std::wstring imagePath = data->lastImageBrowseDirectory;
    std::wstring displayName;
    if (!BrowseForImage(hwnd, &imagePath, &displayName, data->lastImageBrowseDirectory)) {
        return;
    }

    CachedImageMetadata metadata;
    std::wstring createdPath;
    std::wstring errorMessage;
    if (!CopyImageToCache(imagePath, displayName, &metadata, &createdPath, &errorMessage)) {
        std::wstring message = L"Unable to copy the selected image.";
        if (!errorMessage.empty()) {
            message += L"\n\n";
            message += errorMessage;
        }
        MessageBoxW(hwnd, message.c_str(), L"ShellTabs", MB_OK | MB_ICONERROR);
        return;
    }
    if (!createdPath.empty()) {
        TrackCreatedCachedImage(data, createdPath);
    }

    FolderBackgroundEntry entry;
    entry.folderPath = folder;
    entry.image = std::move(metadata);
    data->workingOptions.folderBackgroundEntries.emplace_back(std::move(entry));
    data->lastFolderBrowsePath = folder;
    data->lastImageBrowseDirectory = ExtractDirectoryFromPath(imagePath);

    HWND list = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
    RefreshFolderBackgroundListView(list, data);
    const int newIndex = static_cast<int>(data->workingOptions.folderBackgroundEntries.size() - 1);
    ListView_SetItemState(list, newIndex, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    UpdateSelectedFolderBackgroundPreview(hwnd, data);
    UpdateFolderBackgroundButtons(hwnd);
    SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
}

void HandleEditFolderBackgroundEntry(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND list = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
    int selection = GetSelectedFolderBackgroundIndex(list);
    if (selection < 0 || static_cast<size_t>(selection) >= data->workingOptions.folderBackgroundEntries.size()) {
        return;
    }

    FolderBackgroundEntry& entry = data->workingOptions.folderBackgroundEntries[static_cast<size_t>(selection)];
    std::wstring folder = entry.folderPath;
    if (!BrowseForFolder(hwnd, &folder)) {
        return;
    }
    folder = NormalizeFileSystemPath(folder);
    if (folder.empty()) {
        return;
    }
    for (size_t i = 0; i < data->workingOptions.folderBackgroundEntries.size(); ++i) {
        if (i == static_cast<size_t>(selection)) {
            continue;
        }
        if (EqualsInsensitive(data->workingOptions.folderBackgroundEntries[i].folderPath, folder)) {
            MessageBoxW(hwnd, L"A background for that folder already exists.", L"ShellTabs",
                        MB_OK | MB_ICONWARNING);
            return;
        }
    }

    bool changed = false;
    const std::wstring initialDirectory = !data->lastImageBrowseDirectory.empty()
                                              ? data->lastImageBrowseDirectory
                                              : ExtractDirectoryFromPath(entry.image.cachedImagePath);
    std::wstring imagePath = entry.image.cachedImagePath;
    std::wstring displayName = entry.image.displayName;
    if (BrowseForImage(hwnd, &imagePath, &displayName, initialDirectory)) {
        CachedImageMetadata metadata = entry.image;
        std::wstring createdPath;
        std::wstring errorMessage;
        const std::wstring previousPath = metadata.cachedImagePath;
        if (!CopyImageToCache(imagePath, displayName, &metadata, &createdPath, &errorMessage)) {
            std::wstring message = L"Unable to copy the selected image.";
            if (!errorMessage.empty()) {
                message += L"\n\n";
                message += errorMessage;
            }
            MessageBoxW(hwnd, message.c_str(), L"ShellTabs", MB_OK | MB_ICONERROR);
            return;
        }
        entry.image = std::move(metadata);
        if (!createdPath.empty()) {
            TrackCreatedCachedImage(data, createdPath);
        }
        if (!previousPath.empty() && !EqualsInsensitive(previousPath, entry.image.cachedImagePath)) {
            ScheduleCachedImageRemoval(data, previousPath);
        }
        data->lastImageBrowseDirectory = ExtractDirectoryFromPath(imagePath);
        changed = true;
    }

    if (!EqualsInsensitive(entry.folderPath, folder)) {
        entry.folderPath = folder;
        data->lastFolderBrowsePath = folder;
        changed = true;
    }

    if (!changed) {
        return;
    }

    RefreshFolderBackgroundListView(list, data);
    ListView_SetItemState(list, selection, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    UpdateSelectedFolderBackgroundPreview(hwnd, data);
    UpdateFolderBackgroundButtons(hwnd);
    SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
}

void HandleRemoveFolderBackgroundEntry(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND list = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
    int selection = GetSelectedFolderBackgroundIndex(list);
    if (selection < 0 || static_cast<size_t>(selection) >= data->workingOptions.folderBackgroundEntries.size()) {
        return;
    }

    FolderBackgroundEntry removed = data->workingOptions.folderBackgroundEntries[static_cast<size_t>(selection)];
    data->workingOptions.folderBackgroundEntries.erase(
        data->workingOptions.folderBackgroundEntries.begin() + static_cast<ptrdiff_t>(selection));
    ScheduleCachedImageRemoval(data, removed.image.cachedImagePath);

    RefreshFolderBackgroundListView(list, data);
    const int newCount = ListView_GetItemCount(list);
    if (newCount > 0) {
        int newSelection = selection;
        if (newSelection >= newCount) {
            newSelection = newCount - 1;
        }
        ListView_SetItemState(list, newSelection, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }
    UpdateSelectedFolderBackgroundPreview(hwnd, data);
    UpdateFolderBackgroundButtons(hwnd);
    SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
}

struct GroupEditorContext {
    SavedGroup working;
    std::wstring originalName;
    bool isNew = false;
    HBRUSH colorBrush = nullptr;
    const std::vector<SavedGroup>* existingGroups = nullptr;
};

struct OutlineStyleOption {
    TabGroupOutlineStyle style;
    const wchar_t* label;
};

constexpr OutlineStyleOption kOutlineStyleOptions[] = {
    {TabGroupOutlineStyle::kSolid, L"Solid"},
    {TabGroupOutlineStyle::kDashed, L"Dashed"},
    {TabGroupOutlineStyle::kDotted, L"Dotted"},
};

int OutlineStyleIndexForStyle(TabGroupOutlineStyle style) {
    for (size_t i = 0; i < ARRAYSIZE(kOutlineStyleOptions); ++i) {
        if (kOutlineStyleOptions[i].style == style) {
            return static_cast<int>(i);
        }
    }
    return 0;
}

TabGroupOutlineStyle OutlineStyleFromIndex(LRESULT index) {
    if (index < 0 || index >= static_cast<LRESULT>(ARRAYSIZE(kOutlineStyleOptions))) {
        return TabGroupOutlineStyle::kSolid;
    }
    return kOutlineStyleOptions[static_cast<size_t>(index)].style;
}

void PopulateOutlineStyleCombo(HWND combo) {
    if (!combo) {
        return;
    }
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (const auto& option : kOutlineStyleOptions) {
        SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(option.label));
    }
}

void UpdateListBoxHorizontalExtent(HWND hwndList) {
    if (!hwndList) {
        return;
    }
    SendMessageW(hwndList, LB_SETHORIZONTALEXTENT, 0, 0);
    const int count = static_cast<int>(SendMessageW(hwndList, LB_GETCOUNT, 0, 0));
    if (count <= 0) {
        return;
    }

    HDC dc = GetDC(hwndList);
    if (!dc) {
        return;
    }
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwndList, WM_GETFONT, 0, 0));
    HFONT oldFont = nullptr;
    if (font) {
        oldFont = static_cast<HFONT>(SelectObject(dc, font));
    }

    int maxWidth = 0;
    for (int i = 0; i < count; ++i) {
        LRESULT length = SendMessageW(hwndList, LB_GETTEXTLEN, i, 0);
        if (length == LB_ERR || length <= 0) {
            continue;
        }
        std::wstring text(static_cast<size_t>(length) + 1, L'\0');
        LRESULT copied = SendMessageW(hwndList, LB_GETTEXT, i, reinterpret_cast<LPARAM>(text.data()));
        if (copied == LB_ERR) {
            continue;
        }
        text.resize(static_cast<size_t>(copied));
        SIZE size{};
        if (GetTextExtentPoint32W(dc, text.c_str(), static_cast<int>(text.size()), &size)) {
            maxWidth = std::max(maxWidth, static_cast<int>(size.cx));
        }
    }

    if (oldFont) {
        SelectObject(dc, oldFont);
    }
    ReleaseDC(hwndList, dc);
    SendMessageW(hwndList, LB_SETHORIZONTALEXTENT, maxWidth + 12, 0);
}

int ClampPercentageValue(int value) {
    return std::clamp(value, 0, 100);
}

int InvertPercentageValue(int value) {
    return 100 - ClampPercentageValue(value);
}

void ConfigurePercentageSlider(HWND hwnd, int controlId, int value) {
    HWND slider = GetDlgItem(hwnd, controlId);
    if (!slider) {
        return;
    }
    SendMessageW(slider, TBM_SETRANGE, TRUE, MAKELPARAM(0, 100));
    SendMessageW(slider, TBM_SETPAGESIZE, 0, 5);
    SendMessageW(slider, TBM_SETLINESIZE, 0, 1);
    SendMessageW(slider, TBM_SETTICFREQ, 10, 0);
    SendMessageW(slider, TBM_SETPOS, TRUE, ClampPercentageValue(value));
}

void UpdatePercentageLabel(HWND hwnd, int controlId, int value) {
    wchar_t buffer[16];
    const int clamped = ClampPercentageValue(value);
    _snwprintf_s(buffer, ARRAYSIZE(buffer), _TRUNCATE, L"%d%%", clamped);
    SetDlgItemTextW(hwnd, controlId, buffer);
}

int ClampMultiplierValue(int value) {
    return std::clamp(value, 0, 200);
}

void ConfigureMultiplierSlider(HWND hwnd, int controlId, int value) {
    HWND slider = GetDlgItem(hwnd, controlId);
    if (!slider) {
        return;
    }
    SendMessageW(slider, TBM_SETRANGE, TRUE, MAKELPARAM(0, 200));
    SendMessageW(slider, TBM_SETPAGESIZE, 0, 10);
    SendMessageW(slider, TBM_SETLINESIZE, 0, 2);
    SendMessageW(slider, TBM_SETTICFREQ, 20, 0);
    SendMessageW(slider, TBM_SETPOS, TRUE, ClampMultiplierValue(value));
}

void UpdateMultiplierLabel(HWND hwnd, int controlId, int value) {
    wchar_t buffer[16];
    const int clamped = ClampMultiplierValue(value);
    _snwprintf_s(buffer, ARRAYSIZE(buffer), _TRUNCATE, L"%d%%", clamped);
    SetDlgItemTextW(hwnd, controlId, buffer);
}

void UpdateGradientControlsEnabled(HWND hwnd, bool backgroundEnabled, bool fontEnabled) {
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_BG_LABEL), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_BG_SLIDER), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_BG_VALUE), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_FONT_LABEL), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_FONT_SLIDER), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_FONT_VALUE), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_LABEL), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_VALUE), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_LABEL), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_VALUE), fontEnabled);
}

void UpdateGradientColorControlsEnabled(HWND hwnd, bool backgroundEnabled, bool fontEnabled) {
    const int bgControls[] = {IDC_MAIN_BREADCRUMB_BG_START_LABEL,   IDC_MAIN_BREADCRUMB_BG_START_PREVIEW,
                              IDC_MAIN_BREADCRUMB_BG_START_BUTTON, IDC_MAIN_BREADCRUMB_BG_END_LABEL,
                              IDC_MAIN_BREADCRUMB_BG_END_PREVIEW,  IDC_MAIN_BREADCRUMB_BG_END_BUTTON};
    for (int id : bgControls) {
        EnableWindow(GetDlgItem(hwnd, id), backgroundEnabled);
    }

    const int fontControls[] = {IDC_MAIN_BREADCRUMB_FONT_START_LABEL,   IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW,
                                IDC_MAIN_BREADCRUMB_FONT_START_BUTTON, IDC_MAIN_BREADCRUMB_FONT_END_LABEL,
                                IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW,  IDC_MAIN_BREADCRUMB_FONT_END_BUTTON};
    for (int id : fontControls) {
        EnableWindow(GetDlgItem(hwnd, id), fontEnabled);
    }
}

void UpdateProgressColorControlsEnabled(HWND hwnd, bool enabled) {
    const int controls[] = {IDC_MAIN_PROGRESS_START_LABEL,   IDC_MAIN_PROGRESS_START_PREVIEW,
                            IDC_MAIN_PROGRESS_START_BUTTON, IDC_MAIN_PROGRESS_END_LABEL,
                            IDC_MAIN_PROGRESS_END_PREVIEW,  IDC_MAIN_PROGRESS_END_BUTTON};
    for (int id : controls) {
        EnableWindow(GetDlgItem(hwnd, id), enabled);
    }
}

void UpdateTabColorControlsEnabled(HWND hwnd, bool selectedEnabled, bool unselectedEnabled) {
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_TAB_SELECTED_PREVIEW), selectedEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_TAB_SELECTED_BUTTON), selectedEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_TAB_UNSELECTED_PREVIEW), unselectedEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_TAB_UNSELECTED_BUTTON), unselectedEnabled);
}

void UpdateGlowControlStates(HWND hwnd) {
    const bool glowEnabled = IsDlgButtonChecked(hwnd, IDC_GLOW_ENABLE) == BST_CHECKED;
    EnableWindow(GetDlgItem(hwnd, IDC_GLOW_CUSTOM_COLORS), glowEnabled);

    for (const auto& mapping : kGlowSurfaceControlMappings) {
        EnableWindow(GetDlgItem(hwnd, mapping.controlId), glowEnabled);
    }

    const bool customColors = glowEnabled &&
                              IsDlgButtonChecked(hwnd, IDC_GLOW_CUSTOM_COLORS) == BST_CHECKED;
    EnableWindow(GetDlgItem(hwnd, IDC_GLOW_USE_GRADIENT), customColors);

    const int primaryControls[] = {IDC_GLOW_PRIMARY_LABEL, IDC_GLOW_PRIMARY_PREVIEW,
                                   IDC_GLOW_PRIMARY_BUTTON};
    for (int id : primaryControls) {
        EnableWindow(GetDlgItem(hwnd, id), customColors);
    }

    const bool gradientEnabled = customColors &&
                                 IsDlgButtonChecked(hwnd, IDC_GLOW_USE_GRADIENT) == BST_CHECKED;
    const int secondaryControls[] = {IDC_GLOW_SECONDARY_LABEL, IDC_GLOW_SECONDARY_PREVIEW,
                                     IDC_GLOW_SECONDARY_BUTTON};
    for (int id : secondaryControls) {
        EnableWindow(GetDlgItem(hwnd, id), gradientEnabled);
    }
}

void SetPreviewColor(HWND hwnd, int controlId, HBRUSH* brush, COLORREF color) {
    if (!brush) {
        return;
    }
    if (*brush) {
        DeleteObject(*brush);
        *brush = nullptr;
    }
    *brush = CreateSolidBrush(color);
    if (HWND ctrl = GetDlgItem(hwnd, controlId)) {
        InvalidateRect(ctrl, nullptr, TRUE);
    }
}

void RefreshGlowControls(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }

    CheckDlgButton(hwnd, IDC_GLOW_ENABLE,
                   data->workingOptions.enableNeonGlow ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hwnd, IDC_GLOW_CUSTOM_COLORS,
                   data->workingOptions.useCustomNeonGlowColors ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hwnd, IDC_GLOW_USE_GRADIENT,
                   data->workingOptions.useNeonGlowGradient ? BST_CHECKED : BST_UNCHECKED);

    for (const auto& mapping : kGlowSurfaceControlMappings) {
        const GlowSurfaceOptions& surface = data->workingOptions.glowPalette.*(mapping.member);
        CheckDlgButton(hwnd, mapping.controlId, surface.enabled ? BST_CHECKED : BST_UNCHECKED);
    }

    SetPreviewColor(hwnd, IDC_GLOW_PRIMARY_PREVIEW, &data->glowPrimaryBrush,
                    data->workingOptions.neonGlowPrimaryColor);
    SetPreviewColor(hwnd, IDC_GLOW_SECONDARY_PREVIEW, &data->glowSecondaryBrush,
                    data->workingOptions.neonGlowSecondaryColor);

    UpdateGlowControlStates(hwnd);
}

bool HandleColorButtonClick(HWND hwnd, OptionsDialogData* data, WORD controlId) {
    if (!data) {
        return false;
    }

    COLORREF initial = RGB(255, 255, 255);
    HBRUSH* targetBrush = nullptr;
    int previewId = 0;
    COLORREF* targetColor = nullptr;
    const bool glowColorControl =
        (controlId == IDC_GLOW_PRIMARY_BUTTON || controlId == IDC_GLOW_SECONDARY_BUTTON);

    switch (controlId) {
        case IDC_MAIN_BREADCRUMB_BG_START_BUTTON:
            initial = data->workingOptions.breadcrumbGradientStartColor;
            targetBrush = &data->breadcrumbBgStartBrush;
            previewId = IDC_MAIN_BREADCRUMB_BG_START_PREVIEW;
            targetColor = &data->workingOptions.breadcrumbGradientStartColor;
            break;
        case IDC_MAIN_BREADCRUMB_BG_END_BUTTON:
            initial = data->workingOptions.breadcrumbGradientEndColor;
            targetBrush = &data->breadcrumbBgEndBrush;
            previewId = IDC_MAIN_BREADCRUMB_BG_END_PREVIEW;
            targetColor = &data->workingOptions.breadcrumbGradientEndColor;
            break;
        case IDC_MAIN_BREADCRUMB_FONT_START_BUTTON:
            initial = data->workingOptions.breadcrumbFontGradientStartColor;
            targetBrush = &data->breadcrumbFontStartBrush;
            previewId = IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW;
            targetColor = &data->workingOptions.breadcrumbFontGradientStartColor;
            break;
        case IDC_MAIN_BREADCRUMB_FONT_END_BUTTON:
            initial = data->workingOptions.breadcrumbFontGradientEndColor;
            targetBrush = &data->breadcrumbFontEndBrush;
            previewId = IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW;
            targetColor = &data->workingOptions.breadcrumbFontGradientEndColor;
            break;
        case IDC_MAIN_PROGRESS_START_BUTTON:
            initial = data->workingOptions.progressBarGradientStartColor;
            targetBrush = &data->progressStartBrush;
            previewId = IDC_MAIN_PROGRESS_START_PREVIEW;
            targetColor = &data->workingOptions.progressBarGradientStartColor;
            break;
        case IDC_MAIN_PROGRESS_END_BUTTON:
            initial = data->workingOptions.progressBarGradientEndColor;
            targetBrush = &data->progressEndBrush;
            previewId = IDC_MAIN_PROGRESS_END_PREVIEW;
            targetColor = &data->workingOptions.progressBarGradientEndColor;
            break;
        case IDC_MAIN_TAB_SELECTED_BUTTON:
            initial = data->workingOptions.customTabSelectedColor;
            targetBrush = &data->tabSelectedBrush;
            previewId = IDC_MAIN_TAB_SELECTED_PREVIEW;
            targetColor = &data->workingOptions.customTabSelectedColor;
            break;
        case IDC_MAIN_TAB_UNSELECTED_BUTTON:
            initial = data->workingOptions.customTabUnselectedColor;
            targetBrush = &data->tabUnselectedBrush;
            previewId = IDC_MAIN_TAB_UNSELECTED_PREVIEW;
            targetColor = &data->workingOptions.customTabUnselectedColor;
            break;
        case IDC_GLOW_PRIMARY_BUTTON:
            initial = data->workingOptions.neonGlowPrimaryColor;
            targetBrush = &data->glowPrimaryBrush;
            previewId = IDC_GLOW_PRIMARY_PREVIEW;
            targetColor = &data->workingOptions.neonGlowPrimaryColor;
            break;
        case IDC_GLOW_SECONDARY_BUTTON:
            initial = data->workingOptions.neonGlowSecondaryColor;
            targetBrush = &data->glowSecondaryBrush;
            previewId = IDC_GLOW_SECONDARY_PREVIEW;
            targetColor = &data->workingOptions.neonGlowSecondaryColor;
            break;
        default:
            return false;
    }

    if (targetColor && PromptForColor(hwnd, initial, targetColor)) {
        SetPreviewColor(hwnd, previewId, targetBrush, *targetColor);
        if (glowColorControl) {
            UpdateGlowPaletteFromLegacySettings(data->workingOptions);
        }
        return true;
    }
    return false;
}

void RefreshGroupList(HWND hwndList, const OptionsDialogData* data) {
    SendMessageW(hwndList, LB_RESETCONTENT, 0, 0);
    SendMessageW(hwndList, LB_SETHORIZONTALEXTENT, 0, 0);
    if (!data) {
        return;
    }
    for (const auto& group : data->workingGroups) {
        SendMessageW(hwndList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(group.name.c_str()));
    }
    UpdateListBoxHorizontalExtent(hwndList);
}

std::wstring GetSelectedGroupName(HWND hwndList) {
    const LRESULT index = SendMessageW(hwndList, LB_GETCURSEL, 0, 0);
    if (index == LB_ERR) {
        return {};
    }
    LRESULT length = SendMessageW(hwndList, LB_GETTEXTLEN, index, 0);
    if (length == LB_ERR) {
        return {};
    }
    std::wstring name(static_cast<size_t>(length) + 1, L'\0');
    LRESULT copied = SendMessageW(hwndList, LB_GETTEXT, index, reinterpret_cast<LPARAM>(name.data()));
    if (copied == LB_ERR) {
        return {};
    }
    name.resize(static_cast<size_t>(copied));
    return name;
}

void UpdateGroupButtons(HWND page) {
    HWND list = GetDlgItem(page, IDC_GROUP_LIST);
    HWND editButton = GetDlgItem(page, IDC_GROUP_EDIT);
    HWND removeButton = GetDlgItem(page, IDC_GROUP_REMOVE);
    const bool hasSelection = SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR;
    EnableWindow(editButton, hasSelection);
    EnableWindow(removeButton, hasSelection);
}

void GenerateRandomColor(COLORREF* color) {
    if (!color) {
        return;
    }
    std::mt19937 rng{static_cast<unsigned int>(GetTickCount64())};
    std::uniform_int_distribution<int> dist(0, 255);
    *color = RGB(dist(rng), dist(rng), dist(rng));
}

bool CaseInsensitiveEquals(const std::wstring& left, const std::wstring& right) {
    return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

std::vector<SavedGroup> NormalizeGroups(const std::vector<SavedGroup>& groups) {
    std::vector<SavedGroup> normalized = groups;
    std::sort(normalized.begin(), normalized.end(),
              [](const SavedGroup& a, const SavedGroup& b) {
                  return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0;
              });
    return normalized;
}

bool AreSavedGroupsEqual(const std::vector<SavedGroup>& left, const std::vector<SavedGroup>& right) {
    auto normalizedLeft = NormalizeGroups(left);
    auto normalizedRight = NormalizeGroups(right);
    if (normalizedLeft.size() != normalizedRight.size()) {
        return false;
    }
    for (size_t i = 0; i < normalizedLeft.size(); ++i) {
        const auto& lhs = normalizedLeft[i];
        const auto& rhs = normalizedRight[i];
        if (!CaseInsensitiveEquals(lhs.name, rhs.name)) {
            return false;
        }
        if (lhs.color != rhs.color) {
            return false;
        }
        if (lhs.outlineStyle != rhs.outlineStyle) {
            return false;
        }
        if (lhs.tabPaths.size() != rhs.tabPaths.size()) {
            return false;
        }
        for (size_t j = 0; j < lhs.tabPaths.size(); ++j) {
            if (_wcsicmp(lhs.tabPaths[j].c_str(), rhs.tabPaths[j].c_str()) != 0) {
                return false;
            }
        }
    }
    return true;
}

void UpdatePathButtons(HWND dialog) {
    HWND list = GetDlgItem(dialog, IDC_EDITOR_PATH_LIST);
    const bool hasSelection = SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR;
    EnableWindow(GetDlgItem(dialog, IDC_EDITOR_EDIT_PATH), hasSelection);
    EnableWindow(GetDlgItem(dialog, IDC_EDITOR_REMOVE_PATH), hasSelection);
}

void RefreshPathList(HWND dialog, const GroupEditorContext& context) {
    HWND list = GetDlgItem(dialog, IDC_EDITOR_PATH_LIST);
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    SendMessageW(list, LB_SETHORIZONTALEXTENT, 0, 0);
    for (const auto& path : context.working.tabPaths) {
        SendMessageW(list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(path.c_str()));
    }
    UpdateListBoxHorizontalExtent(list);
    UpdatePathButtons(dialog);
}

bool ValidateUniqueName(const std::wstring& name, const std::wstring& original,
                        const std::vector<SavedGroup>* groups) {
    if (_wcsicmp(name.c_str(), original.c_str()) == 0) {
        return true;
    }
    if (!groups) {
        return true;
    }
    for (const auto& group : *groups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            return false;
        }
    }
    return true;
}

INT_PTR CALLBACK GroupEditorProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* context = reinterpret_cast<GroupEditorContext*>(lParam);
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(context));
            if (context) {
                SetDlgItemTextW(hwnd, IDC_EDITOR_NAME, context->working.name.c_str());
                if (context->working.tabPaths.empty()) {
                    context->working.tabPaths.emplace_back(L"C:\\");
                }
                if (!context->colorBrush) {
                    context->colorBrush = CreateSolidBrush(context->working.color);
                }
                HWND styleCombo = GetDlgItem(hwnd, IDC_EDITOR_STYLE_COMBO);
                PopulateOutlineStyleCombo(styleCombo);
                if (styleCombo) {
                    const int index = OutlineStyleIndexForStyle(context->working.outlineStyle);
                    SendMessageW(styleCombo, CB_SETCURSEL, index, 0);
                }
                RefreshPathList(hwnd, *context);
            }
            return TRUE;
        }
        case WM_CTLCOLORSTATIC: {
            auto* context = reinterpret_cast<GroupEditorContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (context) {
                HWND target = reinterpret_cast<HWND>(lParam);
                if (GetDlgCtrlID(target) == IDC_EDITOR_COLOR_PREVIEW) {
                    if (!context->colorBrush) {
                        context->colorBrush = CreateSolidBrush(context->working.color);
                    }
                    SetBkMode(reinterpret_cast<HDC>(wParam), OPAQUE);
                    SetBkColor(reinterpret_cast<HDC>(wParam), context->working.color);
                    return reinterpret_cast<INT_PTR>(context->colorBrush);
                }
            }
            break;
        }
        case WM_COMMAND: {
            auto* context = reinterpret_cast<GroupEditorContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!context) {
                break;
            }
            switch (LOWORD(wParam)) {
                case IDC_EDITOR_ADD_PATH: {
                    std::wstring path;
                    if (BrowseForFolder(hwnd, &path)) {
                        context->working.tabPaths.push_back(path);
                        RefreshPathList(hwnd, *context);
                    }
                    return TRUE;
                }
                case IDC_EDITOR_EDIT_PATH: {
                    HWND list = GetDlgItem(hwnd, IDC_EDITOR_PATH_LIST);
                    const LRESULT index = SendMessageW(list, LB_GETCURSEL, 0, 0);
                    if (index != LB_ERR) {
                        std::wstring path = context->working.tabPaths[static_cast<size_t>(index)];
                        if (BrowseForFolder(hwnd, &path)) {
                            context->working.tabPaths[static_cast<size_t>(index)] = path;
                            RefreshPathList(hwnd, *context);
                            SendMessageW(list, LB_SETCURSEL, index, 0);
                        }
                    }
                    return TRUE;
                }
                case IDC_EDITOR_REMOVE_PATH: {
                    HWND list = GetDlgItem(hwnd, IDC_EDITOR_PATH_LIST);
                    const LRESULT index = SendMessageW(list, LB_GETCURSEL, 0, 0);
                    if (index != LB_ERR) {
                        context->working.tabPaths.erase(
                            context->working.tabPaths.begin() + static_cast<size_t>(index));
                        RefreshPathList(hwnd, *context);
                    }
                    return TRUE;
                }
                case IDC_EDITOR_COLOR_BUTTON: {
                    COLORREF color = context->working.color;
                    if (PromptForColor(hwnd, color, &color)) {
                        context->working.color = color;
                        if (context->colorBrush) {
                            DeleteObject(context->colorBrush);
                            context->colorBrush = nullptr;
                        }
                        InvalidateRect(GetDlgItem(hwnd, IDC_EDITOR_COLOR_PREVIEW), nullptr, TRUE);
                    }
                    return TRUE;
                }
                case IDC_EDITOR_PATH_LIST: {
                    if (HIWORD(wParam) == LBN_SELCHANGE) {
                        UpdatePathButtons(hwnd);
                    }
                    return TRUE;
                }
                case IDC_EDITOR_STYLE_COMBO: {
                    if (HIWORD(wParam) == CBN_SELCHANGE) {
                        HWND combo = reinterpret_cast<HWND>(lParam);
                        if (combo) {
                            const LRESULT selection = SendMessageW(combo, CB_GETCURSEL, 0, 0);
                            context->working.outlineStyle = OutlineStyleFromIndex(selection);
                        }
                    }
                    return TRUE;
                }
                case IDOK: {
                    wchar_t nameBuffer[256];
                    GetDlgItemTextW(hwnd, IDC_EDITOR_NAME, nameBuffer, ARRAYSIZE(nameBuffer));
                    std::wstring name = nameBuffer;
                    if (name.empty()) {
                        MessageBoxW(hwnd, L"Group name cannot be empty.", L"ShellTabs", MB_OK | MB_ICONWARNING);
                        return TRUE;
                    }
                    if (!ValidateUniqueName(name, context->originalName, context->existingGroups)) {
                        MessageBoxW(hwnd, L"A group with that name already exists.", L"ShellTabs",
                                     MB_OK | MB_ICONWARNING);
                        return TRUE;
                    }
                    context->working.name = std::move(name);
                    HWND combo = GetDlgItem(hwnd, IDC_EDITOR_STYLE_COMBO);
                    if (combo) {
                        const LRESULT selection = SendMessageW(combo, CB_GETCURSEL, 0, 0);
                        context->working.outlineStyle = OutlineStyleFromIndex(selection);
                    }
                    if (context->working.tabPaths.empty()) {
                        context->working.tabPaths.emplace_back(L"C:\\");
                    }
                    EndDialog(hwnd, IDOK);
                    return TRUE;
                }
                case IDCANCEL: {
                    EndDialog(hwnd, IDCANCEL);
                    return TRUE;
                }
            }
            break;
        }
        case WM_DESTROY: {
            auto* context = reinterpret_cast<GroupEditorContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (context && context->colorBrush) {
                DeleteObject(context->colorBrush);
                context->colorBrush = nullptr;
            }
            break;
        }
    }
    return FALSE;
}

bool RunGroupEditor(HWND parent, const SavedGroup* existing, SavedGroup* result,
                    const std::vector<SavedGroup>* groups) {
    if (!result) {
        return false;
    }
    GroupEditorContext context;
    if (existing) {
        context.working = *existing;
        context.originalName = existing->name;
    } else {
        context.working.name = L"New Group";
        GenerateRandomColor(&context.working.color);
        context.working.tabPaths = {L"C:\\"};
        context.originalName = context.working.name;
        context.isNew = true;
    }
    context.existingGroups = groups;

    std::vector<BYTE> dialogTemplate = BuildGroupEditorTemplate();
    INT_PTR resultCode = DialogBoxIndirectParamW(GetModuleHandleInstance(),
                                                 reinterpret_cast<DLGTEMPLATE*>(dialogTemplate.data()), parent,
                                                 GroupEditorProc, reinterpret_cast<LPARAM>(&context));
    if (resultCode == IDOK) {
        *result = context.working;
        return true;
    }
    return false;
}

void HandleNewGroup(HWND page, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    SavedGroup group;
    if (!RunGroupEditor(GetAncestor(page, GA_ROOT), nullptr, &group, &data->workingGroups)) {
        return;
    }

    data->workingGroups.push_back(group);
    data->workingGroupIds.push_back(L"");
    data->groupsChanged = true;
    HWND list = GetDlgItem(page, IDC_GROUP_LIST);
    RefreshGroupList(list, data);
    const int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    if (count > 0) {
        SendMessageW(list, LB_SETCURSEL, count - 1, 0);
    }
    SendMessageW(GetParent(page), PSM_CHANGED, reinterpret_cast<WPARAM>(page), 0);
    UpdateGroupButtons(page);
}

void HandleEditGroup(HWND page, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND list = GetDlgItem(page, IDC_GROUP_LIST);
    const LRESULT index = SendMessageW(list, LB_GETCURSEL, 0, 0);
    if (index == LB_ERR) {
        return;
    }
    if (index < 0 || static_cast<size_t>(index) >= data->workingGroups.size()) {
        return;
    }
    const SavedGroup existing = data->workingGroups[static_cast<size_t>(index)];
    SavedGroup updated;
    if (!RunGroupEditor(GetAncestor(page, GA_ROOT), &existing, &updated, &data->workingGroups)) {
        return;
    }

    data->workingGroups[static_cast<size_t>(index)] = updated;
    data->groupsChanged = true;
    RefreshGroupList(list, data);
    SendMessageW(list, LB_SETCURSEL, index, 0);
    SendMessageW(GetParent(page), PSM_CHANGED, reinterpret_cast<WPARAM>(page), 0);
    UpdateGroupButtons(page);
}

void HandleRemoveGroup(HWND page, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND list = GetDlgItem(page, IDC_GROUP_LIST);
    const LRESULT index = SendMessageW(list, LB_GETCURSEL, 0, 0);
    if (index == LB_ERR) {
        return;
    }
    if (MessageBoxW(GetAncestor(page, GA_ROOT), L"Remove the selected group?", L"ShellTabs",
                    MB_YESNO | MB_ICONQUESTION) != IDYES) {
        return;
    }
    if (index < 0 || static_cast<size_t>(index) >= data->workingGroups.size()) {
        return;
    }
    if (static_cast<size_t>(index) < data->workingGroupIds.size()) {
        std::wstring removedId = data->workingGroupIds[static_cast<size_t>(index)];
        if (!removedId.empty()) {
            data->removedGroupIds.push_back(std::move(removedId));
        }
        data->workingGroupIds.erase(data->workingGroupIds.begin() + static_cast<size_t>(index));
    }
    data->workingGroups.erase(data->workingGroups.begin() + static_cast<size_t>(index));
    data->groupsChanged = true;
    RefreshGroupList(list, data);
    const LRESULT count = SendMessageW(list, LB_GETCOUNT, 0, 0);
    if (count > 0) {
        LRESULT newIndex = index;
        if (newIndex >= count) {
            newIndex = count - 1;
        }
        SendMessageW(list, LB_SETCURSEL, newIndex, 0);
    }
    SendMessageW(GetParent(page), PSM_CHANGED, reinterpret_cast<WPARAM>(page), 0);
    UpdateGroupButtons(page);
}

void ApplyFocusedGroupSelection(HWND page, OptionsDialogData* data) {
    if (!data || data->focusHandled || data->focusSavedGroupId.empty()) {
        return;
    }

    HWND list = GetDlgItem(page, IDC_GROUP_LIST);
    if (!list) {
        data->focusHandled = true;
        return;
    }

    int targetIndex = -1;
    for (size_t i = 0; i < data->workingGroups.size(); ++i) {
        if (CaseInsensitiveEquals(data->workingGroups[i].name, data->focusSavedGroupId)) {
            targetIndex = static_cast<int>(i);
            break;
        }
        if (i < data->workingGroupIds.size() &&
            CaseInsensitiveEquals(data->workingGroupIds[i], data->focusSavedGroupId)) {
            targetIndex = static_cast<int>(i);
            break;
        }
    }

    data->focusHandled = true;
    if (targetIndex < 0) {
        return;
    }

    SendMessageW(list, LB_SETCURSEL, targetIndex, 0);
    UpdateGroupButtons(page);
    if (data->focusShouldEdit) {
        HandleEditGroup(page, data);
    }
}

INT_PTR CALLBACK MainOptionsPageProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* data = reinterpret_cast<OptionsDialogData*>(reinterpret_cast<PROPSHEETPAGEW*>(lParam)->lParam);
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(data));
            if (data) {
                CheckDlgButton(hwnd, IDC_MAIN_REOPEN, data->workingOptions.reopenOnCrash ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(hwnd, IDC_MAIN_PERSIST,
                               data->workingOptions.persistGroupPaths ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(hwnd, IDC_MAIN_LISTVIEW_ACCENT,
                               data->workingOptions.useExplorerAccentColors ? BST_CHECKED : BST_UNCHECKED);
                const wchar_t example[] =
                    L"Example: if a group opens to C:\\test and you browse to C\\test\\child, "
                    L"enabling this option reopens the child folder next time.";
                SetDlgItemTextW(hwnd, IDC_MAIN_EXAMPLE, example);

                PopulateNewTabTemplateCombo(hwnd, data);
                SetDlgItemTextW(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT, data->workingOptions.newTabCustomPath.c_str());
                PopulateNewTabGroupCombo(hwnd, data);
                UpdateNewTabTemplateControls(hwnd, data);

                HWND combo = GetDlgItem(hwnd, IDC_MAIN_DOCK_COMBO);
                if (combo) {
                    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
                    const uint32_t mask = TabBandWindow::GetAvailableDockMask();
                    struct DockEntry {
                        TabBandDockMode mode;
                        const wchar_t* label;
                        uint32_t requiredMask;
                    } entries[] = {
                        {TabBandDockMode::kAutomatic, L"Let Explorer decide", 0},
                        {TabBandDockMode::kTop, L"Top toolbar", 1u << static_cast<uint32_t>(TabBandDockMode::kTop)},
                        {TabBandDockMode::kBottom, L"Bottom toolbar", 1u << static_cast<uint32_t>(TabBandDockMode::kBottom)},
                        {TabBandDockMode::kLeft, L"Left vertical band", 1u << static_cast<uint32_t>(TabBandDockMode::kLeft)},
                        {TabBandDockMode::kRight, L"Right vertical band", 1u << static_cast<uint32_t>(TabBandDockMode::kRight)},
                    };

                    int selectionIndex = -1;
                    for (const auto& entry : entries) {
                        if (entry.mode != TabBandDockMode::kAutomatic && entry.requiredMask != 0 &&
                            (mask & entry.requiredMask) == 0) {
                            continue;
                        }

                        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0,
                                                                         reinterpret_cast<LPARAM>(entry.label)));
                        if (index >= 0) {
                            SendMessageW(combo, CB_SETITEMDATA, index,
                                         static_cast<LPARAM>(entry.mode));
                            if (data->workingOptions.tabDockMode == entry.mode && selectionIndex < 0) {
                                selectionIndex = index;
                            }
                        }
                    }

                    if (selectionIndex < 0) {
                        selectionIndex = 0;
                    }
                    SendMessageW(combo, CB_SETCURSEL, selectionIndex, 0);
                }
            }
            return TRUE;
        }
        case WM_CTLCOLORDLG: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (dc) {
                SetBkColor(dc, GetSysColor(COLOR_3DFACE));
            }
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
        }
        case WM_COMMAND: {
            switch (LOWORD(wParam)) {
                case IDC_MAIN_REOPEN:
                case IDC_MAIN_PERSIST:
                case IDC_MAIN_LISTVIEW_ACCENT:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_MAIN_NEW_TAB_COMBO:
                    if (HIWORD(wParam) == CBN_SELCHANGE) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        UpdateNewTabTemplateControls(hwnd, data);
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_MAIN_NEW_TAB_PATH_EDIT:
                    if (HIWORD(wParam) == EN_CHANGE) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data) {
                            HWND edit = reinterpret_cast<HWND>(lParam);
                            if (!edit) {
                                edit = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT);
                            }
                            data->workingOptions.newTabCustomPath = Trim(GetWindowTextString(edit));
                        }
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_MAIN_NEW_TAB_BROWSE:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        std::wstring path = Trim(GetWindowTextString(GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT)));
                        if (path.empty() && data) {
                            path = data->workingOptions.newTabCustomPath;
                        }
                        if (BrowseForFolder(hwnd, &path)) {
                            path = Trim(path);
                            SetDlgItemTextW(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT, path.c_str());
                            if (data) {
                                data->workingOptions.newTabCustomPath = path;
                            }
                            SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                        }
                    }
                    return TRUE;
                case IDC_MAIN_NEW_TAB_GROUP_COMBO:
                    if (HIWORD(wParam) == CBN_SELCHANGE) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data) {
                            HWND combo = reinterpret_cast<HWND>(lParam);
                            if (!combo) {
                                combo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_GROUP_COMBO);
                            }
                            std::wstring selected = Trim(GetWindowTextString(combo));
                            if (data->workingGroups.empty()) {
                                data->workingOptions.newTabSavedGroup.clear();
                            } else {
                                data->workingOptions.newTabSavedGroup = selected;
                            }
                        }
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_MAIN_DOCK_COMBO:
                    if (HIWORD(wParam) == CBN_SELCHANGE) {
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                default:
                    break;
            }
            break;
        }
        case WM_NOTIFY: {
            const auto code = reinterpret_cast<LPNMHDR>(lParam)->code;
            if (code == PSN_SETACTIVE) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                if (data) {
                    SetDlgItemTextW(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT, data->workingOptions.newTabCustomPath.c_str());
                    PopulateNewTabGroupCombo(hwnd, data);
                    UpdateNewTabTemplateControls(hwnd, data);
                }
                SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, 0);
                return TRUE;
            }
            if (code == PSN_APPLY) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                if (data) {
                    data->workingOptions.reopenOnCrash =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_REOPEN) == BST_CHECKED;
                    data->workingOptions.persistGroupPaths =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_PERSIST) == BST_CHECKED;
                    data->workingOptions.useExplorerAccentColors =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_LISTVIEW_ACCENT) == BST_CHECKED;
                    UpdateGlowPaletteFromLegacySettings(data->workingOptions);

                    if (HWND templateCombo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_COMBO)) {
                        const LRESULT selection = SendMessageW(templateCombo, CB_GETCURSEL, 0, 0);
                        if (selection >= 0) {
                            const LRESULT value = SendMessageW(templateCombo, CB_GETITEMDATA, selection, 0);
                            if (value != CB_ERR) {
                                data->workingOptions.newTabTemplate = static_cast<NewTabTemplate>(value);
                            }
                        }
                    }

                    data->workingOptions.newTabCustomPath =
                        Trim(GetWindowTextString(GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_PATH_EDIT)));

                    if (HWND groupCombo = GetDlgItem(hwnd, IDC_MAIN_NEW_TAB_GROUP_COMBO)) {
                        if (data->workingGroups.empty()) {
                            data->workingOptions.newTabSavedGroup.clear();
                        } else {
                            data->workingOptions.newTabSavedGroup =
                                Trim(GetWindowTextString(groupCombo));
                        }
                    }

                    HWND dockCombo = GetDlgItem(hwnd, IDC_MAIN_DOCK_COMBO);
                    if (dockCombo) {
                        const LRESULT selection = SendMessageW(dockCombo, CB_GETCURSEL, 0, 0);
                        if (selection >= 0) {
                            const LRESULT value = SendMessageW(dockCombo, CB_GETITEMDATA, selection, 0);
                            if (value != CB_ERR) {
                                data->workingOptions.tabDockMode =
                                    static_cast<TabBandDockMode>(value);
                            }
                        }
                    }
                    data->applyInvoked = true;
                }
                SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, PSNRET_NOERROR);
                return TRUE;
            }
            break;
        }
    }
    return FALSE;
}

class CustomizationsPageController {
public:
    static void Initialize(HWND hwnd, OptionsDialogData* data) {
        if (!data) {
            return;
        }
        CheckDlgButton(hwnd, IDC_MAIN_BREADCRUMB,
                       data->workingOptions.enableBreadcrumbGradient ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, IDC_MAIN_BREADCRUMB_FONT,
                       data->workingOptions.enableBreadcrumbFontGradient ? BST_CHECKED : BST_UNCHECKED);
        ConfigurePercentageSlider(hwnd, IDC_MAIN_BREADCRUMB_BG_SLIDER,
                                  data->workingOptions.breadcrumbGradientTransparency);
        ConfigurePercentageSlider(hwnd, IDC_MAIN_BREADCRUMB_FONT_SLIDER,
                                  InvertPercentageValue(data->workingOptions.breadcrumbFontBrightness));
        ConfigureMultiplierSlider(hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER,
                                  data->workingOptions.breadcrumbHighlightAlphaMultiplier);
        ConfigureMultiplierSlider(hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER,
                                  data->workingOptions.breadcrumbDropdownAlphaMultiplier);
        UpdatePercentageLabel(hwnd, IDC_MAIN_BREADCRUMB_BG_VALUE,
                              data->workingOptions.breadcrumbGradientTransparency);
        UpdatePercentageLabel(hwnd, IDC_MAIN_BREADCRUMB_FONT_VALUE,
                              data->workingOptions.breadcrumbFontBrightness);
        UpdateMultiplierLabel(hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_VALUE,
                              data->workingOptions.breadcrumbHighlightAlphaMultiplier);
        UpdateMultiplierLabel(hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_VALUE,
                              data->workingOptions.breadcrumbDropdownAlphaMultiplier);
        CheckDlgButton(hwnd, IDC_MAIN_BREADCRUMB_BG_CUSTOM,
                       data->workingOptions.useCustomBreadcrumbGradientColors ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, IDC_MAIN_BREADCRUMB_FONT_CUSTOM,
                       data->workingOptions.useCustomBreadcrumbFontColors ? BST_CHECKED : BST_UNCHECKED);
        UpdateGradientStates(hwnd);
        SetPreviewColor(hwnd, IDC_MAIN_BREADCRUMB_BG_START_PREVIEW, &data->breadcrumbBgStartBrush,
                        data->workingOptions.breadcrumbGradientStartColor);
        SetPreviewColor(hwnd, IDC_MAIN_BREADCRUMB_BG_END_PREVIEW, &data->breadcrumbBgEndBrush,
                        data->workingOptions.breadcrumbGradientEndColor);
        SetPreviewColor(hwnd, IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW, &data->breadcrumbFontStartBrush,
                        data->workingOptions.breadcrumbFontGradientStartColor);
        SetPreviewColor(hwnd, IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW, &data->breadcrumbFontEndBrush,
                        data->workingOptions.breadcrumbFontGradientEndColor);
        CheckDlgButton(hwnd, IDC_MAIN_PROGRESS_CUSTOM,
                       data->workingOptions.useCustomProgressBarGradientColors ? BST_CHECKED : BST_UNCHECKED);
        SetPreviewColor(hwnd, IDC_MAIN_PROGRESS_START_PREVIEW, &data->progressStartBrush,
                        data->workingOptions.progressBarGradientStartColor);
        SetPreviewColor(hwnd, IDC_MAIN_PROGRESS_END_PREVIEW, &data->progressEndBrush,
                        data->workingOptions.progressBarGradientEndColor);
        UpdateProgressColorControlsEnabled(hwnd, data->workingOptions.useCustomProgressBarGradientColors);
        CheckDlgButton(hwnd, IDC_MAIN_TAB_SELECTED_CHECK,
                       data->workingOptions.useCustomTabSelectedColor ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, IDC_MAIN_TAB_UNSELECTED_CHECK,
                       data->workingOptions.useCustomTabUnselectedColor ? BST_CHECKED : BST_UNCHECKED);
        SetPreviewColor(hwnd, IDC_MAIN_TAB_SELECTED_PREVIEW, &data->tabSelectedBrush,
                        data->workingOptions.customTabSelectedColor);
        SetPreviewColor(hwnd, IDC_MAIN_TAB_UNSELECTED_PREVIEW, &data->tabUnselectedBrush,
                        data->workingOptions.customTabUnselectedColor);
        UpdateTabColorStates(hwnd);
        CheckDlgButton(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE,
                       data->workingOptions.enableFolderBackgrounds ? BST_CHECKED : BST_UNCHECKED);
        HWND backgroundList = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
        InitializeFolderBackgroundList(backgroundList);
        RefreshFolderBackgroundListView(backgroundList, data);
        if (backgroundList && !data->workingOptions.folderBackgroundEntries.empty()) {
            ListView_SetItemState(backgroundList, 0, LVIS_SELECTED | LVIS_FOCUSED,
                                  LVIS_SELECTED | LVIS_FOCUSED);
        }
        UpdateUniversalBackgroundPreview(hwnd, data);
        UpdateSelectedFolderBackgroundPreview(hwnd, data);
        UpdateFolderBackgroundControlsEnabled(hwnd, data->workingOptions.enableFolderBackgrounds);
        UpdateFolderBackgroundButtons(hwnd);
        data->lastFolderBrowsePath =
            data->workingOptions.folderBackgroundEntries.empty()
                ? std::wstring{}
                : data->workingOptions.folderBackgroundEntries.front().folderPath;
        data->lastImageBrowseDirectory =
            ExtractDirectoryFromPath(data->workingOptions.universalFolderBackgroundImage.cachedImagePath);
        CaptureCustomizationChildPlacements(hwnd, data);
        UpdateCustomizationScrollInfo(hwnd, data);
    }

    static bool HandleCommand(HWND hwnd, WPARAM wParam, LPARAM) {
        const WORD controlId = LOWORD(wParam);
        const WORD notification = HIWORD(wParam);
        if (notification != BN_CLICKED) {
            return false;
        }

        auto* data = GetDialogData(hwnd);

        if (std::find(kGradientToggleIds.begin(), kGradientToggleIds.end(), controlId) !=
            kGradientToggleIds.end()) {
            UpdateGradientStates(hwnd);
            NotifyParentOfChange(hwnd);
            return true;
        }

        if (std::find(kTabToggleIds.begin(), kTabToggleIds.end(), controlId) != kTabToggleIds.end()) {
            UpdateTabColorStates(hwnd);
            NotifyParentOfChange(hwnd);
            return true;
        }

        if (std::find(kColorButtonIds.begin(), kColorButtonIds.end(), controlId) != kColorButtonIds.end()) {
            if (HandleColorSelection(hwnd, controlId, data)) {
                NotifyParentOfChange(hwnd);
            }
            return true;
        }

        if (controlId == IDC_CUSTOM_BACKGROUND_ENABLE) {
            const bool enabled = IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED;
            UpdateFolderBackgroundControlsEnabled(hwnd, enabled);
            UpdateFolderBackgroundButtons(hwnd);
            NotifyParentOfChange(hwnd);
            return true;
        }

        return HandleFolderBackgroundCommand(hwnd, controlId, data);
    }

    static void HandleHScroll(HWND hwnd, WPARAM, LPARAM lParam) {
        HWND slider = reinterpret_cast<HWND>(lParam);
        if (!slider) {
            return;
        }

        auto* data = GetDialogData(hwnd);
        bool previewNeeded = false;
        if (HandleSliderChange(hwnd, slider, data, &previewNeeded)) {
            if (previewNeeded) {
                ApplyCustomizationPreview(hwnd, data);
            }
            NotifyParentOfChange(hwnd);
        }
    }

    static bool HandleNotify(HWND hwnd, WPARAM, LPARAM lParam) {
        auto* header = reinterpret_cast<LPNMHDR>(lParam);
        if (!header) {
            return false;
        }

        if (header->idFrom == IDC_CUSTOM_BACKGROUND_LIST) {
            return HandleListViewNotify(hwnd, header, GetDialogData(hwnd));
        }

        if (header->code == PSN_APPLY) {
            return Apply(hwnd, GetDialogData(hwnd));
        }

        return false;
    }

private:
    using SliderClampFn = int (*)(int);
    using SliderTransformFn = int (*)(int);
    using LabelUpdateFn = void (*)(HWND, int, int);
    using SliderApplyFn = bool (*)(OptionsDialogData*, int);

    struct SliderBinding {
        WORD sliderId;
        WORD labelId;
        SliderClampFn clamp;
        SliderTransformFn transform;
        LabelUpdateFn updateLabel;
        SliderApplyFn apply;
    };

    static OptionsDialogData* GetDialogData(HWND hwnd) {
        return reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
    }

    static void NotifyParentOfChange(HWND hwnd) {
        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
    }

    static void UpdateGradientStates(HWND hwnd) {
        const bool backgroundEnabled = IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB) == BST_CHECKED;
        const bool fontEnabled = IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT) == BST_CHECKED;
        UpdateGradientControlsEnabled(hwnd, backgroundEnabled, fontEnabled);

        const bool bgCustom = IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_BG_CUSTOM) == BST_CHECKED;
        const bool fontCustom = IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT_CUSTOM) == BST_CHECKED;
        UpdateGradientColorControlsEnabled(hwnd, backgroundEnabled && bgCustom, fontEnabled && fontCustom);

        const bool progressCustom = IsDlgButtonChecked(hwnd, IDC_MAIN_PROGRESS_CUSTOM) == BST_CHECKED;
        UpdateProgressColorControlsEnabled(hwnd, progressCustom);
    }

    static void UpdateTabColorStates(HWND hwnd) {
        const bool tabSelectedCustom = IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_SELECTED_CHECK) == BST_CHECKED;
        const bool tabUnselectedCustom = IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_UNSELECTED_CHECK) == BST_CHECKED;
        UpdateTabColorControlsEnabled(hwnd, tabSelectedCustom, tabUnselectedCustom);
    }

    static bool HandleColorSelection(HWND hwnd, WORD controlId, OptionsDialogData* data) {
        return HandleColorButtonClick(hwnd, data, controlId);
    }

    static bool HandleFolderBackgroundCommand(HWND hwnd, WORD controlId, OptionsDialogData* data) {
        switch (controlId) {
            case IDC_CUSTOM_BACKGROUND_BROWSE:
                HandleUniversalBackgroundBrowse(hwnd, data);
                return true;
            case IDC_CUSTOM_BACKGROUND_ADD:
                if (data && IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                    HandleAddFolderBackgroundEntry(hwnd, data);
                }
                return true;
            case IDC_CUSTOM_BACKGROUND_EDIT:
                if (data && IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                    HandleEditFolderBackgroundEntry(hwnd, data);
                }
                return true;
            case IDC_CUSTOM_BACKGROUND_REMOVE:
                if (data && IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                    HandleRemoveFolderBackgroundEntry(hwnd, data);
                }
                return true;
            case IDC_CUSTOM_BACKGROUND_CLEAN:
                if (data && IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                    HandleBackgroundCacheMaintenance(hwnd, data);
                }
                return true;
            default:
                return false;
        }
    }

    static bool HandleSliderChange(HWND hwnd, HWND slider, OptionsDialogData* data, bool* previewNeeded) {
        if (!slider || !previewNeeded) {
            return false;
        }
        const int controlId = GetDlgCtrlID(slider);
        const auto it =
            std::find_if(kSliderBindings.begin(), kSliderBindings.end(), [controlId](const SliderBinding& binding) {
                return binding.sliderId == controlId;
            });
        if (it == kSliderBindings.end()) {
            return false;
        }

        int position = static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0));
        if (it->clamp) {
            position = it->clamp(position);
        }
        const int value = it->transform ? it->transform(position) : position;
        it->updateLabel(hwnd, it->labelId, value);
        const bool changed = it->apply(data, value);
        *previewNeeded = changed;
        return true;
    }

    static bool HandleListViewNotify(HWND hwnd, LPNMHDR header, OptionsDialogData* data) {
        switch (header->code) {
            case LVN_ITEMCHANGED: {
                auto* changed = reinterpret_cast<NMLISTVIEW*>(header);
                if (changed && (changed->uChanged & LVIF_STATE) != 0) {
                    UpdateFolderBackgroundButtons(hwnd);
                    UpdateSelectedFolderBackgroundPreview(hwnd, data);
                }
                return true;
            }
            case NM_DBLCLK:
                if (data && IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                    HandleEditFolderBackgroundEntry(hwnd, data);
                }
                return true;
            default:
                return false;
        }
    }

    static bool Apply(HWND hwnd, OptionsDialogData* data) {
        if (data) {
            data->workingOptions.enableBreadcrumbGradient =
                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB) == BST_CHECKED;
            data->workingOptions.enableBreadcrumbFontGradient =
                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT) == BST_CHECKED;
            data->workingOptions.breadcrumbGradientTransparency =
                ClampPercentageValue(static_cast<int>(SendDlgItemMessageW(
                    hwnd, IDC_MAIN_BREADCRUMB_BG_SLIDER, TBM_GETPOS, 0, 0)));
            const int brightnessSliderValue = ClampPercentageValue(static_cast<int>(SendDlgItemMessageW(
                hwnd, IDC_MAIN_BREADCRUMB_FONT_SLIDER, TBM_GETPOS, 0, 0)));
            data->workingOptions.breadcrumbFontBrightness = InvertPercentageValue(brightnessSliderValue);
            data->workingOptions.breadcrumbHighlightAlphaMultiplier =
                ClampMultiplierValue(static_cast<int>(SendDlgItemMessageW(
                    hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER, TBM_GETPOS, 0, 0)));
            data->workingOptions.breadcrumbDropdownAlphaMultiplier =
                ClampMultiplierValue(static_cast<int>(SendDlgItemMessageW(
                    hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER, TBM_GETPOS, 0, 0)));
            data->workingOptions.useCustomBreadcrumbGradientColors =
                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_BG_CUSTOM) == BST_CHECKED;
            data->workingOptions.useCustomBreadcrumbFontColors =
                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT_CUSTOM) == BST_CHECKED;
            data->workingOptions.useCustomProgressBarGradientColors =
                IsDlgButtonChecked(hwnd, IDC_MAIN_PROGRESS_CUSTOM) == BST_CHECKED;
            data->workingOptions.useCustomTabSelectedColor =
                IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_SELECTED_CHECK) == BST_CHECKED;
            data->workingOptions.useCustomTabUnselectedColor =
                IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_UNSELECTED_CHECK) == BST_CHECKED;
            data->workingOptions.enableFolderBackgrounds =
                IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED;
            data->applyInvoked = true;
        }
        SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, PSNRET_NOERROR);
        return true;
    }

    static bool ApplyBreadcrumbTransparency(OptionsDialogData* data, int value) {
        if (!data) {
            return false;
        }
        if (data->workingOptions.breadcrumbGradientTransparency == value) {
            return false;
        }
        data->workingOptions.breadcrumbGradientTransparency = value;
        return true;
    }

    static bool ApplyBreadcrumbFontBrightness(OptionsDialogData* data, int value) {
        if (!data) {
            return false;
        }
        if (data->workingOptions.breadcrumbFontBrightness == value) {
            return false;
        }
        data->workingOptions.breadcrumbFontBrightness = value;
        return true;
    }

    static bool ApplyBreadcrumbHighlightMultiplier(OptionsDialogData* data, int value) {
        if (!data) {
            return false;
        }
        if (data->workingOptions.breadcrumbHighlightAlphaMultiplier == value) {
            return false;
        }
        data->workingOptions.breadcrumbHighlightAlphaMultiplier = value;
        return true;
    }

    static bool ApplyBreadcrumbDropdownMultiplier(OptionsDialogData* data, int value) {
        if (!data) {
            return false;
        }
        if (data->workingOptions.breadcrumbDropdownAlphaMultiplier == value) {
            return false;
        }
        data->workingOptions.breadcrumbDropdownAlphaMultiplier = value;
        return true;
    }

    static inline constexpr std::array<WORD, 5> kGradientToggleIds = {
        IDC_MAIN_BREADCRUMB,        IDC_MAIN_BREADCRUMB_FONT,        IDC_MAIN_BREADCRUMB_BG_CUSTOM,
        IDC_MAIN_BREADCRUMB_FONT_CUSTOM, IDC_MAIN_PROGRESS_CUSTOM};

    static inline constexpr std::array<WORD, 2> kTabToggleIds = {
        IDC_MAIN_TAB_SELECTED_CHECK, IDC_MAIN_TAB_UNSELECTED_CHECK};

    static inline constexpr std::array<WORD, 8> kColorButtonIds = {
        IDC_MAIN_BREADCRUMB_BG_START_BUTTON,  IDC_MAIN_BREADCRUMB_BG_END_BUTTON,
        IDC_MAIN_BREADCRUMB_FONT_START_BUTTON, IDC_MAIN_BREADCRUMB_FONT_END_BUTTON,
        IDC_MAIN_PROGRESS_START_BUTTON,        IDC_MAIN_PROGRESS_END_BUTTON,
        IDC_MAIN_TAB_SELECTED_BUTTON,          IDC_MAIN_TAB_UNSELECTED_BUTTON};

    static inline constexpr std::array<SliderBinding, 4> kSliderBindings = {{
        {IDC_MAIN_BREADCRUMB_BG_SLIDER, IDC_MAIN_BREADCRUMB_BG_VALUE, &ClampPercentageValue, nullptr,
         &UpdatePercentageLabel, &ApplyBreadcrumbTransparency},
        {IDC_MAIN_BREADCRUMB_FONT_SLIDER, IDC_MAIN_BREADCRUMB_FONT_VALUE, &ClampPercentageValue,
         &InvertPercentageValue, &UpdatePercentageLabel, &ApplyBreadcrumbFontBrightness},
        {IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER, IDC_MAIN_BREADCRUMB_HIGHLIGHT_VALUE, &ClampMultiplierValue,
         nullptr, &UpdateMultiplierLabel, &ApplyBreadcrumbHighlightMultiplier},
        {IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER, IDC_MAIN_BREADCRUMB_DROPDOWN_VALUE, &ClampMultiplierValue,
         nullptr, &UpdateMultiplierLabel, &ApplyBreadcrumbDropdownMultiplier},
    }};
};

INT_PTR CALLBACK CustomizationsPageProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* data = reinterpret_cast<OptionsDialogData*>(reinterpret_cast<PROPSHEETPAGEW*>(lParam)->lParam);
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(data));
            CustomizationsPageController::Initialize(hwnd, data);
            return TRUE;
        }
        case WM_CTLCOLORDLG: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (dc) {
                SetBkColor(dc, GetSysColor(COLOR_3DFACE));
            }
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
        }
        case WM_COMMAND: {
            if (CustomizationsPageController::HandleCommand(hwnd, wParam, lParam)) {
                return TRUE;
            }
            break;
        }
        case WM_PREVIEW_BITMAP_READY: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            auto* result = reinterpret_cast<PreviewBitmapResult*>(lParam);
            if (!result) {
                return TRUE;
            }

            const int controlId = static_cast<int>(wParam);
            bool applied = false;
            if (data) {
                if (controlId == IDC_CUSTOM_BACKGROUND_PREVIEW && result->token == data->universalPreviewToken) {
                    SetPreviewBitmap(hwnd, IDC_CUSTOM_BACKGROUND_PREVIEW, &data->universalBackgroundPreview,
                                     result->bitmap);
                    result->bitmap = nullptr;
                    applied = true;
                } else if (controlId == IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW &&
                           result->token == data->folderPreviewToken) {
                    SetPreviewBitmap(hwnd, IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW, &data->folderBackgroundPreview,
                                     result->bitmap);
                    result->bitmap = nullptr;
                    applied = true;
                }
            }

            if (!applied && result->bitmap) {
                DeleteObject(result->bitmap);
            }
            delete result;
            return TRUE;
        }
        case WM_CTLCOLORSTATIC: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!data) {
                break;
            }
            HWND target = reinterpret_cast<HWND>(lParam);
            if (!target) {
                break;
            }
            HDC dc = reinterpret_cast<HDC>(wParam);
            const int controlId = GetDlgCtrlID(target);
            HBRUSH brush = nullptr;
            COLORREF color = 0;
            switch (controlId) {
                case IDC_MAIN_BREADCRUMB_BG_START_PREVIEW:
                    brush = data->breadcrumbBgStartBrush;
                    color = data->workingOptions.breadcrumbGradientStartColor;
                    break;
                case IDC_MAIN_BREADCRUMB_BG_END_PREVIEW:
                    brush = data->breadcrumbBgEndBrush;
                    color = data->workingOptions.breadcrumbGradientEndColor;
                    break;
                case IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW:
                    brush = data->breadcrumbFontStartBrush;
                    color = data->workingOptions.breadcrumbFontGradientStartColor;
                    break;
                case IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW:
                    brush = data->breadcrumbFontEndBrush;
                    color = data->workingOptions.breadcrumbFontGradientEndColor;
                    break;
                case IDC_MAIN_PROGRESS_START_PREVIEW:
                    brush = data->progressStartBrush;
                    color = data->workingOptions.progressBarGradientStartColor;
                    break;
                case IDC_MAIN_PROGRESS_END_PREVIEW:
                    brush = data->progressEndBrush;
                    color = data->workingOptions.progressBarGradientEndColor;
                    break;
                case IDC_MAIN_TAB_SELECTED_PREVIEW:
                    brush = data->tabSelectedBrush;
                    color = data->workingOptions.customTabSelectedColor;
                    break;
                case IDC_MAIN_TAB_UNSELECTED_PREVIEW:
                    brush = data->tabUnselectedBrush;
                    color = data->workingOptions.customTabUnselectedColor;
                    break;
                default:
                    break;
            }
            if (brush) {
                SetBkMode(dc, OPAQUE);
                SetBkColor(dc, color);
                return reinterpret_cast<INT_PTR>(brush);
            }
            wchar_t className[32];
            if (GetClassNameW(target, className, ARRAYSIZE(className))) {
                if (_wcsicmp(className, L"Button") == 0) {
                    const LONG style = GetWindowLongW(target, GWL_STYLE);
                    if ((style & BS_GROUPBOX) == BS_GROUPBOX) {
                        SetBkMode(dc, TRANSPARENT);
                        SetBkColor(dc, GetSysColor(COLOR_3DFACE));
                        return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
                    }
                }
            }
            SetBkMode(dc, TRANSPARENT);
            SetBkColor(dc, GetSysColor(COLOR_3DFACE));
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
            break;
        }
        case WM_HSCROLL:
            CustomizationsPageController::HandleHScroll(hwnd, wParam, lParam);
            return TRUE;
        case WM_SIZE: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            UpdateCustomizationScrollInfo(hwnd, data);
            return TRUE;
        }
        case WM_VSCROLL: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!data) {
                return TRUE;
            }
            bool handled = false;
            switch (LOWORD(wParam)) {
                case SB_LINEUP:
                    handled = ApplyCustomizationScrollDelta(hwnd, data, -kCustomizationScrollLineStep);
                    break;
                case SB_LINEDOWN:
                    handled = ApplyCustomizationScrollDelta(hwnd, data, kCustomizationScrollLineStep);
                    break;
                case SB_PAGEUP:
                    handled = ApplyCustomizationScrollDelta(hwnd, data, -kCustomizationScrollPageStep);
                    break;
                case SB_PAGEDOWN:
                    handled = ApplyCustomizationScrollDelta(hwnd, data, kCustomizationScrollPageStep);
                    break;
                case SB_TOP:
                    handled = UpdateCustomizationScrollPosition(hwnd, data, 0);
                    break;
                case SB_BOTTOM:
                    handled = UpdateCustomizationScrollPosition(hwnd, data, data->customizationScrollMax);
                    break;
                case SB_THUMBTRACK:
                case SB_THUMBPOSITION: {
                    SCROLLINFO info{sizeof(info)};
                    info.fMask = SIF_TRACKPOS;
                    if (GetScrollInfo(hwnd, SB_VERT, &info)) {
                        handled = UpdateCustomizationScrollPosition(hwnd, data, info.nTrackPos);
                    }
                    break;
                }
                default:
                    break;
            }
            if (handled) {
                data->customizationWheelRemainder = 0;
            }
            return TRUE;
        }
        case WM_MOUSEWHEEL: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!data) {
                return TRUE;
            }
            data->customizationWheelRemainder += GET_WHEEL_DELTA_WPARAM(wParam);
            const int increment = data->customizationWheelRemainder / WHEEL_DELTA;
            if (increment != 0) {
                data->customizationWheelRemainder -= increment * WHEEL_DELTA;
                UINT wheelLines = 3;
                if (!SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &wheelLines, 0)) {
                    wheelLines = 3;
                }
                if (wheelLines == WHEEL_PAGESCROLL) {
                    RECT client{};
                    if (GetClientRect(hwnd, &client)) {
                        const int clientHeight = static_cast<int>(client.bottom - client.top);
                        const int page = std::max(1, clientHeight - kCustomizationScrollLineStep);
                        ApplyCustomizationScrollDelta(hwnd, data, -increment * page);
                    } else {
                        ApplyCustomizationScrollDelta(hwnd, data, -increment * kCustomizationScrollPageStep);
                    }
                } else if (wheelLines > 0) {
                    const int delta = static_cast<int>(wheelLines) * kCustomizationScrollLineStep;
                    ApplyCustomizationScrollDelta(hwnd, data, -increment * delta);
                }
            }
            return TRUE;
        }
        case WM_NOTIFY:
            if (CustomizationsPageController::HandleNotify(hwnd, wParam, lParam)) {
                return TRUE;
            }
            break;
    }
    return FALSE;
}

INT_PTR CALLBACK GlowPageProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* page = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
            auto* data = page ? reinterpret_cast<OptionsDialogData*>(page->lParam) : nullptr;
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(data));
            RefreshGlowControls(hwnd, data);
            return TRUE;
        }
        case WM_CTLCOLORDLG: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (dc) {
                SetBkColor(dc, GetSysColor(COLOR_3DFACE));
            }
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
        }
        case WM_CTLCOLORSTATIC: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!data) {
                break;
            }
            HWND target = reinterpret_cast<HWND>(lParam);
            if (!target) {
                break;
            }
            HDC dc = reinterpret_cast<HDC>(wParam);
            const int controlId = GetDlgCtrlID(target);
            if (controlId == IDC_GLOW_PRIMARY_PREVIEW && data->glowPrimaryBrush) {
                SetBkMode(dc, OPAQUE);
                SetBkColor(dc, data->workingOptions.neonGlowPrimaryColor);
                return reinterpret_cast<INT_PTR>(data->glowPrimaryBrush);
            }
            if (controlId == IDC_GLOW_SECONDARY_PREVIEW && data->glowSecondaryBrush) {
                SetBkMode(dc, OPAQUE);
                SetBkColor(dc, data->workingOptions.neonGlowSecondaryColor);
                return reinterpret_cast<INT_PTR>(data->glowSecondaryBrush);
            }
            wchar_t className[32];
            if (GetClassNameW(target, className, ARRAYSIZE(className))) {
                if (_wcsicmp(className, L"Button") == 0) {
                    const LONG style = GetWindowLongW(target, GWL_STYLE);
                    if ((style & BS_GROUPBOX) == BS_GROUPBOX) {
                        SetBkMode(dc, TRANSPARENT);
                        SetBkColor(dc, GetSysColor(COLOR_3DFACE));
                        return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
                    }
                }
            }
            SetBkMode(dc, TRANSPARENT);
            SetBkColor(dc, GetSysColor(COLOR_3DFACE));
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_3DFACE));
        }
        case WM_COMMAND: {
            if (HIWORD(wParam) != BN_CLICKED) {
                break;
            }
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            const WORD controlId = LOWORD(wParam);
            for (const auto& mapping : kGlowSurfaceControlMappings) {
                if (controlId == mapping.controlId) {
                    if (data) {
                        const bool enabled =
                            IsDlgButtonChecked(hwnd, controlId) == BST_CHECKED;
                        GlowSurfaceOptions& surface =
                            data->workingOptions.glowPalette.*(mapping.member);
                        if (surface.enabled != enabled) {
                            surface.enabled = enabled;
                            SendMessageW(GetParent(hwnd), PSM_CHANGED,
                                         reinterpret_cast<WPARAM>(hwnd), 0);
                            ApplyCustomizationPreview(hwnd, data);
                        }
                    }
                    return TRUE;
                }
            }
            switch (controlId) {
                case IDC_GLOW_ENABLE:
                    if (data) {
                        data->workingOptions.enableNeonGlow =
                            IsDlgButtonChecked(hwnd, IDC_GLOW_ENABLE) == BST_CHECKED;
                        UpdateGlowControlStates(hwnd);
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                        ApplyCustomizationPreview(hwnd, data);
                    }
                    return TRUE;
                case IDC_GLOW_CUSTOM_COLORS:
                    if (data) {
                        data->workingOptions.useCustomNeonGlowColors =
                            IsDlgButtonChecked(hwnd, IDC_GLOW_CUSTOM_COLORS) == BST_CHECKED;
                        UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                        UpdateGlowControlStates(hwnd);
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                        ApplyCustomizationPreview(hwnd, data);
                    }
                    return TRUE;
                case IDC_GLOW_USE_GRADIENT:
                    if (data) {
                        data->workingOptions.useNeonGlowGradient =
                            IsDlgButtonChecked(hwnd, IDC_GLOW_USE_GRADIENT) == BST_CHECKED;
                        UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                        UpdateGlowControlStates(hwnd);
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                        ApplyCustomizationPreview(hwnd, data);
                    }
                    return TRUE;
                case IDC_GLOW_PRIMARY_BUTTON:
                case IDC_GLOW_SECONDARY_BUTTON:
                    if (data && HandleColorButtonClick(hwnd, data, controlId)) {
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                        ApplyCustomizationPreview(hwnd, data);
                    }
                    return TRUE;
                default:
                    break;
            }
            break;
        }
        case WM_NOTIFY: {
            auto* header = reinterpret_cast<LPNMHDR>(lParam);
            if (!header) {
                break;
            }
            if (header->code == PSN_SETACTIVE) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                RefreshGlowControls(hwnd, data);
                SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, 0);
                return TRUE;
            }
            if (header->code == PSN_APPLY) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                if (data) {
                    data->workingOptions.enableNeonGlow =
                        IsDlgButtonChecked(hwnd, IDC_GLOW_ENABLE) == BST_CHECKED;
                    data->workingOptions.useCustomNeonGlowColors =
                        IsDlgButtonChecked(hwnd, IDC_GLOW_CUSTOM_COLORS) == BST_CHECKED;
                    data->workingOptions.useNeonGlowGradient =
                        IsDlgButtonChecked(hwnd, IDC_GLOW_USE_GRADIENT) == BST_CHECKED;
                    for (const auto& mapping : kGlowSurfaceControlMappings) {
                        GlowSurfaceOptions& surface =
                            data->workingOptions.glowPalette.*(mapping.member);
                        surface.enabled =
                            IsDlgButtonChecked(hwnd, mapping.controlId) == BST_CHECKED;
                    }
                    UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                    data->applyInvoked = true;
                }
                SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, PSNRET_NOERROR);
                return TRUE;
            }
            break;
        }
    }
    return FALSE;
}

INT_PTR CALLBACK ContextMenuPageProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* page = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
            auto* data = page ? reinterpret_cast<OptionsDialogData*>(page->lParam) : nullptr;
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(data));
            PopulateContextMenuAnchorCombo(GetDlgItem(hwnd, IDC_CONTEXT_ANCHOR_COMBO));
            if (HWND hints = GetDlgItem(hwnd, IDC_CONTEXT_HINTS_STATIC)) {
                wchar_t buffer[256] = {};
                if (LoadStringW(GetModuleHandleInstance(), IDS_OPTIONS_COMMAND_HINTS, buffer,
                                ARRAYSIZE(buffer)) > 0) {
                    SetWindowTextW(hints, buffer);
                }
            }
            RefreshContextMenuTree(hwnd, data, nullptr);
            PopulateContextMenuDetailControls(hwnd, data);
            UpdateContextMenuButtonStates(hwnd, data);
            return TRUE;
        }
        case WM_COMMAND: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!data) {
                break;
            }
            const WORD controlId = LOWORD(wParam);
            const WORD notifyCode = HIWORD(wParam);
            switch (controlId) {
                case IDC_CONTEXT_ADD_COMMAND:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuAddItem(hwnd, data, ContextMenuItemType::kCommand);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_ADD_SUBMENU:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuAddItem(hwnd, data, ContextMenuItemType::kSubmenu);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_ADD_SEPARATOR:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuAddItem(hwnd, data, ContextMenuItemType::kSeparator);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_REMOVE:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuRemoveItem(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_MOVE_UP:
                    if (notifyCode == BN_CLICKED) {
                        MoveContextMenuItem(hwnd, data, true);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_MOVE_DOWN:
                    if (notifyCode == BN_CLICKED) {
                        MoveContextMenuItem(hwnd, data, false);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_INDENT:
                    if (notifyCode == BN_CLICKED) {
                        IndentContextMenuItem(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_OUTDENT:
                    if (notifyCode == BN_CLICKED) {
                        OutdentContextMenuItem(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_ICON_BROWSE:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuBrowseIcon(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_COMMAND_BROWSE:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuBrowseCommand(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_SCOPE_FILES:
                case IDC_CONTEXT_SCOPE_FOLDERS:
                    if (notifyCode == BN_CLICKED) {
                        ApplyContextMenuDetailsFromControls(hwnd, data, true);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_SEPARATOR_CHECK:
                    if (notifyCode == BN_CLICKED) {
                        const bool ensureSeparator =
                            Button_GetCheck(GetDlgItem(hwnd, IDC_CONTEXT_SEPARATOR_CHECK)) == BST_CHECKED;
                        ToggleSeparatorAbove(hwnd, data, ensureSeparator);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_EXTENSION_ADD:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuExtensionAdd(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_EXTENSION_REMOVE:
                    if (notifyCode == BN_CLICKED) {
                        HandleContextMenuExtensionRemove(hwnd, data);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_EXTENSION_LIST:
                    if (notifyCode == LBN_SELCHANGE) {
                        HWND removeButton = GetDlgItem(hwnd, IDC_CONTEXT_EXTENSION_REMOVE);
                        if (removeButton) {
                            HWND list = reinterpret_cast<HWND>(lParam);
                            const int selection =
                                static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
                            EnableWindow(removeButton, selection >= 0);
                        }
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_ANCHOR_COMBO:
                    if (notifyCode == CBN_SELCHANGE) {
                        ApplyContextMenuDetailsFromControls(hwnd, data, true);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_LABEL_EDIT:
                case IDC_CONTEXT_ICON_EDIT:
                case IDC_CONTEXT_COMMAND_PATH:
                case IDC_CONTEXT_COMMAND_ARGS:
                case IDC_CONTEXT_SELECTION_MIN:
                case IDC_CONTEXT_SELECTION_MAX:
                    if (notifyCode == EN_CHANGE) {
                        ApplyContextMenuDetailsFromControls(hwnd, data, true);
                        return TRUE;
                    }
                    break;
                case IDC_CONTEXT_EXTENSION_EDIT:
                    if (notifyCode == EN_CHANGE) {
                        // no-op; handled on add
                        return TRUE;
                    }
                    break;
                default:
                    break;
            }
            break;
        }
        case WM_NOTIFY: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (!data) {
                break;
            }
            auto* header = reinterpret_cast<LPNMHDR>(lParam);
            if (!header) {
                break;
            }
            if (header->idFrom == IDC_CONTEXT_TREE) {
                switch (header->code) {
                    case TVN_SELCHANGING:
                        ApplyContextMenuDetailsFromControls(hwnd, data, true);
                        break;
                    case TVN_SELCHANGED:
                        PopulateContextMenuDetailControls(hwnd, data);
                        UpdateContextMenuButtonStates(hwnd, data);
                        break;
                    default:
                        break;
                }
                return TRUE;
            }
            if (header->code == PSN_APPLY) {
                ApplyContextMenuDetailsFromControls(hwnd, data, true);
                ContextMenuValidationError error;
                std::vector<size_t> validationPath;
                if (!ValidateContextMenuItems(data->workingOptions.contextMenuItems, &validationPath, &error)) {
                    if (!error.message.empty()) {
                        MessageBoxW(hwnd, error.message.c_str(), L"ShellTabs", MB_OK | MB_ICONWARNING);
                    }
                    if (!error.path.empty()) {
                        data->contextSelectionPath = error.path;
                        data->contextSelectionValid = true;
                        RefreshContextMenuTree(hwnd, data, &data->contextSelectionPath);
                        PopulateContextMenuDetailControls(hwnd, data);
                        UpdateContextMenuButtonStates(hwnd, data);
                    }
                    SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, PSNRET_INVALID_NOCHANGEPAGE);
                    return TRUE;
                }
                data->applyInvoked = true;
                SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, PSNRET_NOERROR);
                return TRUE;
            }
            break;
        }
    }
    return FALSE;
}

INT_PTR CALLBACK GroupManagementPageProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* data = reinterpret_cast<OptionsDialogData*>(reinterpret_cast<PROPSHEETPAGEW*>(lParam)->lParam);
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(data));
            RefreshGroupList(GetDlgItem(hwnd, IDC_GROUP_LIST), data);
            UpdateGroupButtons(hwnd);
            ApplyFocusedGroupSelection(hwnd, data);
            return TRUE;
        }
        case WM_COMMAND: {
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            switch (LOWORD(wParam)) {
                case IDC_GROUP_NEW:
                    HandleNewGroup(hwnd, data);
                    return TRUE;
                case IDC_GROUP_EDIT:
                    HandleEditGroup(hwnd, data);
                    return TRUE;
                case IDC_GROUP_REMOVE:
                    HandleRemoveGroup(hwnd, data);
                    return TRUE;
                case IDC_GROUP_LIST:
                    if (HIWORD(wParam) == LBN_SELCHANGE) {
                        UpdateGroupButtons(hwnd);
                    } else if (HIWORD(wParam) == LBN_DBLCLK) {
                        HandleEditGroup(hwnd, data);
                    }
                    return TRUE;
            }
            break;
        }
        case WM_NOTIFY: {
            if (((LPNMHDR)lParam)->code == PSN_APPLY) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                if (data) {
                    data->applyInvoked = true;
                }
                SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, PSNRET_NOERROR);
                return TRUE;
            }
            break;
        }
    }
    return FALSE;
}

int CALLBACK OptionsSheetCallback(HWND hwnd, UINT message, LPARAM) {
    if (message == PSCB_INITIALIZED) {
        if (HWND okButton = GetDlgItem(hwnd, IDOK)) {
            SetWindowTextW(okButton, L"Save");
        }
    }
    return 0;
}

}  // namespace

OptionsDialogResult ShowOptionsDialog(HWND parent, OptionsDialogPage initialPage,
                                     const wchar_t* focusSavedGroupId, bool editFocusedGroup) {
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_TAB_CLASSES | ICC_BAR_CLASSES | ICC_LISTVIEW_CLASSES};
    InitCommonControlsEx(&icc);

    OptionsDialogResult result;
    OptionsDialogData data;
    auto& store = OptionsStore::Instance();
    std::wstring optionsLoadError;
    if (!store.Load(&optionsLoadError)) {
        if (!optionsLoadError.empty()) {
            LogMessage(LogLevel::Warning, L"ShowOptionsDialog failed to load options: %ls", optionsLoadError.c_str());
        } else {
            LogMessage(LogLevel::Warning, L"ShowOptionsDialog failed to load options");
        }
    }
    data.originalOptions = store.Get();
    data.workingOptions = data.originalOptions;
    const int initialTabIndex = static_cast<int>(initialPage);
    data.initialTab = initialTabIndex;
    auto& groupStore = GroupStore::Instance();
    std::wstring groupLoadError;
    if (!groupStore.Load(&groupLoadError)) {
        if (!groupLoadError.empty()) {
            LogMessage(LogLevel::Warning, L"ShowOptionsDialog failed to load saved groups: %ls",
                       groupLoadError.c_str());
        } else {
            LogMessage(LogLevel::Warning, L"ShowOptionsDialog failed to load saved groups");
        }
    }
    data.originalGroups = groupStore.Groups();
    data.workingGroups = data.originalGroups;
    data.workingGroupIds.clear();
    data.workingGroupIds.reserve(data.workingGroups.size());
    for (const auto& group : data.workingGroups) {
        data.workingGroupIds.push_back(group.name);
    }
    data.removedGroupIds.clear();
    if (focusSavedGroupId && *focusSavedGroupId) {
        data.focusSavedGroupId = focusSavedGroupId;
        data.focusShouldEdit = editFocusedGroup;
        data.focusHandled = false;
    } else {
        data.focusSavedGroupId.clear();
        data.focusShouldEdit = false;
        data.focusHandled = true;
    }

    std::vector<BYTE> mainTemplate = BuildMainPageTemplate();
    std::vector<BYTE> customizationTemplate = BuildCustomizationPageTemplate();
    std::vector<BYTE> glowTemplate = BuildGlowPageTemplate();
    std::vector<BYTE> contextTemplate = BuildContextMenuPageTemplate();
    std::vector<BYTE> groupTemplate = BuildGroupPageTemplate();

    auto mainTemplateMemory = AllocateAlignedTemplate(mainTemplate);
    auto customizationTemplateMemory = AllocateAlignedTemplate(customizationTemplate);
    auto glowTemplateMemory = AllocateAlignedTemplate(glowTemplate);
    auto contextTemplateMemory = AllocateAlignedTemplate(contextTemplate);
    auto groupTemplateMemory = AllocateAlignedTemplate(groupTemplate);
    if (!mainTemplateMemory || !customizationTemplateMemory || !glowTemplateMemory ||
        !contextTemplateMemory || !groupTemplateMemory) {
        result.saved = false;
        result.groupsChanged = false;
        result.optionsChanged = false;
        return result;
    }

    std::array<std::wstring, 5> pageTitles;
    auto loadTitle = [&](size_t index, UINT resourceId, const wchar_t* fallback) {
        wchar_t buffer[128] = {};
        if (LoadStringW(GetModuleHandleInstance(), resourceId, buffer, ARRAYSIZE(buffer)) > 0) {
            pageTitles[index] = buffer;
        } else {
            pageTitles[index] = fallback;
        }
        return pageTitles[index].c_str();
    };

    PROPSHEETPAGEW pages[5] = {};
    pages[0].dwSize = sizeof(PROPSHEETPAGEW);
    pages[0].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[0].hInstance = GetModuleHandleInstance();
    pages[0].pResource = mainTemplateMemory.get();
    pages[0].pfnDlgProc = MainOptionsPageProc;
    pages[0].lParam = reinterpret_cast<LPARAM>(&data);
    pages[0].pszTitle = loadTitle(0, IDS_OPTIONS_TITLE_GENERAL, L"General");

    pages[1].dwSize = sizeof(PROPSHEETPAGEW);
    pages[1].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[1].hInstance = GetModuleHandleInstance();
    pages[1].pResource = customizationTemplateMemory.get();
    pages[1].pfnDlgProc = CustomizationsPageProc;
    pages[1].lParam = reinterpret_cast<LPARAM>(&data);
    pages[1].pszTitle = loadTitle(1, IDS_OPTIONS_TITLE_CUSTOMIZATIONS, L"Customizations");

    pages[2].dwSize = sizeof(PROPSHEETPAGEW);
    pages[2].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[2].hInstance = GetModuleHandleInstance();
    pages[2].pResource = glowTemplateMemory.get();
    pages[2].pfnDlgProc = GlowPageProc;
    pages[2].lParam = reinterpret_cast<LPARAM>(&data);
    pages[2].pszTitle = loadTitle(2, IDS_OPTIONS_TITLE_GLOW, L"Glow");

    pages[3].dwSize = sizeof(PROPSHEETPAGEW);
    pages[3].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[3].hInstance = GetModuleHandleInstance();
    pages[3].pResource = contextTemplateMemory.get();
    pages[3].pfnDlgProc = ContextMenuPageProc;
    pages[3].lParam = reinterpret_cast<LPARAM>(&data);
    pages[3].pszTitle = loadTitle(3, IDS_OPTIONS_TITLE_CONTEXT_MENUS, L"Context Menus");

    pages[4].dwSize = sizeof(PROPSHEETPAGEW);
    pages[4].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[4].hInstance = GetModuleHandleInstance();
    pages[4].pResource = groupTemplateMemory.get();
    pages[4].pfnDlgProc = GroupManagementPageProc;
    pages[4].lParam = reinterpret_cast<LPARAM>(&data);
    pages[4].pszTitle = loadTitle(4, IDS_OPTIONS_TITLE_GROUPS, L"Groups && Islands");

    PROPSHEETHEADERW header{};
    header.dwSize = sizeof(PROPSHEETHEADERW);
    header.dwFlags = PSH_PROPSHEETPAGE | PSH_NOAPPLYNOW | PSH_USECALLBACK;
    header.hwndParent = parent;
    header.hInstance = GetModuleHandleInstance();
    header.pszCaption = L"ShellTabs Options";
    header.nPages = ARRAYSIZE(pages);
    header.nStartPage =
        (initialTabIndex >= 0 && initialTabIndex < static_cast<int>(header.nPages)) ? initialTabIndex : 0;
    header.ppsp = pages;
    header.pfnCallback = OptionsSheetCallback;

    INT_PTR dialogResult = PropertySheetW(&header);
    if (dialogResult == IDOK && data.applyInvoked) {
        result.saved = true;
        result.optionsChanged = data.workingOptions != data.originalOptions;
        const bool groupsChanged = !AreSavedGroupsEqual(data.originalGroups, data.workingGroups);
        result.groupsChanged = groupsChanged;
        result.savedGroups = data.workingGroups;
        result.removedGroupIds = data.removedGroupIds;
        result.renamedGroups.clear();
        for (size_t i = 0; i < data.workingGroups.size() && i < data.workingGroupIds.size(); ++i) {
            const std::wstring& originalId = data.workingGroupIds[i];
            const std::wstring& updatedName = data.workingGroups[i].name;
            if (originalId.empty()) {
                continue;
            }
            if (!CaseInsensitiveEquals(originalId, updatedName)) {
                result.renamedGroups.emplace_back(originalId, updatedName);
            }
        }
        store.Set(data.workingOptions);
        store.Save();
        if (result.optionsChanged) {
            ForceExplorerUIRefresh(parent);
        }
        if (groupsChanged) {
            if (!groupStore.Load(&groupLoadError)) {
                if (!groupLoadError.empty()) {
                    LogMessage(LogLevel::Warning, L"ShowOptionsDialog failed to reload saved groups: %ls",
                               groupLoadError.c_str());
                } else {
                    LogMessage(LogLevel::Warning, L"ShowOptionsDialog failed to reload saved groups");
                }
            }
            std::vector<SavedGroup> existingGroups = groupStore.Groups();
            for (const auto& existing : existingGroups) {
                bool found = false;
                for (const auto& updated : data.workingGroups) {
                    if (CaseInsensitiveEquals(existing.name, updated.name)) {
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    groupStore.Remove(existing.name);
                }
            }
            for (const auto& group : data.workingGroups) {
                groupStore.Upsert(group);
            }
            result.savedGroups = groupStore.Groups();
        }
        if (!result.removedGroupIds.empty()) {
            std::vector<std::wstring> filtered;
            filtered.reserve(result.removedGroupIds.size());
            for (const auto& removedId : result.removedGroupIds) {
                bool stillRemoved = true;
                for (const auto& group : result.savedGroups) {
                    if (CaseInsensitiveEquals(group.name, removedId)) {
                        stillRemoved = false;
                        break;
                    }
                }
                if (stillRemoved) {
                    filtered.push_back(removedId);
                }
            }
            result.removedGroupIds.swap(filtered);
        }
        groupStore.RecordChanges(result.renamedGroups, result.removedGroupIds);
        const UINT savedGroupsMessage = GetSavedGroupsChangedMessage();
        if (savedGroupsMessage != 0) {
            SendNotifyMessageW(HWND_BROADCAST, savedGroupsMessage, 0, 0);
            if (parent) {
                SendNotifyMessageW(parent, savedGroupsMessage, 0, 0);
            }
        }
        for (const auto& path : data.pendingCachedImageRemovals) {
            if (!path.empty()) {
                DeleteFileW(path.c_str());
            }
        }
    } else {
        result.saved = false;
        result.groupsChanged = false;
        result.optionsChanged = false;
        if (data.previewOptionsBroadcasted) {
            store.Set(data.originalOptions);
            ForceExplorerUIRefresh(parent);
        }
        for (const auto& path : data.createdCachedImagePaths) {
            if (!path.empty()) {
                DeleteFileW(path.c_str());
            }
        }
    }

    if (data.breadcrumbBgStartBrush) {
        DeleteObject(data.breadcrumbBgStartBrush);
        data.breadcrumbBgStartBrush = nullptr;
    }
    if (data.breadcrumbBgEndBrush) {
        DeleteObject(data.breadcrumbBgEndBrush);
        data.breadcrumbBgEndBrush = nullptr;
    }
    if (data.breadcrumbFontStartBrush) {
        DeleteObject(data.breadcrumbFontStartBrush);
        data.breadcrumbFontStartBrush = nullptr;
    }
    if (data.breadcrumbFontEndBrush) {
        DeleteObject(data.breadcrumbFontEndBrush);
        data.breadcrumbFontEndBrush = nullptr;
    }
    if (data.progressStartBrush) {
        DeleteObject(data.progressStartBrush);
        data.progressStartBrush = nullptr;
    }
    if (data.progressEndBrush) {
        DeleteObject(data.progressEndBrush);
        data.progressEndBrush = nullptr;
    }
    if (data.tabSelectedBrush) {
        DeleteObject(data.tabSelectedBrush);
        data.tabSelectedBrush = nullptr;
    }
    if (data.tabUnselectedBrush) {
        DeleteObject(data.tabUnselectedBrush);
        data.tabUnselectedBrush = nullptr;
    }
    if (data.glowPrimaryBrush) {
        DeleteObject(data.glowPrimaryBrush);
        data.glowPrimaryBrush = nullptr;
    }
    if (data.glowSecondaryBrush) {
        DeleteObject(data.glowSecondaryBrush);
        data.glowSecondaryBrush = nullptr;
    }
    if (data.universalBackgroundPreview) {
        DeleteObject(data.universalBackgroundPreview);
        data.universalBackgroundPreview = nullptr;
    }
    if (data.folderBackgroundPreview) {
        DeleteObject(data.folderBackgroundPreview);
        data.folderBackgroundPreview = nullptr;
    }
    return result;
}

}  // namespace shelltabs
