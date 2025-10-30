#include "OptionsDialog.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <CommCtrl.h>
#include <prsht.h>

#include <algorithm>
#include <cwchar>
#include <malloc.h>
#include <memory>
#include <random>
#include <string>
#include <utility>
#include <vector>
#include <cstring>
#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commdlg.h>
#include <wrl/client.h>
#include <objbase.h>

#include "BackgroundCache.h"
#include "GroupStore.h"
#include "Module.h"
#include "OptionsStore.h"
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
constexpr SIZE kUniversalPreviewSize = {96, 72};
constexpr SIZE kFolderPreviewSize = {64, 64};

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
    IDC_MAIN_DOCK_LABEL = 5045,
    IDC_MAIN_DOCK_COMBO = 5046,

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
};

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
    HBRUSH breadcrumbBgStartBrush = nullptr;
    HBRUSH breadcrumbBgEndBrush = nullptr;
    HBRUSH breadcrumbFontStartBrush = nullptr;
    HBRUSH breadcrumbFontEndBrush = nullptr;
    HBRUSH progressStartBrush = nullptr;
    HBRUSH progressEndBrush = nullptr;
    HBRUSH tabSelectedBrush = nullptr;
    HBRUSH tabUnselectedBrush = nullptr;
    HBITMAP universalBackgroundPreview = nullptr;
    HBITMAP folderBackgroundPreview = nullptr;
    std::wstring lastFolderBrowsePath;
    std::wstring lastImageBrowseDirectory;
    std::vector<std::wstring> createdCachedImagePaths;
    std::vector<std::wstring> pendingCachedImageRemovals;
    std::vector<ChildPlacement> customizationChildPlacements;
    int customizationScrollPos = 0;
    int customizationContentHeight = 0;
    int customizationScrollMax = 0;
};

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

    OptionsStore::Instance().Set(data->workingOptions);
    data->previewOptionsBroadcasted = true;
    ForceExplorerUIRefresh(GetParent(pageWindow));
}

std::vector<BYTE> BuildMainPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 5;
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
    auto* exampleStatic = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    exampleStatic->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    exampleStatic->dwExtendedStyle = 0;
    exampleStatic->x = 10;
    exampleStatic->y = 56;
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
    auto* dockLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    dockLabel->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    dockLabel->dwExtendedStyle = 0;
    dockLabel->x = 10;
    dockLabel->y = 122;
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
    dockCombo->y = 136;
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

    AlignDialogBuffer(data);
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

bool CopyImageToCache(const std::wstring& sourcePath, const std::wstring& displayName, CachedImageMetadata* metadata,
                      std::wstring* createdPath) {
    return CopyImageToBackgroundCache(sourcePath, displayName, metadata, createdPath);
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
                            IDC_CUSTOM_BACKGROUND_FOLDER_NAME};
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
}

HBITMAP LoadPreviewBitmap(const std::wstring& path, const SIZE& size) {
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

void UpdateUniversalBackgroundPreview(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HBITMAP bitmap = LoadPreviewBitmap(data->workingOptions.universalFolderBackgroundImage.cachedImagePath,
                                       kUniversalPreviewSize);
    SetPreviewBitmap(hwnd, IDC_CUSTOM_BACKGROUND_PREVIEW, &data->universalBackgroundPreview, bitmap);
    const std::wstring& name = data->workingOptions.universalFolderBackgroundImage.displayName;
    SetDlgItemTextW(hwnd, IDC_CUSTOM_BACKGROUND_UNIVERSAL_NAME, name.empty() ? L"(None)" : name.c_str());
}

void UpdateSelectedFolderBackgroundPreview(HWND hwnd, OptionsDialogData* data) {
    if (!data) {
        return;
    }
    HWND list = GetDlgItem(hwnd, IDC_CUSTOM_BACKGROUND_LIST);
    const int selection = GetSelectedFolderBackgroundIndex(list);
    HBITMAP bitmap = nullptr;
    std::wstring name;
    if (selection >= 0 && static_cast<size_t>(selection) < data->workingOptions.folderBackgroundEntries.size()) {
        const auto& entry = data->workingOptions.folderBackgroundEntries[static_cast<size_t>(selection)];
        name = entry.image.displayName;
        bitmap = LoadPreviewBitmap(entry.image.cachedImagePath, kFolderPreviewSize);
    }
    SetPreviewBitmap(hwnd, IDC_CUSTOM_BACKGROUND_FOLDER_PREVIEW, &data->folderBackgroundPreview, bitmap);
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
    const std::wstring previousPath = metadata.cachedImagePath;
    if (!CopyImageToCache(imagePath, displayName, &metadata, &createdPath)) {
        MessageBoxW(hwnd, L"Unable to copy the selected image.", L"ShellTabs", MB_OK | MB_ICONERROR);
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
    if (!CopyImageToCache(imagePath, displayName, &metadata, &createdPath)) {
        MessageBoxW(hwnd, L"Unable to copy the selected image.", L"ShellTabs", MB_OK | MB_ICONERROR);
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
        const std::wstring previousPath = metadata.cachedImagePath;
        if (!CopyImageToCache(imagePath, displayName, &metadata, &createdPath)) {
            MessageBoxW(hwnd, L"Unable to copy the selected image.", L"ShellTabs", MB_OK | MB_ICONERROR);
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

bool HandleColorButtonClick(HWND hwnd, OptionsDialogData* data, WORD controlId) {
    if (!data) {
        return false;
    }

    COLORREF initial = RGB(255, 255, 255);
    HBRUSH* targetBrush = nullptr;
    int previewId = 0;
    COLORREF* targetColor = nullptr;

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
        default:
            return false;
    }

    if (targetColor && PromptForColor(hwnd, initial, targetColor)) {
        SetPreviewColor(hwnd, previewId, targetBrush, *targetColor);
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
    HWND hwndParent = GetAncestor(page, GA_ROOT);
    std::wstring name = L"New Group";
    COLORREF color = RGB(0, 120, 215);
    if (!PromptForTextInput(hwndParent, L"Create Tab Group", L"Group name:", &name, &color)) {
        return;
    }
    if (name.empty()) {
        MessageBoxW(hwndParent, L"Group name cannot be empty.", L"ShellTabs", MB_OK | MB_ICONWARNING);
        return;
    }
    for (const auto& group : data->workingGroups) {
        if (_wcsicmp(group.name.c_str(), name.c_str()) == 0) {
            MessageBoxW(hwndParent, L"A saved group with that name already exists.", L"ShellTabs",
                        MB_OK | MB_ICONWARNING);
            return;
        }
    }

    SavedGroup group;
    group.name = name;
    group.color = color;
    group.tabPaths = {L"C:\\"};
    data->workingGroups.push_back(std::move(group));
    data->groupsChanged = true;
    HWND list = GetDlgItem(page, IDC_GROUP_LIST);
    RefreshGroupList(list, data);
    const int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    for (int i = 0; i < count; ++i) {
        SendMessageW(list, LB_SETCURSEL, i, 0);
        std::wstring current = GetSelectedGroupName(list);
        if (_wcsicmp(current.c_str(), name.c_str()) == 0) {
            break;
        }
    }
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
    UpdateGroupButtons(page);
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
                const wchar_t example[] =
                    L"Example: if a group opens to C:\\test and you browse to C\\test\\child, "
                    L"enabling this option reopens the child folder next time.";
                SetDlgItemTextW(hwnd, IDC_MAIN_EXAMPLE, example);

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
                    if (HIWORD(wParam) == BN_CLICKED) {
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
            if (((LPNMHDR)lParam)->code == PSN_APPLY) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                if (data) {
                    data->workingOptions.reopenOnCrash =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_REOPEN) == BST_CHECKED;
                    data->workingOptions.persistGroupPaths =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_PERSIST) == BST_CHECKED;
                    HWND combo = GetDlgItem(hwnd, IDC_MAIN_DOCK_COMBO);
                    if (combo) {
                        const LRESULT selection = SendMessageW(combo, CB_GETCURSEL, 0, 0);
                        if (selection >= 0) {
                            const LRESULT value = SendMessageW(combo, CB_GETITEMDATA, selection, 0);
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

INT_PTR CALLBACK CustomizationsPageProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* data = reinterpret_cast<OptionsDialogData*>(reinterpret_cast<PROPSHEETPAGEW*>(lParam)->lParam);
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(data));
            if (data) {
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
                UpdateGradientControlsEnabled(hwnd, data->workingOptions.enableBreadcrumbGradient,
                                              data->workingOptions.enableBreadcrumbFontGradient);
                UpdateGradientColorControlsEnabled(
                    hwnd,
                    data->workingOptions.enableBreadcrumbGradient &&
                        data->workingOptions.useCustomBreadcrumbGradientColors,
                    data->workingOptions.enableBreadcrumbFontGradient &&
                        data->workingOptions.useCustomBreadcrumbFontColors);
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
                UpdateTabColorControlsEnabled(hwnd, data->workingOptions.useCustomTabSelectedColor,
                                              data->workingOptions.useCustomTabUnselectedColor);
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
                data->lastImageBrowseDirectory = ExtractDirectoryFromPath(
                    data->workingOptions.universalFolderBackgroundImage.cachedImagePath);
                CaptureCustomizationChildPlacements(hwnd, data);
                UpdateCustomizationScrollInfo(hwnd, data);
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
                case IDC_MAIN_BREADCRUMB:
                case IDC_MAIN_BREADCRUMB_FONT:
                case IDC_MAIN_BREADCRUMB_BG_CUSTOM:
                case IDC_MAIN_BREADCRUMB_FONT_CUSTOM:
                case IDC_MAIN_PROGRESS_CUSTOM:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data) {
                            const bool backgroundEnabled =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB) == BST_CHECKED;
                            const bool fontEnabled =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT) == BST_CHECKED;
                            UpdateGradientControlsEnabled(hwnd, backgroundEnabled, fontEnabled);
                            const bool bgCustom =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_BG_CUSTOM) == BST_CHECKED;
                            const bool fontCustom =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT_CUSTOM) == BST_CHECKED;
                            UpdateGradientColorControlsEnabled(hwnd, backgroundEnabled && bgCustom,
                                                               fontEnabled && fontCustom);
                            const bool progressCustom =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_PROGRESS_CUSTOM) == BST_CHECKED;
                            UpdateProgressColorControlsEnabled(hwnd, progressCustom);
                        }
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_MAIN_TAB_SELECTED_CHECK:
                case IDC_MAIN_TAB_UNSELECTED_CHECK:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data) {
                            const bool tabSelectedCustom =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_SELECTED_CHECK) == BST_CHECKED;
                            const bool tabUnselectedCustom =
                                IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_UNSELECTED_CHECK) == BST_CHECKED;
                            UpdateTabColorControlsEnabled(hwnd, tabSelectedCustom, tabUnselectedCustom);
                        }
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_MAIN_BREADCRUMB_BG_START_BUTTON:
                case IDC_MAIN_BREADCRUMB_BG_END_BUTTON:
                case IDC_MAIN_BREADCRUMB_FONT_START_BUTTON:
                case IDC_MAIN_BREADCRUMB_FONT_END_BUTTON:
                case IDC_MAIN_PROGRESS_START_BUTTON:
                case IDC_MAIN_PROGRESS_END_BUTTON:
                case IDC_MAIN_TAB_SELECTED_BUTTON:
                case IDC_MAIN_TAB_UNSELECTED_BUTTON:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (HandleColorButtonClick(hwnd, data, LOWORD(wParam))) {
                            SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                        }
                    }
                    return TRUE;
                case IDC_CUSTOM_BACKGROUND_ENABLE:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data) {
                            const bool enabled =
                                IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED;
                            UpdateFolderBackgroundControlsEnabled(hwnd, enabled);
                            UpdateFolderBackgroundButtons(hwnd);
                        }
                        SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
                    }
                    return TRUE;
                case IDC_CUSTOM_BACKGROUND_BROWSE:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        HandleUniversalBackgroundBrowse(hwnd, data);
                    }
                    return TRUE;
                case IDC_CUSTOM_BACKGROUND_ADD:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data &&
                            IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                            HandleAddFolderBackgroundEntry(hwnd, data);
                        }
                    }
                    return TRUE;
                case IDC_CUSTOM_BACKGROUND_EDIT:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data &&
                            IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                            HandleEditFolderBackgroundEntry(hwnd, data);
                        }
                    }
                    return TRUE;
                case IDC_CUSTOM_BACKGROUND_REMOVE:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (data &&
                            IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                            HandleRemoveFolderBackgroundEntry(hwnd, data);
                        }
                    }
                    return TRUE;
            }
            break;
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
        case WM_HSCROLL: {
            HWND slider = reinterpret_cast<HWND>(lParam);
            if (!slider) {
                return TRUE;
            }
            const int controlId = GetDlgCtrlID(slider);
            auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            bool previewNeeded = false;
            switch (controlId) {
                case IDC_MAIN_BREADCRUMB_BG_SLIDER: {
                    const int sliderValue = ClampPercentageValue(
                        static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0)));
                    UpdatePercentageLabel(hwnd, IDC_MAIN_BREADCRUMB_BG_VALUE, sliderValue);
                    if (data && data->workingOptions.breadcrumbGradientTransparency != sliderValue) {
                        data->workingOptions.breadcrumbGradientTransparency = sliderValue;
                        previewNeeded = true;
                    }
                    break;
                }
                case IDC_MAIN_BREADCRUMB_FONT_SLIDER: {
                    const int sliderValue = ClampPercentageValue(
                        static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0)));
                    const int brightnessValue = InvertPercentageValue(sliderValue);
                    UpdatePercentageLabel(hwnd, IDC_MAIN_BREADCRUMB_FONT_VALUE, brightnessValue);
                    if (data && data->workingOptions.breadcrumbFontBrightness != brightnessValue) {
                        data->workingOptions.breadcrumbFontBrightness = brightnessValue;
                        previewNeeded = true;
                    }
                    break;
                }
                case IDC_MAIN_BREADCRUMB_HIGHLIGHT_SLIDER: {
                    const int sliderValue = ClampMultiplierValue(
                        static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0)));
                    UpdateMultiplierLabel(hwnd, IDC_MAIN_BREADCRUMB_HIGHLIGHT_VALUE, sliderValue);
                    if (data && data->workingOptions.breadcrumbHighlightAlphaMultiplier != sliderValue) {
                        data->workingOptions.breadcrumbHighlightAlphaMultiplier = sliderValue;
                        previewNeeded = true;
                    }
                    break;
                }
                case IDC_MAIN_BREADCRUMB_DROPDOWN_SLIDER: {
                    const int sliderValue = ClampMultiplierValue(
                        static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0)));
                    UpdateMultiplierLabel(hwnd, IDC_MAIN_BREADCRUMB_DROPDOWN_VALUE, sliderValue);
                    if (data && data->workingOptions.breadcrumbDropdownAlphaMultiplier != sliderValue) {
                        data->workingOptions.breadcrumbDropdownAlphaMultiplier = sliderValue;
                        previewNeeded = true;
                    }
                    break;
                }
                default:
                    return TRUE;
            }
            if (previewNeeded) {
                ApplyCustomizationPreview(hwnd, data);
            }
            SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
            return TRUE;
        }
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
            int newPos = data->customizationScrollPos;
            switch (LOWORD(wParam)) {
                case SB_LINEUP:
                    newPos -= 16;
                    break;
                case SB_LINEDOWN:
                    newPos += 16;
                    break;
                case SB_PAGEUP:
                    newPos -= 80;
                    break;
                case SB_PAGEDOWN:
                    newPos += 80;
                    break;
                case SB_TOP:
                    newPos = 0;
                    break;
                case SB_BOTTOM:
                    newPos = data->customizationScrollMax;
                    break;
                case SB_THUMBTRACK:
                case SB_THUMBPOSITION: {
                    SCROLLINFO info{sizeof(info)};
                    info.fMask = SIF_TRACKPOS;
                    if (GetScrollInfo(hwnd, SB_VERT, &info)) {
                        newPos = info.nTrackPos;
                    }
                    break;
                }
                default:
                    break;
            }
            newPos = std::clamp(newPos, 0, data->customizationScrollMax);
            if (newPos != data->customizationScrollPos) {
                data->customizationScrollPos = newPos;
                SetScrollPos(hwnd, SB_VERT, newPos, TRUE);
                RepositionCustomizationChildren(hwnd, data);
            }
            return TRUE;
        }
        case WM_NOTIFY: {
            auto* header = reinterpret_cast<LPNMHDR>(lParam);
            if (!header) {
                break;
            }
            if (header->idFrom == IDC_CUSTOM_BACKGROUND_LIST) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                switch (header->code) {
                    case LVN_ITEMCHANGED: {
                        auto* changed = reinterpret_cast<NMLISTVIEW*>(lParam);
                        if (changed && (changed->uChanged & LVIF_STATE) != 0) {
                            UpdateFolderBackgroundButtons(hwnd);
                            UpdateSelectedFolderBackgroundPreview(hwnd, data);
                        }
                        return TRUE;
                    }
                    case NM_DBLCLK:
                        if (data &&
                            IsDlgButtonChecked(hwnd, IDC_CUSTOM_BACKGROUND_ENABLE) == BST_CHECKED) {
                            HandleEditFolderBackgroundEntry(hwnd, data);
                        }
                        return TRUE;
                    default:
                        break;
                }
            }
            if (header->code == PSN_APPLY) {
                auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
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

OptionsDialogResult ShowOptionsDialog(HWND parent, int initialTab) {
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_TAB_CLASSES | ICC_BAR_CLASSES | ICC_LISTVIEW_CLASSES};
    InitCommonControlsEx(&icc);

    OptionsDialogResult result;
    OptionsDialogData data;
    auto& store = OptionsStore::Instance();
    store.Load();
    data.originalOptions = store.Get();
    data.workingOptions = data.originalOptions;
    data.initialTab = initialTab;
    auto& groupStore = GroupStore::Instance();
    groupStore.Load();
    data.originalGroups = groupStore.Groups();
    data.workingGroups = data.originalGroups;

    std::vector<BYTE> mainTemplate = BuildMainPageTemplate();
    std::vector<BYTE> customizationTemplate = BuildCustomizationPageTemplate();
    std::vector<BYTE> groupTemplate = BuildGroupPageTemplate();

    auto mainTemplateMemory = AllocateAlignedTemplate(mainTemplate);
    auto customizationTemplateMemory = AllocateAlignedTemplate(customizationTemplate);
    auto groupTemplateMemory = AllocateAlignedTemplate(groupTemplate);
    if (!mainTemplateMemory || !customizationTemplateMemory || !groupTemplateMemory) {
        result.saved = false;
        result.groupsChanged = false;
        result.optionsChanged = false;
        return result;
    }

    PROPSHEETPAGEW pages[3] = {};
    pages[0].dwSize = sizeof(PROPSHEETPAGEW);
    pages[0].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[0].hInstance = GetModuleHandleInstance();
    pages[0].pResource = mainTemplateMemory.get();
    pages[0].pfnDlgProc = MainOptionsPageProc;
    pages[0].lParam = reinterpret_cast<LPARAM>(&data);
    pages[0].pszTitle = L"General";

    pages[1].dwSize = sizeof(PROPSHEETPAGEW);
    pages[1].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[1].hInstance = GetModuleHandleInstance();
    pages[1].pResource = customizationTemplateMemory.get();
    pages[1].pfnDlgProc = CustomizationsPageProc;
    pages[1].lParam = reinterpret_cast<LPARAM>(&data);
    pages[1].pszTitle = L"Customizations";

    pages[2].dwSize = sizeof(PROPSHEETPAGEW);
    pages[2].dwFlags = PSP_DLGINDIRECT | PSP_USETITLE;
    pages[2].hInstance = GetModuleHandleInstance();
    pages[2].pResource = groupTemplateMemory.get();
    pages[2].pfnDlgProc = GroupManagementPageProc;
    pages[2].lParam = reinterpret_cast<LPARAM>(&data);
    pages[2].pszTitle = L"Groups && Islands";

    PROPSHEETHEADERW header{};
    header.dwSize = sizeof(PROPSHEETHEADERW);
    header.dwFlags = PSH_PROPSHEETPAGE | PSH_NOAPPLYNOW | PSH_USECALLBACK;
    header.hwndParent = parent;
    header.hInstance = GetModuleHandleInstance();
    header.pszCaption = L"ShellTabs Options";
    header.nPages = ARRAYSIZE(pages);
    header.nStartPage = (initialTab >= 0 && initialTab < static_cast<int>(header.nPages)) ? initialTab : 0;
    header.ppsp = pages;
    header.pfnCallback = OptionsSheetCallback;

    INT_PTR dialogResult = PropertySheetW(&header);
    if (dialogResult == IDOK && data.applyInvoked) {
        result.saved = true;
        result.optionsChanged = data.workingOptions != data.originalOptions;
        const bool groupsChanged = !AreSavedGroupsEqual(data.originalGroups, data.workingGroups);
        result.groupsChanged = groupsChanged;
        store.Set(data.workingOptions);
        store.Save();
        if (result.optionsChanged) {
            ForceExplorerUIRefresh(parent);
        }
        if (groupsChanged) {
            auto& groupStoreToUpdate = GroupStore::Instance();
            groupStoreToUpdate.Load();
            std::vector<SavedGroup> existingGroups = groupStoreToUpdate.Groups();
            for (const auto& existing : existingGroups) {
                bool found = false;
                for (const auto& updated : data.workingGroups) {
                    if (CaseInsensitiveEquals(existing.name, updated.name)) {
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    groupStoreToUpdate.Remove(existing.name);
                }
            }
            for (const auto& group : data.workingGroups) {
                groupStoreToUpdate.Upsert(group);
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
