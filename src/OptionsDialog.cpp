#include "OptionsDialog.h"

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

#include "GroupStore.h"
#include "Module.h"
#include "OptionsStore.h"
#include "Utilities.h"

namespace shelltabs {
namespace {

constexpr int kMainCheckboxWidth = 210;
constexpr int kMainDialogWidth = 260;
constexpr int kMainDialogHeight = 360;
constexpr int kGroupDialogWidth = 320;
constexpr int kGroupDialogHeight = 200;
constexpr int kEditorWidth = 340;
constexpr int kEditorHeight = 220;

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
    IDC_MAIN_BREADCRUMB_BG_CUSTOM = 5012,
    IDC_MAIN_BREADCRUMB_BG_START_LABEL = 5013,
    IDC_MAIN_BREADCRUMB_BG_START_PREVIEW = 5014,
    IDC_MAIN_BREADCRUMB_BG_START_BUTTON = 5015,
    IDC_MAIN_BREADCRUMB_BG_END_LABEL = 5016,
    IDC_MAIN_BREADCRUMB_BG_END_PREVIEW = 5017,
    IDC_MAIN_BREADCRUMB_BG_END_BUTTON = 5018,
    IDC_MAIN_BREADCRUMB_FONT_CUSTOM = 5019,
    IDC_MAIN_BREADCRUMB_FONT_START_LABEL = 5020,
    IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW = 5021,
    IDC_MAIN_BREADCRUMB_FONT_START_BUTTON = 5022,
    IDC_MAIN_BREADCRUMB_FONT_END_LABEL = 5023,
    IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW = 5024,
    IDC_MAIN_BREADCRUMB_FONT_END_BUTTON = 5025,
    IDC_MAIN_TAB_SELECTED_CHECK = 5026,
    IDC_MAIN_TAB_SELECTED_PREVIEW = 5027,
    IDC_MAIN_TAB_SELECTED_BUTTON = 5028,
    IDC_MAIN_TAB_UNSELECTED_CHECK = 5029,
    IDC_MAIN_TAB_UNSELECTED_PREVIEW = 5030,
    IDC_MAIN_TAB_UNSELECTED_BUTTON = 5031,

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

std::vector<BYTE> BuildMainPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 3;
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

    return data;
}

std::vector<BYTE> BuildCustomizationPageTemplate() {
    std::vector<BYTE> data(sizeof(DLGTEMPLATE), 0);
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
    dlg->cdit = 30;
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
    breadcrumbGroup->cy = 260;
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
    addCheckbox(IDC_MAIN_BREADCRUMB_BG_CUSTOM, 16, 138, L"Use custom background gradient colors");
    addStatic(IDC_MAIN_BREADCRUMB_BG_START_LABEL, 24, 156, 60, 10, L"Start:");
    addPreview(IDC_MAIN_BREADCRUMB_BG_START_PREVIEW, 86, 154);
    addButton(IDC_MAIN_BREADCRUMB_BG_START_BUTTON, 124, 153, L"Choose");
    addStatic(IDC_MAIN_BREADCRUMB_BG_END_LABEL, 24, 176, 60, 10, L"End:");
    addPreview(IDC_MAIN_BREADCRUMB_BG_END_PREVIEW, 86, 174);
    addButton(IDC_MAIN_BREADCRUMB_BG_END_BUTTON, 124, 173, L"Choose");
    addCheckbox(IDC_MAIN_BREADCRUMB_FONT_CUSTOM, 16, 198, L"Use custom breadcrumb text colors");
    addStatic(IDC_MAIN_BREADCRUMB_FONT_START_LABEL, 24, 216, 60, 10, L"Start:");
    addPreview(IDC_MAIN_BREADCRUMB_FONT_START_PREVIEW, 86, 214);
    addButton(IDC_MAIN_BREADCRUMB_FONT_START_BUTTON, 124, 213, L"Choose");
    addStatic(IDC_MAIN_BREADCRUMB_FONT_END_LABEL, 24, 236, 60, 10, L"End:");
    addPreview(IDC_MAIN_BREADCRUMB_FONT_END_PREVIEW, 86, 234);
    addButton(IDC_MAIN_BREADCRUMB_FONT_END_BUTTON, 124, 233, L"Choose");

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* tabsGroup = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    tabsGroup->style = WS_CHILD | WS_VISIBLE | BS_GROUPBOX;
    tabsGroup->dwExtendedStyle = 0;
    tabsGroup->x = 6;
    tabsGroup->y = 272;
    tabsGroup->cx = kMainDialogWidth - 12;
    tabsGroup->cy = 88;
    tabsGroup->id = 0;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Tabs");
    AppendWord(data, 0);

    addCheckbox(IDC_MAIN_TAB_SELECTED_CHECK, 16, 288, L"Use custom selected tab color");
    addPreview(IDC_MAIN_TAB_SELECTED_PREVIEW, 24, 306);
    addButton(IDC_MAIN_TAB_SELECTED_BUTTON, 62, 305, L"Choose");
    addCheckbox(IDC_MAIN_TAB_UNSELECTED_CHECK, 16, 324, L"Use custom unselected tab color");
    addPreview(IDC_MAIN_TAB_UNSELECTED_PREVIEW, 24, 342);
    addButton(IDC_MAIN_TAB_UNSELECTED_BUTTON, 62, 341, L"Choose");

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
    dlg->cdit = 9;
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
    auto* pathsLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    pathsLabel->style = WS_CHILD | WS_VISIBLE;
    pathsLabel->dwExtendedStyle = 0;
    pathsLabel->x = 10;
    pathsLabel->y = 64;
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
    pathList->y = 76;
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
    addButton->y = 76;
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
    editButton->y = 96;
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
    removeButton->y = 116;
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

struct OptionsDialogData {
    ShellTabsOptions originalOptions;
    ShellTabsOptions workingOptions;
    bool applyInvoked = false;
    bool groupsChanged = false;
    int initialTab = 0;
    std::vector<SavedGroup> originalGroups;
    std::vector<SavedGroup> workingGroups;
    HBRUSH breadcrumbBgStartBrush = nullptr;
    HBRUSH breadcrumbBgEndBrush = nullptr;
    HBRUSH breadcrumbFontStartBrush = nullptr;
    HBRUSH breadcrumbFontEndBrush = nullptr;
    HBRUSH tabSelectedBrush = nullptr;
    HBRUSH tabUnselectedBrush = nullptr;
};

struct GroupEditorContext {
    SavedGroup working;
    std::wstring originalName;
    bool isNew = false;
    HBRUSH colorBrush = nullptr;
    const std::vector<SavedGroup>* existingGroups = nullptr;
};

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

void UpdateGradientControlsEnabled(HWND hwnd, bool backgroundEnabled, bool fontEnabled) {
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_BG_LABEL), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_BG_SLIDER), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_BG_VALUE), backgroundEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_FONT_LABEL), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_FONT_SLIDER), fontEnabled);
    EnableWindow(GetDlgItem(hwnd, IDC_MAIN_BREADCRUMB_FONT_VALUE), fontEnabled);
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
            }
            return TRUE;
        }
        case WM_CTLCOLORDLG: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (dc) {
                SetBkColor(dc, GetSysColor(COLOR_WINDOW));
            }
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_WINDOW));
        }
        case WM_COMMAND: {
            switch (LOWORD(wParam)) {
                case IDC_MAIN_REOPEN:
                case IDC_MAIN_PERSIST:
                    if (HIWORD(wParam) == BN_CLICKED) {
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
                UpdatePercentageLabel(hwnd, IDC_MAIN_BREADCRUMB_BG_VALUE,
                                     data->workingOptions.breadcrumbGradientTransparency);
                UpdatePercentageLabel(hwnd, IDC_MAIN_BREADCRUMB_FONT_VALUE,
                                     data->workingOptions.breadcrumbFontBrightness);
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
            }
            return TRUE;
        }
        case WM_CTLCOLORDLG: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (dc) {
                SetBkColor(dc, GetSysColor(COLOR_WINDOW));
            }
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_WINDOW));
        }
        case WM_COMMAND: {
            switch (LOWORD(wParam)) {
                case IDC_MAIN_BREADCRUMB:
                case IDC_MAIN_BREADCRUMB_FONT:
                case IDC_MAIN_BREADCRUMB_BG_CUSTOM:
                case IDC_MAIN_BREADCRUMB_FONT_CUSTOM:
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
                case IDC_MAIN_TAB_SELECTED_BUTTON:
                case IDC_MAIN_TAB_UNSELECTED_BUTTON:
                    if (HIWORD(wParam) == BN_CLICKED) {
                        auto* data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        if (HandleColorButtonClick(hwnd, data, LOWORD(wParam))) {
                            SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
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
                        SetBkColor(dc, GetSysColor(COLOR_WINDOW));
                        return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_WINDOW));
                    }
                }
            }
            SetBkMode(dc, TRANSPARENT);
            SetBkColor(dc, GetSysColor(COLOR_WINDOW));
            return reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_WINDOW));
            break;
        }
        case WM_HSCROLL: {
            HWND slider = reinterpret_cast<HWND>(lParam);
            if (!slider) {
                return TRUE;
            }
            const int controlId = GetDlgCtrlID(slider);
            if (controlId == IDC_MAIN_BREADCRUMB_BG_SLIDER || controlId == IDC_MAIN_BREADCRUMB_FONT_SLIDER) {
                const int labelId = (controlId == IDC_MAIN_BREADCRUMB_BG_SLIDER)
                                        ? IDC_MAIN_BREADCRUMB_BG_VALUE
                                        : IDC_MAIN_BREADCRUMB_FONT_VALUE;
                const int sliderValue = ClampPercentageValue(
                    static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0)));
                const int displayValue =
                    (controlId == IDC_MAIN_BREADCRUMB_FONT_SLIDER) ? InvertPercentageValue(sliderValue)
                                                                   : sliderValue;
                UpdatePercentageLabel(hwnd, labelId, displayValue);
                SendMessageW(GetParent(hwnd), PSM_CHANGED, reinterpret_cast<WPARAM>(hwnd), 0);
            }
            return TRUE;
        }
        case WM_NOTIFY: {
            if (((LPNMHDR)lParam)->code == PSN_APPLY) {
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
                    data->workingOptions.useCustomBreadcrumbGradientColors =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_BG_CUSTOM) == BST_CHECKED;
                    data->workingOptions.useCustomBreadcrumbFontColors =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_BREADCRUMB_FONT_CUSTOM) == BST_CHECKED;
                    data->workingOptions.useCustomTabSelectedColor =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_SELECTED_CHECK) == BST_CHECKED;
                    data->workingOptions.useCustomTabUnselectedColor =
                        IsDlgButtonChecked(hwnd, IDC_MAIN_TAB_UNSELECTED_CHECK) == BST_CHECKED;
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
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_TAB_CLASSES | ICC_BAR_CLASSES};
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
    } else {
        result.saved = false;
        result.groupsChanged = false;
        result.optionsChanged = false;
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
    if (data.tabSelectedBrush) {
        DeleteObject(data.tabSelectedBrush);
        data.tabSelectedBrush = nullptr;
    }
    if (data.tabUnselectedBrush) {
        DeleteObject(data.tabUnselectedBrush);
        data.tabUnselectedBrush = nullptr;
    }
    return result;
}

}  // namespace shelltabs
