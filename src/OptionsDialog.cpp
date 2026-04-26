#include "OptionsDialog.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <prsht.h>
#include <commdlg.h>
#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <wrl/client.h>
#include <objbase.h>

#include <algorithm>
#include <array>
#include <cassert>
#include <cstring>
#include <limits>
#include <malloc.h>
#include <memory>
#include <random>
#include <string>
#include <thread>
#include <utility>
#include <vector>

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

// Modern, clean dimensions with proper spacing
constexpr int kPageWidth = 440;
constexpr int kPageHeight = 480;
constexpr int kMargin = 16;
constexpr int kSpacing = 10;
constexpr int kGroupMargin = 12;
constexpr int kLabelHeight = 14;
constexpr int kEditHeight = 22;
constexpr int kButtonHeight = 24;
constexpr int kCheckHeight = 16;
constexpr int kComboHeight = 200;
constexpr int kSliderHeight = 24;
constexpr int kPreviewSize = 64;
constexpr int kColorBoxSize = 28;

constexpr UINT WM_PREVIEW_BITMAP_READY = WM_APP + 101;

// Modern control IDs - clearly organized by page
enum ControlIds : int {
    // General Page (6000-6099)
    IDC_GEN_REOPEN = 6001,
    IDC_GEN_PERSIST = 6002,
    IDC_GEN_DOCK_LABEL = 6003,
    IDC_GEN_DOCK_COMBO = 6004,
    IDC_GEN_NEWTAB_LABEL = 6005,
    IDC_GEN_NEWTAB_COMBO = 6006,
    IDC_GEN_NEWTAB_PATH_LABEL = 6007,
    IDC_GEN_NEWTAB_PATH = 6008,
    IDC_GEN_NEWTAB_PATH_BROWSE = 6009,
    IDC_GEN_NEWTAB_GROUP_LABEL = 6010,
    IDC_GEN_NEWTAB_GROUP = 6011,
    IDC_GEN_REUSE_WINDOW = 6012,

    // Appearance Page (6100-6199)
    IDC_APP_BREADCRUMB_GROUP = 6100,
    IDC_APP_BREADCRUMB_ENABLE = 6101,
    IDC_APP_BREADCRUMB_TRANS_LABEL = 6102,
    IDC_APP_BREADCRUMB_TRANS = 6103,
    IDC_APP_BREADCRUMB_TRANS_VAL = 6104,
    IDC_APP_BREADCRUMB_CUSTOM = 6105,
    IDC_APP_BREADCRUMB_START_LABEL = 6106,
    IDC_APP_BREADCRUMB_START_PREVIEW = 6107,
    IDC_APP_BREADCRUMB_START_BTN = 6108,
    IDC_APP_BREADCRUMB_END_LABEL = 6109,
    IDC_APP_BREADCRUMB_END_PREVIEW = 6110,
    IDC_APP_BREADCRUMB_END_BTN = 6111,

    IDC_APP_FONT_GROUP = 6120,
    IDC_APP_FONT_ENABLE = 6121,
    IDC_APP_FONT_BRIGHT_LABEL = 6122,
    IDC_APP_FONT_BRIGHT = 6123,
    IDC_APP_FONT_BRIGHT_VAL = 6124,
    IDC_APP_FONT_CUSTOM = 6125,
    IDC_APP_FONT_START_LABEL = 6126,
    IDC_APP_FONT_START_PREVIEW = 6127,
    IDC_APP_FONT_START_BTN = 6128,
    IDC_APP_FONT_END_LABEL = 6129,
    IDC_APP_FONT_END_PREVIEW = 6130,
    IDC_APP_FONT_END_BTN = 6131,

    IDC_APP_HIGHLIGHT_LABEL = 6140,
    IDC_APP_HIGHLIGHT = 6141,
    IDC_APP_HIGHLIGHT_VAL = 6142,
    IDC_APP_DROPDOWN_LABEL = 6143,
    IDC_APP_DROPDOWN = 6144,
    IDC_APP_DROPDOWN_VAL = 6145,

    IDC_APP_TAB_GROUP = 6150,
    IDC_APP_TAB_SEL_CHECK = 6151,
    IDC_APP_TAB_SEL_PREVIEW = 6152,
    IDC_APP_TAB_SEL_BTN = 6153,
    IDC_APP_TAB_UNSEL_CHECK = 6154,
    IDC_APP_TAB_UNSEL_PREVIEW = 6155,
    IDC_APP_TAB_UNSEL_BTN = 6156,

    IDC_APP_PROGRESS_GROUP = 6160,
    IDC_APP_PROGRESS_CUSTOM = 6161,
    IDC_APP_PROGRESS_START_LABEL = 6162,
    IDC_APP_PROGRESS_START_PREVIEW = 6163,
    IDC_APP_PROGRESS_START_BTN = 6164,
    IDC_APP_PROGRESS_END_LABEL = 6165,
    IDC_APP_PROGRESS_END_PREVIEW = 6166,
    IDC_APP_PROGRESS_END_BTN = 6167,
    IDC_APP_LISTVIEW_ACCENT = 6168,

    // Glow Effects (6200-6299)
    IDC_GLOW_ENABLE = 6200,
    IDC_GLOW_USE_GRADIENT = 6201,
    IDC_GLOW_CUSTOM = 6202,
    IDC_GLOW_PRIMARY_LABEL = 6203,
    IDC_GLOW_PRIMARY_PREVIEW = 6204,
    IDC_GLOW_PRIMARY_BTN = 6205,
    IDC_GLOW_SECONDARY_LABEL = 6206,
    IDC_GLOW_SECONDARY_PREVIEW = 6207,
    IDC_GLOW_SECONDARY_BTN = 6208,
    IDC_GLOW_BITMAP_INTERCEPT = 6209,
    IDC_GLOW_FILE_GRADIENT = 6210,
    IDC_GLOW_EXPLORER_ACCENT = 6211,

    IDC_GLOW_SURFACES = 6220,
    IDC_GLOW_SURF_LISTVIEW = 6221,
    IDC_GLOW_SURF_HEADER = 6222,
    IDC_GLOW_SURF_REBAR = 6223,
    IDC_GLOW_SURF_TOOLBAR = 6224,
    IDC_GLOW_SURF_EDIT = 6225,
    IDC_GLOW_SURF_DIRECTUI = 6226,
    IDC_GLOW_SURF_SCROLLBAR = 6227,

    // Backgrounds (6300-6399)
    IDC_BG_ENABLE = 6300,
    IDC_BG_UNIVERSAL_GROUP = 6301,
    IDC_BG_UNIVERSAL_PREVIEW = 6302,
    IDC_BG_UNIVERSAL_BROWSE = 6303,
    IDC_BG_UNIVERSAL_CLEAR = 6304,
    IDC_BG_UNIVERSAL_NAME = 6305,
    IDC_BG_FOLDER_GROUP = 6310,
    IDC_BG_FOLDER_LIST = 6311,
    IDC_BG_FOLDER_ADD = 6312,
    IDC_BG_FOLDER_EDIT = 6313,
    IDC_BG_FOLDER_REMOVE = 6314,
    IDC_BG_FOLDER_PREVIEW = 6315,
    IDC_BG_FOLDER_NAME = 6316,
    IDC_BG_FOLDER_CLEAN = 6317,
    IDC_BG_OPACITY_LABEL = 6320,
    IDC_BG_OPACITY_SLIDER = 6321,
    IDC_BG_OPACITY_VAL = 6322,
    IDC_BG_POS_GROUP = 6330,
    IDC_BG_POS_TILE = 6331,
    IDC_BG_POS_STRETCH = 6332,
    IDC_BG_POS_CENTER = 6333,
    IDC_BG_POS_BOTTOMLEFT = 6334,
    IDC_BG_POS_BOTTOMRIGHT = 6335,

    // Context Menus (6400-6499)
    IDC_CTX_TREE = 6400,
    IDC_CTX_TEMPLATE = 6401,
    IDC_CTX_ADD_COMMAND = 6402,
    IDC_CTX_ADD_SUBMENU = 6403,
    IDC_CTX_ADD_SEPARATOR = 6404,
    IDC_CTX_REMOVE = 6405,
    IDC_CTX_MOVE_UP = 6406,
    IDC_CTX_MOVE_DOWN = 6407,
    IDC_CTX_INDENT = 6408,
    IDC_CTX_OUTDENT = 6409,

    IDC_CTX_LABEL = 6420,
    IDC_CTX_ICON = 6421,
    IDC_CTX_ICON_BROWSE = 6422,
    IDC_CTX_COMMAND = 6423,
    IDC_CTX_COMMAND_BROWSE = 6424,
    IDC_CTX_ARGS = 6425,
    IDC_CTX_WORKDIR = 6426,
    IDC_CTX_WORKDIR_BROWSE = 6427,
    IDC_CTX_RUN_ADMIN = 6428,
    IDC_CTX_WAIT = 6429,
    IDC_CTX_WINDOW_STATE = 6430,
    IDC_CTX_ENABLED = 6431,
    IDC_CTX_DESCRIPTION = 6432,
    IDC_CTX_ID = 6433,

    IDC_CTX_MIN_SEL = 6440,
    IDC_CTX_MAX_SEL = 6441,
    IDC_CTX_FILES = 6442,
    IDC_CTX_FOLDERS = 6443,
    IDC_CTX_MULTIPLE = 6444,
    IDC_CTX_PATTERNS = 6445,
    IDC_CTX_EXCLUDE = 6446,
    IDC_CTX_ANCHOR = 6447,

    IDC_CTX_FOLDER_FILTERS = 6450,
    IDC_CTX_CONFIRM = 6451,
    IDC_CTX_CONFIRM_MSG = 6452,
    IDC_CTX_ADDL_CMDS = 6453,
    IDC_CTX_ADDL_ADD = 6454,
    IDC_CTX_ADDL_REMOVE = 6455,
    IDC_CTX_EXPAND_ENV = 6456,
    IDC_CTX_PROPS_PANEL = 6457,

    // Groups (6500-6599)
    IDC_GRP_LIST = 6500,
    IDC_GRP_NEW = 6501,
    IDC_GRP_EDIT = 6502,
    IDC_GRP_REMOVE = 6503,

    IDC_GRP_ED_NAME = 6510,
    IDC_GRP_ED_PATHS = 6511,
    IDC_GRP_ED_ADD = 6512,
    IDC_GRP_ED_EDIT_PATH = 6513,
    IDC_GRP_ED_REMOVE_PATH = 6514,
    IDC_GRP_ED_COLOR_PREVIEW = 6515,
    IDC_GRP_ED_COLOR_BTN = 6516,
    IDC_GRP_ED_STYLE = 6517,

    // Web Folders (6600-6699)
    IDC_WEB_LIST = 6600,
    IDC_WEB_ADD = 6601,
    IDC_WEB_EDIT = 6602,
    IDC_WEB_REMOVE = 6603,
    IDC_WEB_URL_LABEL = 6604,

    IDC_WEB_ED_NAME = 6610,
    IDC_WEB_ED_URL = 6611,
    IDC_WEB_ED_PARALLEL = 6612,
    IDC_WEB_ED_MAX_CONCURRENT = 6613,
    IDC_WEB_ED_SPEED_LIMIT = 6614,
};

// Preview bitmap async result
struct PreviewBitmapResult {
    UINT64 token = 0;
    HBITMAP bitmap = nullptr;
};

// Main options dialog data
struct OptionsDialogData {
    ShellTabsOptions originalOptions;
    ShellTabsOptions workingOptions;

    std::vector<SavedGroup> originalGroups;
    std::vector<SavedGroup> workingGroups;
    std::vector<std::wstring> workingGroupIds;
    std::vector<std::wstring> removedGroupIds;

    bool applyInvoked = false;
    bool groupsChanged = false;
    bool previewBroadcasted = false;
    int initialTab = 0;

    // Color brushes for previews
    HBRUSH breadcrumbStartBrush = nullptr;
    HBRUSH breadcrumbEndBrush = nullptr;
    HBRUSH fontStartBrush = nullptr;
    HBRUSH fontEndBrush = nullptr;
    HBRUSH tabSelectedBrush = nullptr;
    HBRUSH tabUnselectedBrush = nullptr;
    HBRUSH progressStartBrush = nullptr;
    HBRUSH progressEndBrush = nullptr;
    HBRUSH glowPrimaryBrush = nullptr;
    HBRUSH glowSecondaryBrush = nullptr;
    HBRUSH groupColorBrush = nullptr;

    // Background previews
    HBITMAP universalPreview = nullptr;
    HBITMAP folderPreview = nullptr;
    UINT64 universalPreviewToken = 0;
    UINT64 folderPreviewToken = 0;
    std::wstring lastImageDir;
    std::wstring lastFolderPath;
    std::vector<std::wstring> createdCachedImages;
    std::vector<std::wstring> pendingCachedRemovals;

    // Context menu tree state
    std::vector<std::vector<size_t>> contextTreePaths;
    std::vector<HTREEITEM> contextTreeItems;
    std::vector<size_t> contextSelection;
    bool contextSelectionValid = false;
    bool contextUpdating = false;
    std::wstring contextCmdBrowseDir;

    // Focus handling
    std::wstring focusGroupId;
    bool focusGroupEdit = false;
    bool focusHandled = false;
};

//=============================================================================
// Dialog Template Builder - Modern, clean helper functions
//=============================================================================

void AlignBuffer(std::vector<BYTE>& buffer) {
    while (buffer.size() % 4 != 0) {
        buffer.push_back(0);
    }
}

void AppendWord(std::vector<BYTE>& buffer, WORD value) {
    buffer.push_back(static_cast<BYTE>(value & 0xFF));
    buffer.push_back(static_cast<BYTE>((value >> 8) & 0xFF));
}

void AppendDWord(std::vector<BYTE>& buffer, DWORD value) {
    buffer.push_back(static_cast<BYTE>(value & 0xFF));
    buffer.push_back(static_cast<BYTE>((value >> 8) & 0xFF));
    buffer.push_back(static_cast<BYTE>((value >> 16) & 0xFF));
    buffer.push_back(static_cast<BYTE>((value >> 24) & 0xFF));
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

DialogTemplatePtr AllocateTemplate(const std::vector<BYTE>& source) {
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

// Modern dialog template builder class
class DialogBuilder {
public:
    DialogBuilder(int width, int height, int controlCount = 0) {
        data_.resize(sizeof(DLGTEMPLATE));
        auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data_.data());
        dlg->style = DS_SETFONT | DS_CONTROL | WS_CHILD | WS_VISIBLE |
                     WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        dlg->dwExtendedStyle = WS_EX_CONTROLPARENT;
        dlg->cdit = static_cast<WORD>(controlCount);
        dlg->x = 0;
        dlg->y = 0;
        dlg->cx = static_cast<short>(width);
        dlg->cy = static_cast<short>(height);

        AppendWord(data_, 0);  // menu
        AppendWord(data_, 0);  // class
        AppendWord(data_, 0);  // title
        AppendWord(data_, 9);  // font size
        AppendString(data_, L"Segoe UI");
    }

    void AddButton(int id, const wchar_t* text, int x, int y, int w, int h,
                   DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                   DWORD exStyle = 0) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = style;
        item->dwExtendedStyle = exStyle;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendWord(data_, 0xFFFF);
        AppendWord(data_, 0x0080);  // BUTTON
        AppendString(data_, text);
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void AddCheckbox(int id, const wchar_t* text, int x, int y, int w, int h) {
        AddButton(id, text, x, y, w, h,
                 WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);
    }

    void AddRadioButton(int id, const wchar_t* text, int x, int y, int w, int h, bool isGroupStart = false) {
        DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTORADIOBUTTON;
        if (isGroupStart) style |= WS_GROUP;
        AddButton(id, text, x, y, w, h, style);
    }

    void AddPushButton(int id, const wchar_t* text, int x, int y, int w, int h) {
        AddButton(id, text, x, y, w, h,
                 WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON);
    }

    void AddGroupBox(int id, const wchar_t* text, int x, int y, int w, int h) {
        AddButton(id, text, x, y, w, h,
                 WS_CHILD | WS_VISIBLE | BS_GROUPBOX);
    }

    void AddStatic(int id, const wchar_t* text, int x, int y, int w, int h,
                   DWORD style = WS_CHILD | WS_VISIBLE) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = style;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendWord(data_, 0xFFFF);
        AppendWord(data_, 0x0082);  // STATIC
        AppendString(data_, text);
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void AddEdit(int id, const wchar_t* text, int x, int y, int w, int h,
                DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_LEFT) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = style;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendWord(data_, 0xFFFF);
        AppendWord(data_, 0x0081);  // EDIT
        AppendString(data_, text ? text : L"");
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void AddComboBox(int id, int x, int y, int w, int h) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendWord(data_, 0xFFFF);
        AppendWord(data_, 0x0085);  // COMBOBOX
        AppendWord(data_, 0);
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void AddListBox(int id, int x, int y, int w, int h,
                   DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER |
                                 WS_VSCROLL | LBS_NOTIFY) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = style;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendWord(data_, 0xFFFF);
        AppendWord(data_, 0x0083);  // LISTBOX
        AppendWord(data_, 0);
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void AddSlider(int id, int x, int y, int w, int h) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | TBS_AUTOTICKS;
        item->dwExtendedStyle = 0;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendString(data_, L"msctls_trackbar32");
        AppendWord(data_, 0);
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void AddTreeView(int id, int x, int y, int w, int h) {
        AlignBuffer(data_);
        size_t offset = data_.size();
        data_.resize(offset + sizeof(DLGITEMTEMPLATE));
        auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(data_.data() + offset);
        item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | WS_VSCROLL |
                      TVS_HASLINES | TVS_HASBUTTONS | TVS_LINESATROOT | TVS_SHOWSELALWAYS;
        item->dwExtendedStyle = WS_EX_CLIENTEDGE;
        item->x = static_cast<short>(x);
        item->y = static_cast<short>(y);
        item->cx = static_cast<short>(w);
        item->cy = static_cast<short>(h);
        item->id = static_cast<WORD>(id);
        AppendString(data_, L"SysTreeView32");
        AppendWord(data_, 0);
        AppendWord(data_, 0);
        IncrementControlCount();
    }

    void EnableVerticalScrollbar() {
        auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data_.data());
        dlg->style |= WS_VSCROLL;
    }

    DialogTemplatePtr Build() {
        return AllocateTemplate(data_);
    }

private:
    void IncrementControlCount() {
        auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data_.data());
        dlg->cdit++;
    }

    std::vector<BYTE> data_;
};

//=============================================================================
// Helper Functions
//=============================================================================

std::wstring GetControlText(HWND hwnd) {
    if (!hwnd) return L"";
    int len = GetWindowTextLengthW(hwnd);
    if (len <= 0) return L"";
    std::wstring text(static_cast<size_t>(len) + 1, L'\0');
    int copied = GetWindowTextW(hwnd, text.data(), len + 1);
    if (copied > 0) {
        text.resize(static_cast<size_t>(copied));
    } else {
        text.clear();
    }
    return text;
}

void UpdateBrush(HBRUSH* brush, COLORREF color) {
    if (brush && *brush) {
        DeleteObject(*brush);
    }
    if (brush) {
        *brush = CreateSolidBrush(color);
    }
}

bool ChooseColor(HWND parent, COLORREF* color) {
    if (!color) return false;

    static COLORREF customColors[16] = {
        RGB(255,255,255), RGB(0,0,0), RGB(255,0,0), RGB(0,255,0),
        RGB(0,0,255), RGB(255,255,0), RGB(255,0,255), RGB(0,255,255),
        RGB(192,192,192), RGB(128,128,128), RGB(128,0,0), RGB(0,128,0),
        RGB(0,0,128), RGB(128,128,0), RGB(128,0,128), RGB(0,128,128)
    };

    CHOOSECOLORW cc{};
    cc.lStructSize = sizeof(cc);
    cc.hwndOwner = parent;
    cc.rgbResult = *color;
    cc.lpCustColors = customColors;
    cc.Flags = CC_FULLOPEN | CC_RGBINIT;

    if (ChooseColorW(&cc)) {
        *color = cc.rgbResult;
        return true;
    }
    return false;
}

void DrawColorBox(HDC hdc, RECT rect, COLORREF color) {
    HBRUSH brush = CreateSolidBrush(color);
    FillRect(hdc, &rect, brush);
    DeleteObject(brush);
    FrameRect(hdc, &rect, static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
}

bool BrowseForFolderImpl(HWND parent, std::wstring* path, const wchar_t* title, bool allowVirtual) {
    if (!path) return false;

    Microsoft::WRL::ComPtr<IFileDialog> dialog;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&dialog));
    if (FAILED(hr)) return false;

    DWORD options = 0;
    dialog->GetOptions(&options);
    options |= FOS_PICKFOLDERS;
    if (!allowVirtual) {
        options |= FOS_FORCEFILESYSTEM;
    }
    dialog->SetOptions(options);

    if (title) dialog->SetTitle(title);

    if (SUCCEEDED(dialog->Show(parent))) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (SUCCEEDED(dialog->GetResult(&item))) {
            if (TryGetShellLocationPath(item.Get(), path, allowVirtual)) {
                return true;
            }
        }
    }
    return false;
}

bool BrowseForFolder(HWND parent, std::wstring* path, const wchar_t* title = nullptr) {
    return BrowseForFolderImpl(parent, path, title, false);
}

bool BrowseForImage(HWND parent, std::wstring* path, std::wstring* dir) {
    if (!path) return false;

    Microsoft::WRL::ComPtr<IFileDialog> dialog;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&dialog));
    if (FAILED(hr)) return false;

    COMDLG_FILTERSPEC filters[] = {
        { L"Images", L"*.jpg;*.jpeg;*.png;*.bmp;*.gif" },
        { L"All Files", L"*.*" }
    };
    dialog->SetFileTypes(2, filters);
    dialog->SetTitle(L"Select Image");

    if (dir && !dir->empty()) {
        Microsoft::WRL::ComPtr<IShellItem> folder;
        if (SUCCEEDED(SHCreateItemFromParsingName(dir->c_str(), nullptr,
                                                  IID_PPV_ARGS(&folder)))) {
            dialog->SetFolder(folder.Get());
        }
    }

    if (SUCCEEDED(dialog->Show(parent))) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (SUCCEEDED(dialog->GetResult(&item))) {
            wchar_t* filePath = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath))) {
                *path = filePath;
                if (dir) {
                    *dir = *path;
                    size_t pos = dir->find_last_of(L"\\");
                    if (pos != std::wstring::npos) {
                        dir->resize(pos);
                    }
                }
                CoTaskMemFree(filePath);
                return true;
            }
        }
    }
    return false;
}

//=============================================================================
// GENERAL PAGE
//=============================================================================

DialogTemplatePtr CreateGeneralPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    int y = kMargin;

    // Crash recovery
    builder.AddCheckbox(IDC_GEN_REOPEN,
        L"Reopen tabs after Explorer crash",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing;

    // Persist groups
    builder.AddCheckbox(IDC_GEN_PERSIST,
        L"Remember folder paths in saved groups",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing;

    // Reuse existing window
    builder.AddCheckbox(IDC_GEN_REUSE_WINDOW,
        L"Open folders as tabs in an existing window",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing * 2;

    // Dock mode
    builder.AddStatic(IDC_GEN_DOCK_LABEL,
        L"Tab band position:",
        kMargin, y, 120, kLabelHeight);
    y += kLabelHeight + 4;

    builder.AddComboBox(IDC_GEN_DOCK_COMBO,
        kMargin, y, 200, kComboHeight);
    y += kEditHeight + kSpacing * 2;

    // New tab behavior
    builder.AddStatic(IDC_GEN_NEWTAB_LABEL,
        L"New tab opens:",
        kMargin, y, 120, kLabelHeight);
    y += kLabelHeight + 4;

    builder.AddComboBox(IDC_GEN_NEWTAB_COMBO,
        kMargin, y, 200, kComboHeight);
    y += kEditHeight + kSpacing;

    // Custom path (shown conditionally)
    builder.AddStatic(IDC_GEN_NEWTAB_PATH_LABEL,
        L"Custom path:",
        kMargin + kGroupMargin, y, 100, kLabelHeight);
    builder.AddEdit(IDC_GEN_NEWTAB_PATH, L"",
        kMargin + kGroupMargin + 100, y - 2, 180, kEditHeight);
    builder.AddPushButton(IDC_GEN_NEWTAB_PATH_BROWSE, L"Browse...",
        kMargin + kGroupMargin + 285, y - 2, 70, kEditHeight);
    y += kEditHeight + kSpacing;

    // Saved group (shown conditionally)
    builder.AddStatic(IDC_GEN_NEWTAB_GROUP_LABEL,
        L"Saved group:",
        kMargin + kGroupMargin, y, 100, kLabelHeight);
    builder.AddComboBox(IDC_GEN_NEWTAB_GROUP,
        kMargin + kGroupMargin + 100, y - 2, 255, kComboHeight);

    return builder.Build();
}

void InitGeneralPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    // Set checkbox states
    CheckDlgButton(page, IDC_GEN_REOPEN,
        data->workingOptions.reopenOnCrash ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GEN_PERSIST,
        data->workingOptions.persistGroupPaths ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GEN_REUSE_WINDOW,
        data->workingOptions.reuseExistingWindow ? BST_CHECKED : BST_UNCHECKED);

    // Populate dock mode combo
    HWND dockCombo = GetDlgItem(page, IDC_GEN_DOCK_COMBO);
    if (dockCombo) {
        const wchar_t* modes[] = { L"Automatic", L"Top", L"Bottom", L"Left", L"Right" };
        for (const auto* mode : modes) {
            ComboBox_AddString(dockCombo, mode);
        }
        ComboBox_SetCurSel(dockCombo, static_cast<int>(data->workingOptions.tabDockMode));
    }

    // Populate new tab combo
    HWND newTabCombo = GetDlgItem(page, IDC_GEN_NEWTAB_COMBO);
    if (newTabCombo) {
        const wchar_t* templates[] = {
            L"Duplicate current tab",
            L"This PC",
            L"Custom path",
            L"Saved group"
        };
        for (const auto* tmpl : templates) {
            ComboBox_AddString(newTabCombo, tmpl);
        }
        ComboBox_SetCurSel(newTabCombo, static_cast<int>(data->workingOptions.newTabTemplate));
    }

    // Set custom path
    SetDlgItemTextW(page, IDC_GEN_NEWTAB_PATH, data->workingOptions.newTabCustomPath.c_str());

    // Populate groups combo
    HWND groupCombo = GetDlgItem(page, IDC_GEN_NEWTAB_GROUP);
    if (groupCombo) {
        for (const auto& group : data->workingGroups) {
            int idx = ComboBox_AddString(groupCombo, group.name.c_str());
            if (group.name == data->workingOptions.newTabSavedGroup) {
                ComboBox_SetCurSel(groupCombo, idx);
            }
        }
    }

    // Show/hide controls based on selection
    bool showPath = (data->workingOptions.newTabTemplate == NewTabTemplate::kCustomPath);
    bool showGroup = (data->workingOptions.newTabTemplate == NewTabTemplate::kSavedGroup);

    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_PATH_LABEL), showPath ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_PATH), showPath ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_PATH_BROWSE), showPath ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_GROUP_LABEL), showGroup ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_GROUP), showGroup ? SW_SHOW : SW_HIDE);
}

void UpdateGeneralPageVisibility(HWND page, NewTabTemplate tmpl) {
    bool showPath = (tmpl == NewTabTemplate::kCustomPath);
    bool showGroup = (tmpl == NewTabTemplate::kSavedGroup);

    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_PATH_LABEL), showPath ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_PATH), showPath ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_PATH_BROWSE), showPath ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_GROUP_LABEL), showGroup ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(page, IDC_GEN_NEWTAB_GROUP), showGroup ? SW_SHOW : SW_HIDE);
}

INT_PTR CALLBACK GeneralPageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitGeneralPage(page, data);
        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_GEN_REOPEN && code == BN_CLICKED) {
                data->workingOptions.reopenOnCrash =
                    (IsDlgButtonChecked(page, IDC_GEN_REOPEN) == BST_CHECKED);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_GEN_PERSIST && code == BN_CLICKED) {
                data->workingOptions.persistGroupPaths =
                    (IsDlgButtonChecked(page, IDC_GEN_PERSIST) == BST_CHECKED);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_GEN_REUSE_WINDOW && code == BN_CLICKED) {
                data->workingOptions.reuseExistingWindow =
                    (IsDlgButtonChecked(page, IDC_GEN_REUSE_WINDOW) == BST_CHECKED);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_GEN_DOCK_COMBO && code == CBN_SELCHANGE) {
                HWND combo = GetDlgItem(page, IDC_GEN_DOCK_COMBO);
                int sel = ComboBox_GetCurSel(combo);
                if (sel >= 0) {
                    data->workingOptions.tabDockMode = static_cast<TabBandDockMode>(sel);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_GEN_NEWTAB_COMBO && code == CBN_SELCHANGE) {
                HWND combo = GetDlgItem(page, IDC_GEN_NEWTAB_COMBO);
                int sel = ComboBox_GetCurSel(combo);
                if (sel >= 0) {
                    data->workingOptions.newTabTemplate = static_cast<NewTabTemplate>(sel);
                    UpdateGeneralPageVisibility(page, data->workingOptions.newTabTemplate);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_GEN_NEWTAB_PATH && code == EN_CHANGE) {
                data->workingOptions.newTabCustomPath = GetControlText(GetDlgItem(page, IDC_GEN_NEWTAB_PATH));
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_GEN_NEWTAB_PATH_BROWSE && code == BN_CLICKED) {
                std::wstring path;
                if (BrowseForFolder(page, &path, L"Select Folder")) {
                    data->workingOptions.newTabCustomPath = path;
                    SetDlgItemTextW(page, IDC_GEN_NEWTAB_PATH, path.c_str());
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_GEN_NEWTAB_GROUP && code == CBN_SELCHANGE) {
                HWND combo = GetDlgItem(page, IDC_GEN_NEWTAB_GROUP);
                int sel = ComboBox_GetCurSel(combo);
                if (sel >= 0 && sel < static_cast<int>(data->workingGroups.size())) {
                    data->workingOptions.newTabSavedGroup = data->workingGroups[sel].name;
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            break;
        }

        case WM_NOTIFY: {
            NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);
            if (nmhdr->code == PSN_APPLY && data) {
                return PSNRET_NOERROR;
            }
            break;
        }
    }

    return FALSE;
}

//=============================================================================
// APPEARANCE PAGE - Simplified placeholder
//=============================================================================

DialogTemplatePtr CreateAppearancePageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    int y = kMargin;

    // Breadcrumb section
    builder.AddGroupBox(IDC_APP_BREADCRUMB_GROUP, L"Breadcrumb Bar",
        kMargin, y, kPageWidth - 2 * kMargin, 100);
    y += 18;

    builder.AddCheckbox(IDC_APP_BREADCRUMB_ENABLE,
        L"Enable gradient background",
        kMargin + kGroupMargin, y, 220, kCheckHeight);
    y += kCheckHeight + kSpacing;

    builder.AddStatic(IDC_APP_BREADCRUMB_TRANS_LABEL,
        L"Transparency:",
        kMargin + kGroupMargin, y, 100, kLabelHeight);
    builder.AddSlider(IDC_APP_BREADCRUMB_TRANS,
        kMargin + kGroupMargin + 100, y - 2, 200, kSliderHeight);
    builder.AddStatic(IDC_APP_BREADCRUMB_TRANS_VAL, L"45%",
        kMargin + kGroupMargin + 310, y, 40, kLabelHeight);

    y += kSliderHeight + kSpacing * 2;

    // Tab colors section
    builder.AddGroupBox(IDC_APP_TAB_GROUP, L"Tab Colors",
        kMargin, y, kPageWidth - 2 * kMargin, 80);
    y += 18;

    builder.AddCheckbox(IDC_APP_TAB_SEL_CHECK,
        L"Custom selected color:",
        kMargin + kGroupMargin, y, 150, kCheckHeight);
    builder.AddStatic(IDC_APP_TAB_SEL_PREVIEW, L"",
        kMargin + kGroupMargin + 160, y - 2, kColorBoxSize, kColorBoxSize,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW);
    builder.AddPushButton(IDC_APP_TAB_SEL_BTN, L"Choose...",
        kMargin + kGroupMargin + 200, y - 2, 70, kButtonHeight);
    y += kButtonHeight + kSpacing;

    builder.AddCheckbox(IDC_APP_TAB_UNSEL_CHECK,
        L"Custom unselected color:",
        kMargin + kGroupMargin, y, 150, kCheckHeight);
    builder.AddStatic(IDC_APP_TAB_UNSEL_PREVIEW, L"",
        kMargin + kGroupMargin + 160, y - 2, kColorBoxSize, kColorBoxSize,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW);
    builder.AddPushButton(IDC_APP_TAB_UNSEL_BTN, L"Choose...",
        kMargin + kGroupMargin + 200, y - 2, 70, kButtonHeight);

    return builder.Build();
}

void InitAppearancePage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    CheckDlgButton(page, IDC_APP_BREADCRUMB_ENABLE,
        data->workingOptions.enableBreadcrumbGradient ? BST_CHECKED : BST_UNCHECKED);

    HWND slider = GetDlgItem(page, IDC_APP_BREADCRUMB_TRANS);
    if (slider) {
        SendMessageW(slider, TBM_SETRANGE, TRUE, MAKELPARAM(0, 100));
        SendMessageW(slider, TBM_SETPOS, TRUE, data->workingOptions.breadcrumbGradientTransparency);
    }

    wchar_t buf[32];
    swprintf_s(buf, L"%d%%", data->workingOptions.breadcrumbGradientTransparency);
    SetDlgItemTextW(page, IDC_APP_BREADCRUMB_TRANS_VAL, buf);

    CheckDlgButton(page, IDC_APP_TAB_SEL_CHECK,
        data->workingOptions.useCustomTabSelectedColor ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_APP_TAB_UNSEL_CHECK,
        data->workingOptions.useCustomTabUnselectedColor ? BST_CHECKED : BST_UNCHECKED);

    UpdateBrush(&data->tabSelectedBrush, data->workingOptions.customTabSelectedColor);
    UpdateBrush(&data->tabUnselectedBrush, data->workingOptions.customTabUnselectedColor);
}

INT_PTR CALLBACK AppearancePageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitAppearancePage(page, data);
        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_APP_BREADCRUMB_ENABLE && code == BN_CLICKED) {
                data->workingOptions.enableBreadcrumbGradient =
                    (IsDlgButtonChecked(page, IDC_APP_BREADCRUMB_ENABLE) == BST_CHECKED);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_APP_TAB_SEL_CHECK && code == BN_CLICKED) {
                data->workingOptions.useCustomTabSelectedColor =
                    (IsDlgButtonChecked(page, IDC_APP_TAB_SEL_CHECK) == BST_CHECKED);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_APP_TAB_SEL_BTN && code == BN_CLICKED) {
                if (ChooseColor(page, &data->workingOptions.customTabSelectedColor)) {
                    UpdateBrush(&data->tabSelectedBrush, data->workingOptions.customTabSelectedColor);
                    InvalidateRect(GetDlgItem(page, IDC_APP_TAB_SEL_PREVIEW), nullptr, TRUE);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_APP_TAB_UNSEL_CHECK && code == BN_CLICKED) {
                data->workingOptions.useCustomTabUnselectedColor =
                    (IsDlgButtonChecked(page, IDC_APP_TAB_UNSEL_CHECK) == BST_CHECKED);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_APP_TAB_UNSEL_BTN && code == BN_CLICKED) {
                if (ChooseColor(page, &data->workingOptions.customTabUnselectedColor)) {
                    UpdateBrush(&data->tabUnselectedBrush, data->workingOptions.customTabUnselectedColor);
                    InvalidateRect(GetDlgItem(page, IDC_APP_TAB_UNSEL_PREVIEW), nullptr, TRUE);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            break;
        }

        case WM_HSCROLL: {
            if (!data) break;

            HWND slider = reinterpret_cast<HWND>(lParam);
            int id = GetDlgCtrlID(slider);

            if (id == IDC_APP_BREADCRUMB_TRANS) {
                int pos = static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0));
                data->workingOptions.breadcrumbGradientTransparency = pos;
                wchar_t buf[32];
                swprintf_s(buf, L"%d%%", pos);
                SetDlgItemTextW(page, IDC_APP_BREADCRUMB_TRANS_VAL, buf);
                PropSheet_Changed(GetParent(page), page);
            }
            break;
        }

        case WM_DRAWITEM: {
            if (!data) break;

            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis->CtlType != ODT_STATIC) break;
            if (dis->CtlID == IDC_APP_TAB_SEL_PREVIEW) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.customTabSelectedColor);
                return TRUE;
            }
            else if (dis->CtlID == IDC_APP_TAB_UNSEL_PREVIEW) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.customTabUnselectedColor);
                return TRUE;
            }
            break;
        }
    }

    return FALSE;
}

//=============================================================================
// GLOW EFFECTS PAGE
//=============================================================================

DialogTemplatePtr CreateGlowEffectsPageTemplate() {
    // Use actual content height instead of fixed page height to enable scrolling
    constexpr int kActualContentHeight = 520;  // Increased height to fit all controls
    DialogBuilder builder(kPageWidth, kActualContentHeight);
    builder.EnableVerticalScrollbar();

    int y = kMargin;

    // Main glow controls
    builder.AddCheckbox(IDC_GLOW_ENABLE,
        L"Enable neon glow effects",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing;

    builder.AddCheckbox(IDC_GLOW_BITMAP_INTERCEPT,
        L"Intercept Explorer bitmaps (may impact performance)",
        kMargin, y, 350, kCheckHeight);
    y += kCheckHeight + kSpacing;

    builder.AddCheckbox(IDC_GLOW_FILE_GRADIENT,
        L"Enable file/folder gradient font",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing;

    builder.AddCheckbox(IDC_GLOW_EXPLORER_ACCENT,
        L"Use Explorer accent colors",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing * 2;

    // Surface group
    builder.AddGroupBox(IDC_GLOW_SURFACES, L"Glow Surfaces",
        kMargin, y, kPageWidth - 2 * kMargin, 150);
    y += 18;

    builder.AddCheckbox(IDC_GLOW_SURF_LISTVIEW,
        L"Enable list view glow",
        kMargin + kGroupMargin, y, 200, kCheckHeight);
    y += kCheckHeight + 6;

    builder.AddCheckbox(IDC_GLOW_SURF_HEADER,
        L"Enable column header glow",
        kMargin + kGroupMargin, y, 200, kCheckHeight);
    y += kCheckHeight + 6;

    builder.AddCheckbox(IDC_GLOW_SURF_REBAR,
        L"Enable rebar glow",
        kMargin + kGroupMargin, y, 200, kCheckHeight);
    y += kCheckHeight + 6;

    builder.AddCheckbox(IDC_GLOW_SURF_TOOLBAR,
        L"Enable toolbar glow",
        kMargin + kGroupMargin, y, 200, kCheckHeight);
    y += kCheckHeight + 6;

    builder.AddCheckbox(IDC_GLOW_SURF_EDIT,
        L"Enable address bar glow",
        kMargin + kGroupMargin, y, 200, kCheckHeight);
    y += kCheckHeight + 6;


    builder.AddCheckbox(IDC_GLOW_SURF_SCROLLBAR,
        L"Enable scrollbar glow",
        kMargin + kGroupMargin, y, 200, kCheckHeight);
    y += kCheckHeight + kSpacing * 2;

    // Custom colors
    builder.AddCheckbox(IDC_GLOW_CUSTOM,
        L"Use custom glow colors",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing;

    builder.AddCheckbox(IDC_GLOW_USE_GRADIENT,
        L"Blend glow with gradient",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing;

    // Color pickers
    builder.AddStatic(IDC_GLOW_PRIMARY_LABEL,
        L"Primary color:",
        kMargin, y, 100, kLabelHeight);
    builder.AddStatic(IDC_GLOW_PRIMARY_PREVIEW, L"",
        kMargin + 105, y - 2, kColorBoxSize, kColorBoxSize,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW);
    builder.AddPushButton(IDC_GLOW_PRIMARY_BTN, L"Choose...",
        kMargin + 145, y - 2, 80, kButtonHeight);
    y += kButtonHeight + kSpacing;

    builder.AddStatic(IDC_GLOW_SECONDARY_LABEL,
        L"Secondary color:",
        kMargin, y, 100, kLabelHeight);
    builder.AddStatic(IDC_GLOW_SECONDARY_PREVIEW, L"",
        kMargin + 105, y - 2, kColorBoxSize, kColorBoxSize,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW);
    builder.AddPushButton(IDC_GLOW_SECONDARY_BTN, L"Choose...",
        kMargin + 145, y - 2, 80, kButtonHeight);

    return builder.Build();
}

void InitGlowEffectsPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    CheckDlgButton(page, IDC_GLOW_ENABLE,
        data->workingOptions.enableNeonGlow ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_BITMAP_INTERCEPT,
        data->workingOptions.enableBitmapIntercept ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_FILE_GRADIENT,
        data->workingOptions.enableFileGradientFont ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_EXPLORER_ACCENT,
        data->workingOptions.useExplorerAccentColors ? BST_CHECKED : BST_UNCHECKED);

    CheckDlgButton(page, IDC_GLOW_SURF_LISTVIEW,
        data->workingOptions.glowPalette.listView.enabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_SURF_HEADER,
        data->workingOptions.glowPalette.header.enabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_SURF_REBAR,
        data->workingOptions.glowPalette.rebar.enabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_SURF_TOOLBAR,
        data->workingOptions.glowPalette.toolbar.enabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_SURF_EDIT,
        data->workingOptions.glowPalette.edits.enabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_SURF_DIRECTUI,
        data->workingOptions.glowPalette.directUi.enabled ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_SURF_SCROLLBAR,
        data->workingOptions.glowPalette.scrollbars.enabled ? BST_CHECKED : BST_UNCHECKED);

    CheckDlgButton(page, IDC_GLOW_CUSTOM,
        data->workingOptions.useCustomNeonGlowColors ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(page, IDC_GLOW_USE_GRADIENT,
        data->workingOptions.useNeonGlowGradient ? BST_CHECKED : BST_UNCHECKED);

    UpdateBrush(&data->glowPrimaryBrush, data->workingOptions.neonGlowPrimaryColor);
    UpdateBrush(&data->glowSecondaryBrush, data->workingOptions.neonGlowSecondaryColor);

    // Update control states
    bool glowEnabled = data->workingOptions.enableNeonGlow;
    bool customColors = glowEnabled && data->workingOptions.useCustomNeonGlowColors;
    bool gradientEnabled = customColors && data->workingOptions.useNeonGlowGradient;

    EnableWindow(GetDlgItem(page, IDC_GLOW_CUSTOM), glowEnabled);
    EnableWindow(GetDlgItem(page, IDC_GLOW_USE_GRADIENT), customColors);
    EnableWindow(GetDlgItem(page, IDC_GLOW_PRIMARY_LABEL), customColors);
    EnableWindow(GetDlgItem(page, IDC_GLOW_PRIMARY_PREVIEW), customColors);
    EnableWindow(GetDlgItem(page, IDC_GLOW_PRIMARY_BTN), customColors);
    EnableWindow(GetDlgItem(page, IDC_GLOW_SECONDARY_LABEL), gradientEnabled);
    EnableWindow(GetDlgItem(page, IDC_GLOW_SECONDARY_PREVIEW), gradientEnabled);
    EnableWindow(GetDlgItem(page, IDC_GLOW_SECONDARY_BTN), gradientEnabled);
}

INT_PTR CALLBACK GlowEffectsPageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitGlowEffectsPage(page, data);

        // Initialize scrollbar for the page
        RECT clientRect;
        GetClientRect(page, &clientRect);
        constexpr int kActualContentHeight = 520;  // Match CreateGlowEffectsPageTemplate
        int pageHeight = clientRect.bottom - clientRect.top;

        SCROLLINFO si = {};
        si.cbSize = sizeof(SCROLLINFO);
        si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin = 0;
        si.nMax = kActualContentHeight;
        si.nPage = static_cast<UINT>(pageHeight);
        si.nPos = 0;
        SetScrollInfo(page, SB_VERT, &si, TRUE);

        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (code == BN_CLICKED) {
                bool changed = false;
                bool updateStates = false;

                if (id == IDC_GLOW_ENABLE) {
                    data->workingOptions.enableNeonGlow =
                        (IsDlgButtonChecked(page, IDC_GLOW_ENABLE) == BST_CHECKED);
                    changed = updateStates = true;
                }
                else if (id == IDC_GLOW_BITMAP_INTERCEPT) {
                    data->workingOptions.enableBitmapIntercept =
                        (IsDlgButtonChecked(page, IDC_GLOW_BITMAP_INTERCEPT) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_FILE_GRADIENT) {
                    data->workingOptions.enableFileGradientFont =
                        (IsDlgButtonChecked(page, IDC_GLOW_FILE_GRADIENT) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_EXPLORER_ACCENT) {
                    data->workingOptions.useExplorerAccentColors =
                        (IsDlgButtonChecked(page, IDC_GLOW_EXPLORER_ACCENT) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_LISTVIEW) {
                    data->workingOptions.glowPalette.listView.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_LISTVIEW) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_HEADER) {
                    data->workingOptions.glowPalette.header.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_HEADER) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_REBAR) {
                    data->workingOptions.glowPalette.rebar.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_REBAR) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_TOOLBAR) {
                    data->workingOptions.glowPalette.toolbar.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_TOOLBAR) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_EDIT) {
                    data->workingOptions.glowPalette.edits.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_EDIT) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_DIRECTUI) {
                    data->workingOptions.glowPalette.directUi.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_DIRECTUI) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_SURF_SCROLLBAR) {
                    data->workingOptions.glowPalette.scrollbars.enabled =
                        (IsDlgButtonChecked(page, IDC_GLOW_SURF_SCROLLBAR) == BST_CHECKED);
                    changed = true;
                }
                else if (id == IDC_GLOW_CUSTOM) {
                    data->workingOptions.useCustomNeonGlowColors =
                        (IsDlgButtonChecked(page, IDC_GLOW_CUSTOM) == BST_CHECKED);
                    UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                    changed = updateStates = true;
                }
                else if (id == IDC_GLOW_USE_GRADIENT) {
                    data->workingOptions.useNeonGlowGradient =
                        (IsDlgButtonChecked(page, IDC_GLOW_USE_GRADIENT) == BST_CHECKED);
                    UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                    changed = updateStates = true;
                }
                else if (id == IDC_GLOW_PRIMARY_BTN) {
                    if (ChooseColor(page, &data->workingOptions.neonGlowPrimaryColor)) {
                        UpdateBrush(&data->glowPrimaryBrush, data->workingOptions.neonGlowPrimaryColor);
                        InvalidateRect(GetDlgItem(page, IDC_GLOW_PRIMARY_PREVIEW), nullptr, TRUE);
                        UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                        changed = true;
                    }
                }
                else if (id == IDC_GLOW_SECONDARY_BTN) {
                    if (ChooseColor(page, &data->workingOptions.neonGlowSecondaryColor)) {
                        UpdateBrush(&data->glowSecondaryBrush, data->workingOptions.neonGlowSecondaryColor);
                        InvalidateRect(GetDlgItem(page, IDC_GLOW_SECONDARY_PREVIEW), nullptr, TRUE);
                        UpdateGlowPaletteFromLegacySettings(data->workingOptions);
                        changed = true;
                    }
                }

                if (updateStates) {
                    bool glowEnabled = data->workingOptions.enableNeonGlow;
                    bool customColors = glowEnabled && data->workingOptions.useCustomNeonGlowColors;
                    bool gradientEnabled = customColors && data->workingOptions.useNeonGlowGradient;

                    EnableWindow(GetDlgItem(page, IDC_GLOW_CUSTOM), glowEnabled);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_USE_GRADIENT), customColors);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_PRIMARY_LABEL), customColors);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_PRIMARY_PREVIEW), customColors);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_PRIMARY_BTN), customColors);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_SECONDARY_LABEL), gradientEnabled);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_SECONDARY_PREVIEW), gradientEnabled);
                    EnableWindow(GetDlgItem(page, IDC_GLOW_SECONDARY_BTN), gradientEnabled);
                }

                if (changed) {
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            break;
        }

        case WM_DRAWITEM: {
            if (!data) break;

            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis->CtlType != ODT_STATIC) break;
            if (dis->CtlID == IDC_GLOW_PRIMARY_PREVIEW) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.neonGlowPrimaryColor);
                return TRUE;
            }
            else if (dis->CtlID == IDC_GLOW_SECONDARY_PREVIEW) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.neonGlowSecondaryColor);
                return TRUE;
            }
            break;
        }

        case WM_VSCROLL: {
            // Handle vertical scrolling
            SCROLLINFO si = {};
            si.cbSize = sizeof(SCROLLINFO);
            si.fMask = SIF_ALL;
            GetScrollInfo(page, SB_VERT, &si);

            int oldPos = si.nPos;

            switch (LOWORD(wParam)) {
                case SB_TOP:
                    si.nPos = si.nMin;
                    break;
                case SB_BOTTOM:
                    si.nPos = si.nMax;
                    break;
                case SB_LINEUP:
                    si.nPos -= 20;
                    break;
                case SB_LINEDOWN:
                    si.nPos += 20;
                    break;
                case SB_PAGEUP:
                    si.nPos -= si.nPage;
                    break;
                case SB_PAGEDOWN:
                    si.nPos += si.nPage;
                    break;
                case SB_THUMBTRACK:
                case SB_THUMBPOSITION:
                    si.nPos = si.nTrackPos;
                    break;
            }

            si.fMask = SIF_POS;
            SetScrollInfo(page, SB_VERT, &si, TRUE);
            GetScrollInfo(page, SB_VERT, &si);

            if (si.nPos != oldPos) {
                ScrollWindow(page, 0, oldPos - si.nPos, nullptr, nullptr);
                UpdateWindow(page);
            }

            return TRUE;
        }

        case WM_MOUSEWHEEL: {
            // Handle mouse wheel scrolling
            int delta = GET_WHEEL_DELTA_WPARAM(wParam);
            int scrollAmount = -delta / WHEEL_DELTA;

            SCROLLINFO si = {};
            si.cbSize = sizeof(SCROLLINFO);
            si.fMask = SIF_ALL;
            GetScrollInfo(page, SB_VERT, &si);

            int oldPos = si.nPos;
            si.nPos += scrollAmount * 40;  // 40 pixels per wheel notch

            si.fMask = SIF_POS;
            SetScrollInfo(page, SB_VERT, &si, TRUE);
            GetScrollInfo(page, SB_VERT, &si);

            if (si.nPos != oldPos) {
                ScrollWindow(page, 0, oldPos - si.nPos, nullptr, nullptr);
                UpdateWindow(page);
            }

            return TRUE;
        }

        case WM_SIZE: {
            // Update scrollbar when window is resized
            RECT clientRect;
            GetClientRect(page, &clientRect);
            constexpr int kActualContentHeight = 520;  // Match CreateGlowEffectsPageTemplate
            int pageHeight = clientRect.bottom - clientRect.top;

            SCROLLINFO si = {};
            si.cbSize = sizeof(SCROLLINFO);
            si.fMask = SIF_RANGE | SIF_PAGE;
            si.nMin = 0;
            si.nMax = kActualContentHeight;
            si.nPage = static_cast<UINT>(pageHeight);
            SetScrollInfo(page, SB_VERT, &si, TRUE);
            break;
        }

        case WM_NOTIFY: {
            NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);
            if (nmhdr->code == PSN_APPLY && data) {
                return PSNRET_NOERROR;
            }
            break;
        }
    }

    return FALSE;
}

//=============================================================================
// BACKGROUNDS PAGE
//=============================================================================

// Helper function to generate unique token for async operations
static UINT64 GenerateToken() {
    static std::atomic<UINT64> counter{1};
    return counter.fetch_add(1, std::memory_order_relaxed);
}

// Creates a kPreviewSize x kPreviewSize HBITMAP thumbnail of the image at |path|.
// Returns nullptr if the image can't be loaded.
static HBITMAP CreatePreviewBitmap(const std::wstring& path, int size) {
    if (path.empty()) return nullptr;

    auto gdiBmp = LoadBackgroundBitmap(path);
    if (!gdiBmp || gdiBmp->GetLastStatus() != Gdiplus::Ok) return nullptr;

    HDC screenDC = GetDC(nullptr);
    if (!screenDC) return nullptr;
    HDC memDC = CreateCompatibleDC(screenDC);
    HBITMAP hBmp = CreateCompatibleBitmap(screenDC, size, size);
    ReleaseDC(nullptr, screenDC);
    if (!memDC || !hBmp) {
        if (memDC) DeleteDC(memDC);
        if (hBmp) DeleteObject(hBmp);
        return nullptr;
    }

    HGDIOBJ oldBmp = SelectObject(memDC, hBmp);

    // Fill background with dialog face color
    RECT fillRect = {0, 0, size, size};
    FillRect(memDC, &fillRect, GetSysColorBrush(COLOR_BTNFACE));

    {
        Gdiplus::Graphics g(memDC);
        g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
        g.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
        g.DrawImage(gdiBmp.get(), 0, 0, size, size);
    }

    SelectObject(memDC, oldBmp);
    DeleteDC(memDC);
    return hBmp;
}

// Loads a preview bitmap for |path| and stores it in |stored|, invalidating |ctrl|.
static void SetPreviewBitmap(HWND ctrl, HBITMAP& stored, const std::wstring& path, int size) {
    HBITMAP newBmp = CreatePreviewBitmap(path, size);
    if (stored) DeleteObject(stored);
    stored = newBmp;
    if (ctrl) InvalidateRect(ctrl, nullptr, TRUE);
}

// Draws the preview bitmap |bmp| (or a placeholder) into the DRAWITEMSTRUCT.
static void DrawPreviewControl(const DRAWITEMSTRUCT* dis, HBITMAP bmp) {
    RECT rc = dis->rcItem;
    FillRect(dis->hDC, &rc, GetSysColorBrush(COLOR_BTNFACE));
    DrawEdge(dis->hDC, &rc, EDGE_SUNKEN, BF_RECT);

    if (bmp) {
        BITMAP bmInfo = {};
        GetObject(bmp, sizeof(bmInfo), &bmInfo);
        HDC memDC = CreateCompatibleDC(dis->hDC);
        HGDIOBJ old = SelectObject(memDC, bmp);
        int x = rc.left + 2, y2 = rc.top + 2;
        int w = rc.right - rc.left - 4;
        int h = rc.bottom - rc.top - 4;
        if (w > 0 && h > 0) {
            StretchBlt(dis->hDC, x, y2, w, h, memDC, 0, 0, bmInfo.bmWidth, bmInfo.bmHeight, SRCCOPY);
        }
        SelectObject(memDC, old);
        DeleteDC(memDC);
    } else {
        // Draw a small "no image" indicator
        RECT inner = {rc.left + 2, rc.top + 2, rc.right - 2, rc.bottom - 2};
        SetBkColor(dis->hDC, GetSysColor(COLOR_BTNFACE));
        SetTextColor(dis->hDC, GetSysColor(COLOR_GRAYTEXT));
        DrawTextW(dis->hDC, L"?", 1, &inner, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
}

void UpdateBackgroundControlStates(HWND page, bool enabled) {
    EnableWindow(GetDlgItem(page, IDC_BG_UNIVERSAL_BROWSE), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_UNIVERSAL_CLEAR), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_LIST), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_ADD), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_OPACITY_SLIDER), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_POS_TILE), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_POS_STRETCH), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_POS_CENTER), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_POS_BOTTOMLEFT), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_POS_BOTTOMRIGHT), enabled);
    // Edit/Remove/Clean depend on selection too
    HWND list = GetDlgItem(page, IDC_BG_FOLDER_LIST);
    bool hasSel = list && (SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_EDIT), enabled && hasSel);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_REMOVE), enabled && hasSel);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_CLEAN), enabled);
}

DialogTemplatePtr CreateBackgroundsPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);
    int y = kMargin;  // 16

    // Enable/disable all backgrounds
    builder.AddCheckbox(IDC_BG_ENABLE,
        L"Enable custom folder backgrounds",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing * 2;  // 52

    // ---- Universal Background group ----
    builder.AddGroupBox(IDC_BG_UNIVERSAL_GROUP, L"Universal Background",
        kMargin, y, kPageWidth - 2 * kMargin, 156);
    y += 18;  // 70

    builder.AddStatic(-1, L"Image:",
        kMargin + kGroupMargin, y, 60, kLabelHeight);
    builder.AddStatic(IDC_BG_UNIVERSAL_PREVIEW, L"",
        kMargin + kGroupMargin + 65, y - 2, kPreviewSize, kPreviewSize,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW);
    y += kPreviewSize + kSpacing;  // 144

    builder.AddPushButton(IDC_BG_UNIVERSAL_BROWSE, L"Browse...",
        kMargin + kGroupMargin, y, 80, kButtonHeight);
    builder.AddPushButton(IDC_BG_UNIVERSAL_CLEAR, L"Clear",
        kMargin + kGroupMargin + 90, y, 60, kButtonHeight);
    y += kButtonHeight + 6;  // 174

    builder.AddStatic(IDC_BG_UNIVERSAL_NAME, L"(no image selected)",
        kMargin + kGroupMargin, y,
        kPageWidth - 2 * kMargin - 2 * kGroupMargin, kLabelHeight);
    y += kLabelHeight + kSpacing * 2;  // 208 (group top=52, height=156, so bottom=208)

    // ---- Folder-Specific group ----
    builder.AddGroupBox(IDC_BG_FOLDER_GROUP, L"Folder-Specific Backgrounds",
        kMargin, y, kPageWidth - 2 * kMargin, 156);
    y += 18;  // 226

    builder.AddListBox(IDC_BG_FOLDER_LIST,
        kMargin + kGroupMargin, y, 200, 80);
    builder.AddStatic(IDC_BG_FOLDER_PREVIEW, L"",
        kMargin + kGroupMargin + 210, y, kPreviewSize, kPreviewSize,
        WS_CHILD | WS_VISIBLE | SS_OWNERDRAW);
    y += 86;  // 312

    builder.AddStatic(IDC_BG_FOLDER_NAME, L"",
        kMargin + kGroupMargin, y,
        kPageWidth - 2 * kMargin - 2 * kGroupMargin, kLabelHeight);
    y += kLabelHeight + 6;  // 332

    builder.AddPushButton(IDC_BG_FOLDER_ADD,    L"Add...",
        kMargin + kGroupMargin, y, 60, kButtonHeight);
    builder.AddPushButton(IDC_BG_FOLDER_EDIT,   L"Edit...",
        kMargin + kGroupMargin + 70, y, 60, kButtonHeight);
    builder.AddPushButton(IDC_BG_FOLDER_REMOVE, L"Remove",
        kMargin + kGroupMargin + 140, y, 60, kButtonHeight);
    builder.AddPushButton(IDC_BG_FOLDER_CLEAN,  L"Clean Up...",
        kMargin + kGroupMargin + 210, y, 80, kButtonHeight);
    y += kButtonHeight + kSpacing * 2;

    // ---- Opacity slider ----
    builder.AddStatic(IDC_BG_OPACITY_LABEL, L"Opacity:",
        kMargin, y + 4, 48, kLabelHeight);
    builder.AddSlider(IDC_BG_OPACITY_SLIDER,
        kMargin + 50, y, 200, kSliderHeight);
    builder.AddStatic(IDC_BG_OPACITY_VAL, L"78%",
        kMargin + 256, y + 4, 40, kLabelHeight);
    y += kSliderHeight + kSpacing;

    // ---- Position mode radio buttons ----
    builder.AddGroupBox(IDC_BG_POS_GROUP, L"Image Position",
        kMargin, y, kPageWidth - 2 * kMargin, 5 * kCheckHeight + 30);
    y += 16;

    builder.AddRadioButton(IDC_BG_POS_BOTTOMRIGHT, L"Bottom right (default)",
        kMargin + kGroupMargin, y, 180, kCheckHeight, /*isGroupStart=*/true);
    builder.AddRadioButton(IDC_BG_POS_BOTTOMLEFT, L"Bottom left",
        kMargin + kGroupMargin + 190, y, 100, kCheckHeight);
    y += kCheckHeight + 4;

    builder.AddRadioButton(IDC_BG_POS_CENTER, L"Center",
        kMargin + kGroupMargin, y, 180, kCheckHeight);
    builder.AddRadioButton(IDC_BG_POS_STRETCH, L"Stretch to fill",
        kMargin + kGroupMargin + 190, y, 100, kCheckHeight);
    y += kCheckHeight + 4;

    builder.AddRadioButton(IDC_BG_POS_TILE, L"Tile (repeat)",
        kMargin + kGroupMargin, y, 180, kCheckHeight);

    return builder.Build();
}

void InitBackgroundsPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    CheckDlgButton(page, IDC_BG_ENABLE,
        data->workingOptions.enableFolderBackgrounds ? BST_CHECKED : BST_UNCHECKED);

    // Set universal background name
    const auto& uni = data->workingOptions.universalFolderBackgroundImage;
    HWND nameLabel = GetDlgItem(page, IDC_BG_UNIVERSAL_NAME);
    if (nameLabel) {
        SetWindowTextW(nameLabel, uni.displayName.empty() ? L"(no image selected)" : uni.displayName.c_str());
    }

    // Load universal preview
    SetPreviewBitmap(GetDlgItem(page, IDC_BG_UNIVERSAL_PREVIEW),
                     data->universalPreview,
                     uni.cachedImagePath,
                     kPreviewSize);

    // Populate folder list
    HWND folderList = GetDlgItem(page, IDC_BG_FOLDER_LIST);
    if (folderList) {
        SendMessageW(folderList, LB_RESETCONTENT, 0, 0);
        for (const auto& entry : data->workingOptions.folderBackgroundEntries) {
            std::wstring display = entry.folderPath;
            if (!entry.image.displayName.empty()) {
                display += L"  ->  " + entry.image.displayName;
            }
            SendMessageW(folderList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(display.c_str()));
        }
    }

    // Load folder preview for current selection (if any)
    {
        int sel = folderList ? (int)SendMessageW(folderList, LB_GETCURSEL, 0, 0) : LB_ERR;
        std::wstring folderImgPath;
        if (sel != LB_ERR && sel < (int)data->workingOptions.folderBackgroundEntries.size()) {
            folderImgPath = data->workingOptions.folderBackgroundEntries[sel].image.cachedImagePath;
        }
        SetPreviewBitmap(GetDlgItem(page, IDC_BG_FOLDER_PREVIEW),
                         data->folderPreview,
                         folderImgPath,
                         kPreviewSize);
    }

    // Initialize opacity slider
    HWND opacitySlider = GetDlgItem(page, IDC_BG_OPACITY_SLIDER);
    if (opacitySlider) {
        SendMessageW(opacitySlider, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
        SendMessageW(opacitySlider, TBM_SETPOS, TRUE, data->workingOptions.backgroundOpacity);
    }
    {
        wchar_t buf[32];
        int pct = static_cast<int>(data->workingOptions.backgroundOpacity) * 100 / 255;
        swprintf_s(buf, L"%d%%", pct);
        SetDlgItemTextW(page, IDC_BG_OPACITY_VAL, buf);
    }

    // Initialize position radio buttons
    {
        int radioId = IDC_BG_POS_BOTTOMRIGHT;  // default
        switch (data->workingOptions.backgroundPositionMode) {
            case BackgroundPositionMode::kTile:       radioId = IDC_BG_POS_TILE; break;
            case BackgroundPositionMode::kStretch:    radioId = IDC_BG_POS_STRETCH; break;
            case BackgroundPositionMode::kCenter:     radioId = IDC_BG_POS_CENTER; break;
            case BackgroundPositionMode::kBottomLeft: radioId = IDC_BG_POS_BOTTOMLEFT; break;
            default:                                  radioId = IDC_BG_POS_BOTTOMRIGHT; break;
        }
        CheckRadioButton(page, IDC_BG_POS_TILE, IDC_BG_POS_BOTTOMRIGHT, radioId);
    }

    UpdateBackgroundControlStates(page, data->workingOptions.enableFolderBackgrounds);
}

// Handle list selection change - show folder info and update preview
void OnBgFolderSelChanged(HWND page, OptionsDialogData* data) {
    if (!data) return;
    HWND list = GetDlgItem(page, IDC_BG_FOLDER_LIST);
    if (!list) return;

    int sel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
    bool hasSel = (sel != LB_ERR) &&
                  (sel < static_cast<int>(data->workingOptions.folderBackgroundEntries.size()));

    bool enabled = data->workingOptions.enableFolderBackgrounds;
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_EDIT),   enabled && hasSel);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_REMOVE), enabled && hasSel);

    HWND nameLabel = GetDlgItem(page, IDC_BG_FOLDER_NAME);
    if (nameLabel) {
        if (hasSel) {
            const auto& entry = data->workingOptions.folderBackgroundEntries[sel];
            std::wstring info = entry.folderPath;
            if (!entry.image.displayName.empty()) {
                info += L"  ->  " + entry.image.displayName;
            }
            SetWindowTextW(nameLabel, info.c_str());
        } else {
            SetWindowTextW(nameLabel, L"");
        }
    }

    // Update folder preview
    std::wstring folderImgPath;
    if (hasSel) {
        folderImgPath = data->workingOptions.folderBackgroundEntries[sel].image.cachedImagePath;
    }
    SetPreviewBitmap(GetDlgItem(page, IDC_BG_FOLDER_PREVIEW),
                     data->folderPreview,
                     folderImgPath,
                     kPreviewSize);
}

// Refresh the folder listbox from workingOptions
void RefreshFolderList(HWND page, OptionsDialogData* data) {
    HWND list = GetDlgItem(page, IDC_BG_FOLDER_LIST);
    if (!list || !data) return;
    int prevSel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    for (const auto& entry : data->workingOptions.folderBackgroundEntries) {
        std::wstring display = entry.folderPath;
        if (!entry.image.displayName.empty()) {
            display += L"  ->  " + entry.image.displayName;
        }
        SendMessageW(list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(display.c_str()));
    }
    int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    if (prevSel != LB_ERR && prevSel < count) {
        SendMessageW(list, LB_SETCURSEL, prevSel, 0);
    }
    OnBgFolderSelChanged(page, data);
}

INT_PTR CALLBACK BackgroundsPageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitBackgroundsPage(page, data);
        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
        case WM_COMMAND: {
            if (!data) break;
            int id   = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_BG_ENABLE && code == BN_CLICKED) {
                data->workingOptions.enableFolderBackgrounds =
                    (IsDlgButtonChecked(page, IDC_BG_ENABLE) == BST_CHECKED);
                UpdateBackgroundControlStates(page, data->workingOptions.enableFolderBackgrounds);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_BG_UNIVERSAL_BROWSE && code == BN_CLICKED) {
                std::wstring path;
                if (BrowseForImage(page, &path, &data->lastImageDir)) {
                    CachedImageMetadata meta;
                    std::wstring createdPath;
                    std::wstring errorMsg;
                    if (CopyImageToBackgroundCache(path, path, &meta, &createdPath, &errorMsg)) {
                        if (!createdPath.empty()) {
                            data->createdCachedImages.push_back(createdPath);
                        }
                        data->workingOptions.universalFolderBackgroundImage = meta;
                        // Use the source filename as the display name
                        size_t slashPos = path.find_last_of(L"\\/");
                        std::wstring fileName = (slashPos != std::wstring::npos)
                            ? path.substr(slashPos + 1) : path;
                        data->workingOptions.universalFolderBackgroundImage.displayName = fileName;
                        SetWindowTextW(GetDlgItem(page, IDC_BG_UNIVERSAL_NAME), fileName.c_str());
                    } else {
                        // Fallback: store path directly if caching fails
                        data->workingOptions.universalFolderBackgroundImage.cachedImagePath = path;
                        data->workingOptions.universalFolderBackgroundImage.displayName = path;
                        SetWindowTextW(GetDlgItem(page, IDC_BG_UNIVERSAL_NAME), path.c_str());
                    }
                    // Update preview
                    SetPreviewBitmap(GetDlgItem(page, IDC_BG_UNIVERSAL_PREVIEW),
                                     data->universalPreview,
                                     data->workingOptions.universalFolderBackgroundImage.cachedImagePath,
                                     kPreviewSize);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_BG_UNIVERSAL_CLEAR && code == BN_CLICKED) {
                data->workingOptions.universalFolderBackgroundImage.displayName.clear();
                data->workingOptions.universalFolderBackgroundImage.cachedImagePath.clear();
                SetWindowTextW(GetDlgItem(page, IDC_BG_UNIVERSAL_NAME), L"(no image selected)");
                SetPreviewBitmap(GetDlgItem(page, IDC_BG_UNIVERSAL_PREVIEW),
                                 data->universalPreview, L"", kPreviewSize);
                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_BG_FOLDER_ADD && code == BN_CLICKED) {
                // Browse for folder
                std::wstring folderPath;
                if (BrowseForFolder(page, &folderPath, L"Select folder to assign a background image")) {
                    // Browse for image
                    std::wstring imagePath;
                    if (BrowseForImage(page, &imagePath, &data->lastImageDir)) {
                        CachedImageMetadata meta;
                        std::wstring createdPath;
                        std::wstring errorMsg;
                        if (!CopyImageToBackgroundCache(imagePath, imagePath, &meta, &createdPath, &errorMsg)) {
                            // Use path directly as fallback
                            meta.cachedImagePath = imagePath;
                            meta.displayName     = imagePath;
                        } else {
                            if (!createdPath.empty()) {
                                data->createdCachedImages.push_back(createdPath);
                            }
                            size_t slashPos = imagePath.find_last_of(L"\\/");
                            meta.displayName = (slashPos != std::wstring::npos)
                                ? imagePath.substr(slashPos + 1) : imagePath;
                        }

                        // Check if this folder already exists; update if so
                        bool found = false;
                        for (auto& entry : data->workingOptions.folderBackgroundEntries) {
                            if (_wcsicmp(entry.folderPath.c_str(), folderPath.c_str()) == 0) {
                                entry.image = meta;
                                found = true;
                                break;
                            }
                        }
                        if (!found) {
                            FolderBackgroundEntry entry;
                            entry.folderPath = folderPath;
                            entry.image      = meta;
                            data->workingOptions.folderBackgroundEntries.push_back(std::move(entry));
                        }

                        RefreshFolderList(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            else if (id == IDC_BG_FOLDER_EDIT && code == BN_CLICKED) {
                HWND list = GetDlgItem(page, IDC_BG_FOLDER_LIST);
                int sel = list ? static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0)) : LB_ERR;
                if (sel != LB_ERR && sel < static_cast<int>(data->workingOptions.folderBackgroundEntries.size())) {
                    std::wstring imagePath;
                    if (BrowseForImage(page, &imagePath, &data->lastImageDir)) {
                        CachedImageMetadata meta;
                        std::wstring createdPath;
                        std::wstring errorMsg;
                        if (!CopyImageToBackgroundCache(imagePath, imagePath, &meta, &createdPath, &errorMsg)) {
                            meta.cachedImagePath = imagePath;
                            meta.displayName     = imagePath;
                        } else {
                            if (!createdPath.empty()) {
                                data->createdCachedImages.push_back(createdPath);
                            }
                            size_t slashPos = imagePath.find_last_of(L"\\/");
                            meta.displayName = (slashPos != std::wstring::npos)
                                ? imagePath.substr(slashPos + 1) : imagePath;
                        }
                        data->workingOptions.folderBackgroundEntries[sel].image = meta;
                        RefreshFolderList(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            else if (id == IDC_BG_FOLDER_REMOVE && code == BN_CLICKED) {
                HWND list = GetDlgItem(page, IDC_BG_FOLDER_LIST);
                int sel = list ? static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0)) : LB_ERR;
                if (sel != LB_ERR && sel < static_cast<int>(data->workingOptions.folderBackgroundEntries.size())) {
                    auto& entry = data->workingOptions.folderBackgroundEntries[sel];
                    if (!entry.image.cachedImagePath.empty()) {
                        data->pendingCachedRemovals.push_back(entry.image.cachedImagePath);
                    }
                    data->workingOptions.folderBackgroundEntries.erase(
                        data->workingOptions.folderBackgroundEntries.begin() + sel);
                    RefreshFolderList(page, data);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_BG_FOLDER_CLEAN && code == BN_CLICKED) {
                UpdateCachedImageUsage(data->workingOptions, /*forceMaintenance=*/true);
                MessageBoxW(page, L"Cache cleanup scheduled.", L"Clean Up", MB_OK | MB_ICONINFORMATION);
            }
            else if (id == IDC_BG_FOLDER_LIST && code == LBN_SELCHANGE) {
                OnBgFolderSelChanged(page, data);
            }
            else if (code == BN_CLICKED && id >= IDC_BG_POS_TILE && id <= IDC_BG_POS_BOTTOMRIGHT) {
                switch (id) {
                    case IDC_BG_POS_TILE:        data->workingOptions.backgroundPositionMode = BackgroundPositionMode::kTile; break;
                    case IDC_BG_POS_STRETCH:     data->workingOptions.backgroundPositionMode = BackgroundPositionMode::kStretch; break;
                    case IDC_BG_POS_CENTER:      data->workingOptions.backgroundPositionMode = BackgroundPositionMode::kCenter; break;
                    case IDC_BG_POS_BOTTOMLEFT:  data->workingOptions.backgroundPositionMode = BackgroundPositionMode::kBottomLeft; break;
                    case IDC_BG_POS_BOTTOMRIGHT: data->workingOptions.backgroundPositionMode = BackgroundPositionMode::kBottomRight; break;
                }
                PropSheet_Changed(GetParent(page), page);
            }
            break;
        }

        case WM_HSCROLL: {
            if (!data) break;
            HWND slider = reinterpret_cast<HWND>(lParam);
            int id = GetDlgCtrlID(slider);
            if (id == IDC_BG_OPACITY_SLIDER) {
                int pos = static_cast<int>(SendMessageW(slider, TBM_GETPOS, 0, 0));
                data->workingOptions.backgroundOpacity = static_cast<BYTE>(pos);
                wchar_t buf[32];
                swprintf_s(buf, L"%d%%", pos * 100 / 255);
                SetDlgItemTextW(page, IDC_BG_OPACITY_VAL, buf);
                PropSheet_Changed(GetParent(page), page);
            }
            break;
        }

        case WM_DRAWITEM: {
            if (!data) break;
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (!dis || dis->CtlType != ODT_STATIC) break;
            if (dis->CtlID == IDC_BG_UNIVERSAL_PREVIEW) {
                DrawPreviewControl(dis, data->universalPreview);
                return TRUE;
            }
            if (dis->CtlID == IDC_BG_FOLDER_PREVIEW) {
                DrawPreviewControl(dis, data->folderPreview);
                return TRUE;
            }
            break;
        }

        case WM_NOTIFY: {
            NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);
            if (nmhdr->code == PSN_APPLY && data) {
                // Nothing special needed on apply; main dialog handles save
                return PSNRET_NOERROR;
            }
            break;
        }
    }

    return FALSE;
}

//=============================================================================
// CONTEXT MENUS PAGE - Simplified
//=============================================================================

DialogTemplatePtr CreateContextMenusPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    // Layout constants for split view
    constexpr int treeWidth = 155;
    constexpr int treeLeft = kMargin;
    constexpr int propLeft = treeLeft + treeWidth + 6;
    constexpr int propWidth = kPageWidth - propLeft - kMargin;

    int y = kMargin;

    // Template selector at top, full width
    builder.AddStatic(-1, L"Template:",
        kMargin, y + 2, 55, kLabelHeight);
    builder.AddComboBox(IDC_CTX_TEMPLATE,
        kMargin + 58, y, 140, kComboHeight);
    y += kEditHeight + 6;

    int treeTop = y;

    // Left panel: Tree view
    builder.AddTreeView(IDC_CTX_TREE,
        treeLeft, treeTop, treeWidth, 340);

    // Left panel: Buttons below tree
    int btnY = treeTop + 344;
    constexpr int btnW2 = 74;
    builder.AddPushButton(IDC_CTX_ADD_COMMAND, L"+ Command",
        treeLeft, btnY, btnW2, kButtonHeight);
    builder.AddPushButton(IDC_CTX_ADD_SUBMENU, L"+ Submenu",
        treeLeft + btnW2 + 3, btnY, btnW2, kButtonHeight);
    btnY += kButtonHeight + 3;
    builder.AddPushButton(IDC_CTX_ADD_SEPARATOR, L"+ Separator",
        treeLeft, btnY, btnW2, kButtonHeight);
    builder.AddPushButton(IDC_CTX_REMOVE, L"Remove",
        treeLeft + btnW2 + 3, btnY, btnW2, kButtonHeight);
    btnY += kButtonHeight + 3;
    builder.AddPushButton(IDC_CTX_MOVE_UP, L"Up",
        treeLeft, btnY, 50, kButtonHeight);
    builder.AddPushButton(IDC_CTX_MOVE_DOWN, L"Down",
        treeLeft + 53, btnY, 50, kButtonHeight);
    builder.AddPushButton(IDC_CTX_INDENT, L">",
        treeLeft + 106, btnY, 22, kButtonHeight);
    builder.AddPushButton(IDC_CTX_OUTDENT, L"<",
        treeLeft + 130, btnY, 22, kButtonHeight);

    // Right panel: Properties (scrollable child window created at runtime)
    // We create a static frame as the host; actual controls are created dynamically
    builder.AddStatic(IDC_CTX_PROPS_PANEL, L"",
        propLeft, treeTop, propWidth, 420,
        WS_CHILD | WS_VISIBLE | WS_BORDER | WS_CLIPCHILDREN);

    return builder.Build();
}

// ---------------------------------------------------------------------------
// Context Menu tree path resolution — resolves HTREEITEM to a
// ContextMenuItem* by walking the tree parent chain to build an index path,
// then indexing into the nested workingOptions.contextMenuItems vectors.
// ---------------------------------------------------------------------------

// Build the index path from root to the given tree item
static bool BuildTreeIndexPath(HWND tree, HTREEITEM hItem, std::vector<size_t>& path) {
    path.clear();
    // Walk up to root, collecting indices
    std::vector<size_t> reversePath;
    HTREEITEM cur = hItem;
    while (cur) {
        TVITEMW tv{};
        tv.hItem = cur;
        tv.mask = TVIF_PARAM;
        if (!TreeView_GetItem(tree, &tv)) return false;
        reversePath.push_back(static_cast<size_t>(tv.lParam));
        cur = TreeView_GetParent(tree, cur);
    }
    // Reverse so index[0] is the top-level index
    path.assign(reversePath.rbegin(), reversePath.rend());
    return !path.empty();
}

// Resolve a tree item to a ContextMenuItem pointer using the index path
static ContextMenuItem* ResolveTreeItem(HWND tree, HTREEITEM hItem,
                                        std::vector<ContextMenuItem>& items) {
    std::vector<size_t> path;
    if (!BuildTreeIndexPath(tree, hItem, path)) return nullptr;

    std::vector<ContextMenuItem>* vec = &items;
    ContextMenuItem* result = nullptr;
    for (size_t idx : path) {
        if (idx >= vec->size()) return nullptr;
        result = &(*vec)[idx];
        vec = &result->children;
    }
    return result;
}

// Resolve the parent vector and index for a tree item (for remove/move ops)
static bool ResolveTreeItemParent(HWND tree, HTREEITEM hItem,
                                  std::vector<ContextMenuItem>& items,
                                  std::vector<ContextMenuItem>** parentVec,
                                  size_t* index) {
    std::vector<size_t> path;
    if (!BuildTreeIndexPath(tree, hItem, path)) return false;

    std::vector<ContextMenuItem>* vec = &items;
    for (size_t i = 0; i + 1 < path.size(); ++i) {
        if (path[i] >= vec->size()) return false;
        vec = &(*vec)[path[i]].children;
    }
    if (path.back() >= vec->size()) return false;
    *parentVec = vec;
    *index = path.back();
    return true;
}

// ---------------------------------------------------------------------------
// Context Menu tree population
// ---------------------------------------------------------------------------

void PopulateContextMenuTree(HWND tree, const std::vector<ContextMenuItem>& items,
                             HTREEITEM parent = TVI_ROOT) {
    for (size_t i = 0; i < items.size(); ++i) {
        const auto& item = items[i];
        TVINSERTSTRUCTW insert{};
        insert.hParent = parent;
        insert.hInsertAfter = TVI_LAST;
        insert.item.mask = TVIF_TEXT | TVIF_PARAM;

        std::wstring displayText;
        if (item.type == ContextMenuItemType::kSeparator) {
            displayText = L"\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500";
        } else if (item.type == ContextMenuItemType::kSubmenu) {
            displayText = L"\u25B6 " + (item.label.empty() ? L"(Submenu)" : item.label);
        } else {
            displayText = item.label.empty() ? L"(Unnamed)" : item.label;
        }
        if (!item.enabled) displayText += L" [disabled]";

        insert.item.pszText = const_cast<wchar_t*>(displayText.c_str());
        insert.item.lParam = static_cast<LPARAM>(i);
        HTREEITEM hItem = TreeView_InsertItem(tree, &insert);

        if (item.type == ContextMenuItemType::kSubmenu && !item.children.empty()) {
            PopulateContextMenuTree(tree, item.children, hItem);
            TreeView_Expand(tree, hItem, TVE_EXPAND);
        }
    }
}

void RefreshContextMenuTree(HWND page, OptionsDialogData* data) {
    if (!data) return;
    HWND tree = GetDlgItem(page, IDC_CTX_TREE);
    if (!tree) return;

    TreeView_DeleteAllItems(tree);
    if (data->workingOptions.contextMenuItems.empty()) {
        TVINSERTSTRUCTW insert{};
        insert.hParent = TVI_ROOT;
        insert.hInsertAfter = TVI_LAST;
        insert.item.mask = TVIF_TEXT;
        std::wstring text = L"(No items)";
        insert.item.pszText = text.data();
        TreeView_InsertItem(tree, &insert);
    } else {
        PopulateContextMenuTree(tree, data->workingOptions.contextMenuItems);
    }
}

// ---------------------------------------------------------------------------
// Properties panel — dynamically created child controls hosted in a
// scrollable child window.  We use raw CreateWindowExW for the panel
// controls rather than a dialog template so we can scroll them.
// ---------------------------------------------------------------------------

struct CtxPropsPanel {
    HWND host = nullptr;       // The scrollable child
    HWND parent = nullptr;     // The property sheet page
    int contentHeight = 0;
    int scrollY = 0;
    bool suppressNotify = false; // Prevents feedback loop when loading item

    // Control HWNDs inside the panel
    HWND lblLabel = nullptr, edLabel = nullptr;
    HWND lblIcon = nullptr, edIcon = nullptr, btnIconBrowse = nullptr;
    HWND lblExe = nullptr, edExe = nullptr, btnExeBrowse = nullptr;
    HWND lblArgs = nullptr, edArgs = nullptr;
    HWND lblWorkDir = nullptr, edWorkDir = nullptr, btnWorkDirBrowse = nullptr;
    HWND cbWindowState = nullptr;
    HWND chkAdmin = nullptr, chkWait = nullptr, chkEnabled = nullptr, chkExpandEnv = nullptr;
    HWND lblMinSel = nullptr, edMinSel = nullptr;
    HWND lblMaxSel = nullptr, edMaxSel = nullptr;
    HWND chkFiles = nullptr, chkFolders = nullptr, chkMultiple = nullptr;
    HWND lblPatterns = nullptr, edPatterns = nullptr;
    HWND lblExclude = nullptr, edExclude = nullptr;
    HWND cbAnchor = nullptr;
    HWND lblFolderFilters = nullptr, edFolderFilters = nullptr;
    HWND chkConfirm = nullptr;
    HWND lblConfirmMsg = nullptr, edConfirmMsg = nullptr;
    HWND lblAddlCmds = nullptr, lbAddlCmds = nullptr;
    HWND btnAddlAdd = nullptr, btnAddlRemove = nullptr;
};

static CtxPropsPanel* s_ctxProps = nullptr;

static LRESULT CALLBACK CtxPropsPanelProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* props = reinterpret_cast<CtxPropsPanel*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg) {
    case WM_MOUSEWHEEL: {
        if (!props) break;
        int delta = GET_WHEEL_DELTA_WPARAM(wParam);
        int step = 30;
        props->scrollY -= (delta > 0) ? step : -step;
        // Clamp
        RECT rc;
        GetClientRect(hwnd, &rc);
        int maxScroll = props->contentHeight - (rc.bottom - rc.top);
        if (maxScroll < 0) maxScroll = 0;
        if (props->scrollY < 0) props->scrollY = 0;
        if (props->scrollY > maxScroll) props->scrollY = maxScroll;
        SetScrollPos(hwnd, SB_VERT, props->scrollY, TRUE);
        ScrollWindowEx(hwnd, 0, 0, nullptr, nullptr, nullptr, nullptr,
                       SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN);
        // Reposition all children
        HDWP dwp = BeginDeferWindowPos(40);
        HWND child = GetWindow(hwnd, GW_CHILD);
        while (child) {
            RECT cr;
            GetWindowRect(child, &cr);
            MapWindowPoints(HWND_DESKTOP, hwnd, reinterpret_cast<POINT*>(&cr), 2);
            // We stored original Y in the child's ID user data — we'll just
            // invalidate and rely on WM_PAINT. For simplicity, use
            // ScrollWindow approach instead.
            child = GetWindow(child, GW_HWNDNEXT);
        }
        if (dwp) EndDeferWindowPos(dwp);
        InvalidateRect(hwnd, nullptr, TRUE);
        return 0;
    }
    case WM_VSCROLL: {
        if (!props) break;
        RECT rc;
        GetClientRect(hwnd, &rc);
        int maxScroll = props->contentHeight - (rc.bottom - rc.top);
        if (maxScroll < 0) maxScroll = 0;
        int oldY = props->scrollY;
        switch (LOWORD(wParam)) {
        case SB_LINEUP: props->scrollY -= 20; break;
        case SB_LINEDOWN: props->scrollY += 20; break;
        case SB_PAGEUP: props->scrollY -= (rc.bottom - rc.top); break;
        case SB_PAGEDOWN: props->scrollY += (rc.bottom - rc.top); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: props->scrollY = HIWORD(wParam); break;
        }
        if (props->scrollY < 0) props->scrollY = 0;
        if (props->scrollY > maxScroll) props->scrollY = maxScroll;
        if (props->scrollY != oldY) {
            SetScrollPos(hwnd, SB_VERT, props->scrollY, TRUE);
            ScrollWindowEx(hwnd, 0, -(props->scrollY - oldY), nullptr, nullptr,
                           nullptr, nullptr, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN);
        }
        return 0;
    }
    case WM_COMMAND: {
        // Forward WM_COMMAND to the property sheet page so ContextMenusPageProc handles it
        if (props && props->parent) {
            SendMessageW(props->parent, msg, wParam, lParam);
        }
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// DLU to pixel helpers (approximate for 9pt Segoe UI)
static int DluToPixelX(int dlu, HWND hwnd) {
    RECT r = {0, 0, dlu, 0};
    MapDialogRect(GetParent(hwnd) ? GetParent(hwnd) : hwnd, &r);
    return r.right;
}
static int DluToPixelY(int dlu, HWND hwnd) {
    RECT r = {0, 0, 0, dlu};
    MapDialogRect(GetParent(hwnd) ? GetParent(hwnd) : hwnd, &r);
    return r.bottom;
}

static HWND CreateLabel(HWND parent, const wchar_t* text, int x, int y, int w, int h) {
    return CreateWindowExW(0, L"STATIC", text,
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        x, y, w, h, parent, nullptr,
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static HWND CreateEditCtrl(HWND parent, int id, int x, int y, int w, int h,
                           DWORD extraStyle = 0) {
    return CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL | extraStyle,
        x, y, w, h, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static HWND CreateCheckCtrl(HWND parent, int id, const wchar_t* text, int x, int y, int w, int h) {
    return CreateWindowExW(0, L"BUTTON", text,
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        x, y, w, h, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static HWND CreateBtnCtrl(HWND parent, int id, const wchar_t* text, int x, int y, int w, int h) {
    return CreateWindowExW(0, L"BUTTON", text,
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        x, y, w, h, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static HWND CreateComboCtrl(HWND parent, int id, int x, int y, int w, int dropH) {
    return CreateWindowExW(0, L"COMBOBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST,
        x, y, w, dropH, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static HWND CreateListBoxCtrl(HWND parent, int id, int x, int y, int w, int h) {
    return CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY,
        x, y, w, h, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static HWND CreateGroupCtrl(HWND parent, const wchar_t* text, int x, int y, int w, int h) {
    return CreateWindowExW(0, L"BUTTON", text,
        WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
        x, y, w, h, parent, nullptr,
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE)), nullptr);
}

static void SetCtrlFont(HWND ctrl, HWND page) {
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(page, WM_GETFONT, 0, 0));
    if (font) SendMessageW(ctrl, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
}

static void SetAllChildFonts(HWND parent, HWND page) {
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(page, WM_GETFONT, 0, 0));
    if (!font) return;
    HWND child = GetWindow(parent, GW_CHILD);
    while (child) {
        SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
        child = GetWindow(child, GW_HWNDNEXT);
    }
}

static void CreatePropertiesPanel(HWND page) {
    HWND hostFrame = GetDlgItem(page, IDC_CTX_PROPS_PANEL);
    if (!hostFrame) return;

    RECT hostRect;
    GetClientRect(hostFrame, &hostRect);
    int panelW = hostRect.right - hostRect.left;

    // Register the scrollable panel class if not yet done
    static ATOM sPanelClass = 0;
    if (!sPanelClass) {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = CtxPropsPanelProc;
        wc.hInstance = reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(page, GWLP_HINSTANCE));
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        wc.lpszClassName = L"ShellTabsCtxPropsPanel";
        wc.style = CS_HREDRAW | CS_VREDRAW;
        sPanelClass = RegisterClassExW(&wc);
    }

    // Create scrollable child inside the static frame
    HWND panel = CreateWindowExW(0, L"ShellTabsCtxPropsPanel", L"",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_CLIPCHILDREN,
        0, 0, hostRect.right, hostRect.bottom,
        hostFrame, nullptr,
        reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(page, GWLP_HINSTANCE)), nullptr);

    auto* props = new CtxPropsPanel();
    props->host = panel;
    props->parent = page;
    SetWindowLongPtrW(panel, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(props));
    s_ctxProps = props;

    // Layout constants (pixels)
    const int pad = 6;
    const int lblH = 14;
    const int edH = 18;
    const int chkH = 16;
    const int lblW = 60;
    const int fullW = panelW - 2 * pad;
    const int edtLeft = pad + lblW + 2;
    const int edtW = fullW - lblW - 2;
    const int browseW = 28;
    const int edtWBrowse = edtW - browseW - 2;

    int y = pad;

    // --- Command section ---
    CreateGroupCtrl(panel, L"Command", pad - 2, y, fullW + 4, 7 * (edH + 4) + lblH + 8);
    y += lblH + 2;

    props->lblLabel = CreateLabel(panel, L"Label:", pad + 4, y + 2, lblW, lblH);
    props->edLabel = CreateEditCtrl(panel, IDC_CTX_LABEL, edtLeft, y, edtW, edH);
    y += edH + 4;

    props->lblIcon = CreateLabel(panel, L"Icon:", pad + 4, y + 2, lblW, lblH);
    props->edIcon = CreateEditCtrl(panel, IDC_CTX_ICON, edtLeft, y, edtWBrowse, edH);
    props->btnIconBrowse = CreateBtnCtrl(panel, IDC_CTX_ICON_BROWSE, L"...",
        edtLeft + edtWBrowse + 2, y, browseW, edH);
    y += edH + 4;

    props->lblExe = CreateLabel(panel, L"Executable:", pad + 4, y + 2, lblW, lblH);
    props->edExe = CreateEditCtrl(panel, IDC_CTX_COMMAND, edtLeft, y, edtWBrowse, edH);
    props->btnExeBrowse = CreateBtnCtrl(panel, IDC_CTX_COMMAND_BROWSE, L"...",
        edtLeft + edtWBrowse + 2, y, browseW, edH);
    y += edH + 4;

    props->lblArgs = CreateLabel(panel, L"Arguments:", pad + 4, y + 2, lblW, lblH);
    props->edArgs = CreateEditCtrl(panel, IDC_CTX_ARGS, edtLeft, y, edtW, edH);
    y += edH + 4;

    props->lblWorkDir = CreateLabel(panel, L"Work Dir:", pad + 4, y + 2, lblW, lblH);
    props->edWorkDir = CreateEditCtrl(panel, IDC_CTX_WORKDIR, edtLeft, y, edtWBrowse, edH);
    props->btnWorkDirBrowse = CreateBtnCtrl(panel, IDC_CTX_WORKDIR_BROWSE, L"...",
        edtLeft + edtWBrowse + 2, y, browseW, edH);
    y += edH + 4;

    CreateLabel(panel, L"Window:", pad + 4, y + 2, lblW, lblH);
    props->cbWindowState = CreateComboCtrl(panel, IDC_CTX_WINDOW_STATE,
        edtLeft, y, edtW, 120);
    ComboBox_AddString(props->cbWindowState, L"Normal");
    ComboBox_AddString(props->cbWindowState, L"Minimized");
    ComboBox_AddString(props->cbWindowState, L"Maximized");
    ComboBox_AddString(props->cbWindowState, L"Hidden");
    y += edH + 4;

    CreateLabel(panel, L"Anchor:", pad + 4, y + 2, lblW, lblH);
    props->cbAnchor = CreateComboCtrl(panel, IDC_CTX_ANCHOR, edtLeft, y, edtW, 120);
    ComboBox_AddString(props->cbAnchor, L"Default");
    ComboBox_AddString(props->cbAnchor, L"Top");
    ComboBox_AddString(props->cbAnchor, L"Bottom");
    ComboBox_AddString(props->cbAnchor, L"Before Shell Items");
    ComboBox_AddString(props->cbAnchor, L"After Shell Items");
    y += edH + 8;

    // --- Behavior section ---
    CreateGroupCtrl(panel, L"Behavior", pad - 2, y, fullW + 4, 4 * (chkH + 4) + lblH + 4);
    y += lblH + 2;

    props->chkEnabled = CreateCheckCtrl(panel, IDC_CTX_ENABLED, L"Enabled", pad + 4, y, fullW - 8, chkH);
    y += chkH + 4;
    props->chkAdmin = CreateCheckCtrl(panel, IDC_CTX_RUN_ADMIN, L"Run as administrator", pad + 4, y, fullW - 8, chkH);
    y += chkH + 4;
    props->chkWait = CreateCheckCtrl(panel, IDC_CTX_WAIT, L"Wait for completion", pad + 4, y, fullW - 8, chkH);
    y += chkH + 4;
    props->chkExpandEnv = CreateCheckCtrl(panel, IDC_CTX_EXPAND_ENV, L"Expand environment variables", pad + 4, y, fullW - 8, chkH);
    y += chkH + 8;

    // --- Visibility section ---
    CreateGroupCtrl(panel, L"Visibility", pad - 2, y, fullW + 4, 7 * (edH + 4) + lblH + 4);
    y += lblH + 2;

    props->lblMinSel = CreateLabel(panel, L"Min sel:", pad + 4, y + 2, lblW, lblH);
    props->edMinSel = CreateEditCtrl(panel, IDC_CTX_MIN_SEL, edtLeft, y, 40, edH, ES_NUMBER);
    CreateLabel(panel, L"Max:", edtLeft + 45, y + 2, 30, lblH);
    props->edMaxSel = CreateEditCtrl(panel, IDC_CTX_MAX_SEL, edtLeft + 75, y, 40, edH, ES_NUMBER);
    y += edH + 4;

    props->chkFiles = CreateCheckCtrl(panel, IDC_CTX_FILES, L"Files", pad + 4, y, 50, chkH);
    props->chkFolders = CreateCheckCtrl(panel, IDC_CTX_FOLDERS, L"Folders", pad + 58, y, 55, chkH);
    props->chkMultiple = CreateCheckCtrl(panel, IDC_CTX_MULTIPLE, L"Multiple", pad + 117, y, 60, chkH);
    y += chkH + 4;

    props->lblPatterns = CreateLabel(panel, L"Patterns:", pad + 4, y + 2, lblW, lblH);
    props->edPatterns = CreateEditCtrl(panel, IDC_CTX_PATTERNS, edtLeft, y, edtW, edH);
    y += edH + 4;

    props->lblExclude = CreateLabel(panel, L"Exclude:", pad + 4, y + 2, lblW, lblH);
    props->edExclude = CreateEditCtrl(panel, IDC_CTX_EXCLUDE, edtLeft, y, edtW, edH);
    y += edH + 8;

    // --- Advanced section ---
    CreateGroupCtrl(panel, L"Advanced", pad - 2, y, fullW + 4, 5 * (edH + 4) + 60 + lblH + 8);
    y += lblH + 2;

    props->lblFolderFilters = CreateLabel(panel, L"Folder filters:", pad + 4, y + 2, lblW, lblH);
    props->edFolderFilters = CreateEditCtrl(panel, IDC_CTX_FOLDER_FILTERS, edtLeft, y, edtW, edH);
    y += edH + 4;

    props->chkConfirm = CreateCheckCtrl(panel, IDC_CTX_CONFIRM, L"Confirm before execute", pad + 4, y, fullW - 8, chkH);
    y += chkH + 4;

    props->lblConfirmMsg = CreateLabel(panel, L"Message:", pad + 4, y + 2, lblW, lblH);
    props->edConfirmMsg = CreateEditCtrl(panel, IDC_CTX_CONFIRM_MSG, edtLeft, y, edtW, edH);
    y += edH + 4;

    props->lblAddlCmds = CreateLabel(panel, L"Additional commands:", pad + 4, y + 2, fullW - 8, lblH);
    y += lblH + 2;
    props->lbAddlCmds = CreateListBoxCtrl(panel, IDC_CTX_ADDL_CMDS, pad + 4, y, fullW - 8 - browseW - 4, 50);
    props->btnAddlAdd = CreateBtnCtrl(panel, IDC_CTX_ADDL_ADD, L"+",
        fullW - browseW, y, browseW, 20);
    props->btnAddlRemove = CreateBtnCtrl(panel, IDC_CTX_ADDL_REMOVE, L"-",
        fullW - browseW, y + 24, browseW, 20);
    y += 50 + 8;

    props->contentHeight = y + pad;

    // Set fonts for all children
    SetAllChildFonts(panel, page);

    // Set up scrollbar
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = props->contentHeight;
    si.nPage = static_cast<UINT>(hostRect.bottom - hostRect.top);
    si.nPos = 0;
    SetScrollInfo(panel, SB_VERT, &si, TRUE);

    // Initially disable all property controls until an item is selected
    EnableWindow(panel, FALSE);
}

// ---------------------------------------------------------------------------
// Load/Save selected item properties to/from the panel controls
// ---------------------------------------------------------------------------

static std::wstring JoinStrings(const std::vector<std::wstring>& v, const wchar_t* sep) {
    std::wstring result;
    for (size_t i = 0; i < v.size(); ++i) {
        if (i > 0) result += sep;
        result += v[i];
    }
    return result;
}

static std::vector<std::wstring> SplitString(const std::wstring& s, wchar_t sep) {
    std::vector<std::wstring> result;
    size_t start = 0;
    for (size_t i = 0; i <= s.size(); ++i) {
        if (i == s.size() || s[i] == sep) {
            std::wstring token = s.substr(start, i - start);
            // Trim whitespace
            size_t first = token.find_first_not_of(L" \t");
            if (first != std::wstring::npos) {
                size_t last = token.find_last_not_of(L" \t");
                token = token.substr(first, last - first + 1);
            } else {
                token.clear();
            }
            if (!token.empty()) result.push_back(token);
            start = i + 1;
        }
    }
    return result;
}

static std::wstring GetEditText(HWND edit) {
    int len = GetWindowTextLengthW(edit);
    if (len <= 0) return {};
    std::wstring buf(static_cast<size_t>(len) + 1, L'\0');
    GetWindowTextW(edit, buf.data(), len + 1);
    buf.resize(static_cast<size_t>(len));
    return buf;
}

static void LoadItemToPanel(ContextMenuItem* item) {
    if (!s_ctxProps) return;
    auto* p = s_ctxProps;
    p->suppressNotify = true;

    // Enable the panel
    EnableWindow(p->host, TRUE);

    bool isCommand = item && item->type == ContextMenuItemType::kCommand;
    bool isSeparator = item && item->type == ContextMenuItemType::kSeparator;

    // Command section — hide for separators
    auto showCmd = isCommand ? SW_SHOW : SW_HIDE;
    auto enableCmd = [&](HWND h) { if (h) { ShowWindow(h, isSeparator ? SW_HIDE : SW_SHOW); EnableWindow(h, isCommand); } };
    enableCmd(p->edExe); enableCmd(p->btnExeBrowse); enableCmd(p->lblExe);
    enableCmd(p->edArgs); enableCmd(p->lblArgs);
    enableCmd(p->edWorkDir); enableCmd(p->btnWorkDirBrowse); enableCmd(p->lblWorkDir);
    if (p->cbWindowState) { ShowWindow(p->cbWindowState, isSeparator ? SW_HIDE : showCmd); EnableWindow(p->cbWindowState, isCommand); }

    // Label/Icon always visible (except separator)
    if (p->edLabel) { ShowWindow(p->edLabel, isSeparator ? SW_HIDE : SW_SHOW); }
    if (p->lblLabel) ShowWindow(p->lblLabel, isSeparator ? SW_HIDE : SW_SHOW);
    if (p->edIcon) { ShowWindow(p->edIcon, isSeparator ? SW_HIDE : SW_SHOW); }
    if (p->lblIcon) ShowWindow(p->lblIcon, isSeparator ? SW_HIDE : SW_SHOW);
    if (p->btnIconBrowse) ShowWindow(p->btnIconBrowse, isSeparator ? SW_HIDE : SW_SHOW);

    if (!item) {
        // Clear all fields
        SetWindowTextW(p->edLabel, L"");
        SetWindowTextW(p->edIcon, L"");
        SetWindowTextW(p->edExe, L"");
        SetWindowTextW(p->edArgs, L"");
        SetWindowTextW(p->edWorkDir, L"");
        ComboBox_SetCurSel(p->cbWindowState, 0);
        ComboBox_SetCurSel(p->cbAnchor, 0);
        Button_SetCheck(p->chkEnabled, BST_UNCHECKED);
        Button_SetCheck(p->chkAdmin, BST_UNCHECKED);
        Button_SetCheck(p->chkWait, BST_UNCHECKED);
        Button_SetCheck(p->chkExpandEnv, BST_CHECKED);
        SetWindowTextW(p->edMinSel, L"0");
        SetWindowTextW(p->edMaxSel, L"0");
        Button_SetCheck(p->chkFiles, BST_CHECKED);
        Button_SetCheck(p->chkFolders, BST_CHECKED);
        Button_SetCheck(p->chkMultiple, BST_CHECKED);
        SetWindowTextW(p->edPatterns, L"");
        SetWindowTextW(p->edExclude, L"");
        SetWindowTextW(p->edFolderFilters, L"");
        Button_SetCheck(p->chkConfirm, BST_UNCHECKED);
        SetWindowTextW(p->edConfirmMsg, L"");
        SendMessageW(p->lbAddlCmds, LB_RESETCONTENT, 0, 0);
        EnableWindow(p->host, FALSE);
        p->suppressNotify = false;
        return;
    }

    // Populate fields
    SetWindowTextW(p->edLabel, item->label.c_str());
    SetWindowTextW(p->edIcon, item->iconSource.c_str());
    SetWindowTextW(p->edExe, item->executable.c_str());
    SetWindowTextW(p->edArgs, item->arguments.c_str());
    SetWindowTextW(p->edWorkDir, item->workingDirectory.c_str());
    ComboBox_SetCurSel(p->cbWindowState, static_cast<int>(item->windowState));
    ComboBox_SetCurSel(p->cbAnchor, static_cast<int>(item->anchor));

    Button_SetCheck(p->chkEnabled, item->enabled ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(p->chkAdmin, item->runAsAdmin ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(p->chkWait, item->waitForCompletion ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(p->chkExpandEnv, item->expandEnvironmentVars ? BST_CHECKED : BST_UNCHECKED);

    SetWindowTextW(p->edMinSel, std::to_wstring(item->visibility.minimumSelection).c_str());
    SetWindowTextW(p->edMaxSel, std::to_wstring(item->visibility.maximumSelection).c_str());
    Button_SetCheck(p->chkFiles, item->visibility.showForFiles ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(p->chkFolders, item->visibility.showForFolders ? BST_CHECKED : BST_UNCHECKED);
    Button_SetCheck(p->chkMultiple, item->visibility.showForMultiple ? BST_CHECKED : BST_UNCHECKED);
    SetWindowTextW(p->edPatterns, JoinStrings(item->visibility.filePatterns, L"; ").c_str());
    SetWindowTextW(p->edExclude, JoinStrings(item->visibility.excludePatterns, L"; ").c_str());

    SetWindowTextW(p->edFolderFilters, JoinStrings(item->folderPathFilters, L"; ").c_str());
    Button_SetCheck(p->chkConfirm, item->confirmBeforeExecute ? BST_CHECKED : BST_UNCHECKED);
    SetWindowTextW(p->edConfirmMsg, item->confirmMessage.c_str());

    SendMessageW(p->lbAddlCmds, LB_RESETCONTENT, 0, 0);
    for (const auto& cmd : item->additionalCommands) {
        SendMessageW(p->lbAddlCmds, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(cmd.c_str()));
    }

    p->suppressNotify = false;
}

static void SavePanelToItem(ContextMenuItem* item) {
    if (!s_ctxProps || !item) return;
    auto* p = s_ctxProps;

    item->label = GetEditText(p->edLabel);
    item->iconSource = GetEditText(p->edIcon);
    item->executable = GetEditText(p->edExe);
    item->arguments = GetEditText(p->edArgs);
    item->workingDirectory = GetEditText(p->edWorkDir);

    int wndSel = ComboBox_GetCurSel(p->cbWindowState);
    if (wndSel >= 0) item->windowState = static_cast<ContextMenuWindowState>(wndSel);
    int ancSel = ComboBox_GetCurSel(p->cbAnchor);
    if (ancSel >= 0) item->anchor = static_cast<ContextMenuInsertionAnchor>(ancSel);

    item->enabled = Button_GetCheck(p->chkEnabled) == BST_CHECKED;
    item->runAsAdmin = Button_GetCheck(p->chkAdmin) == BST_CHECKED;
    item->waitForCompletion = Button_GetCheck(p->chkWait) == BST_CHECKED;
    item->expandEnvironmentVars = Button_GetCheck(p->chkExpandEnv) == BST_CHECKED;

    item->visibility.minimumSelection = _wtoi(GetEditText(p->edMinSel).c_str());
    item->visibility.maximumSelection = _wtoi(GetEditText(p->edMaxSel).c_str());
    item->visibility.showForFiles = Button_GetCheck(p->chkFiles) == BST_CHECKED;
    item->visibility.showForFolders = Button_GetCheck(p->chkFolders) == BST_CHECKED;
    item->visibility.showForMultiple = Button_GetCheck(p->chkMultiple) == BST_CHECKED;
    item->visibility.filePatterns = SplitString(GetEditText(p->edPatterns), L';');
    item->visibility.excludePatterns = SplitString(GetEditText(p->edExclude), L';');

    item->folderPathFilters = SplitString(GetEditText(p->edFolderFilters), L';');
    item->confirmBeforeExecute = Button_GetCheck(p->chkConfirm) == BST_CHECKED;
    item->confirmMessage = GetEditText(p->edConfirmMsg);

    // Read additional commands from listbox
    item->additionalCommands.clear();
    int count = static_cast<int>(SendMessageW(p->lbAddlCmds, LB_GETCOUNT, 0, 0));
    for (int i = 0; i < count; ++i) {
        int len = static_cast<int>(SendMessageW(p->lbAddlCmds, LB_GETTEXTLEN, i, 0));
        if (len > 0) {
            std::wstring buf(static_cast<size_t>(len) + 1, L'\0');
            SendMessageW(p->lbAddlCmds, LB_GETTEXT, i, reinterpret_cast<LPARAM>(buf.data()));
            buf.resize(static_cast<size_t>(len));
            item->additionalCommands.push_back(buf);
        }
    }
}

// ---------------------------------------------------------------------------
// Template menu item creators
// ---------------------------------------------------------------------------

static ContextMenuItem CreateCommandPromptMenuItem() {
    ContextMenuItem item{};
    item.type = ContextMenuItemType::kCommand;
    item.label = L"Command Prompt Here";
    item.executable = L"cmd.exe";
    item.arguments = L"/k cd /d \"%V\"";
    item.workingDirectory = L"%V";
    item.windowState = ContextMenuWindowState::kNormal;
    item.runAsAdmin = false;
    item.enabled = true;
    item.anchor = ContextMenuInsertionAnchor::kDefault;
    item.visibility.showForFiles = true;
    item.visibility.showForFolders = true;
    item.visibility.minimumSelection = 0;
    item.visibility.maximumSelection = 100;
    return item;
}

static ContextMenuItem CreatePowerShellMenuItem() {
    ContextMenuItem item{};
    item.type = ContextMenuItemType::kCommand;
    item.label = L"PowerShell Here";
    item.executable = L"powershell.exe";
    item.arguments = L"-NoExit -Command \"Set-Location -Path '%V'\"";
    item.workingDirectory = L"%V";
    item.windowState = ContextMenuWindowState::kNormal;
    item.runAsAdmin = false;
    item.enabled = true;
    item.anchor = ContextMenuInsertionAnchor::kDefault;
    item.visibility.showForFiles = true;
    item.visibility.showForFolders = true;
    item.visibility.minimumSelection = 0;
    item.visibility.maximumSelection = 100;
    return item;
}

static ContextMenuItem CreateVSCodeMenuItem() {
    ContextMenuItem item{};
    item.type = ContextMenuItemType::kCommand;
    item.label = L"Open with VS Code";
    item.executable = L"code";
    item.arguments = L"\"%V\"";
    item.workingDirectory = L"%V";
    item.windowState = ContextMenuWindowState::kNormal;
    item.runAsAdmin = false;
    item.enabled = true;
    item.anchor = ContextMenuInsertionAnchor::kDefault;
    item.visibility.showForFiles = true;
    item.visibility.showForFolders = true;
    item.visibility.minimumSelection = 0;
    item.visibility.maximumSelection = 100;
    return item;
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

static void InitContextMenusPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    HWND combo = GetDlgItem(page, IDC_CTX_TEMPLATE);
    if (combo) {
        ComboBox_AddString(combo, L"(Select template)");
        ComboBox_AddString(combo, L"Command Prompt Here");
        ComboBox_AddString(combo, L"PowerShell Here");
        ComboBox_AddString(combo, L"Open with VS Code");
        ComboBox_SetCurSel(combo, 0);
    }

    CreatePropertiesPanel(page);
    RefreshContextMenuTree(page, data);
}

// ---------------------------------------------------------------------------
// ContextMenusPageProc
// ---------------------------------------------------------------------------

INT_PTR CALLBACK ContextMenusPageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitContextMenusPage(page, data);
        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
    case WM_COMMAND: {
        if (!data) break;

        const UINT id = LOWORD(wParam);
        const UINT code = HIWORD(wParam);

        // --- Property changes from the right panel ---
        // EN_CHANGE from edit controls
        if (code == EN_CHANGE && s_ctxProps && !s_ctxProps->suppressNotify) {
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                auto* item = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
                if (item) {
                    SavePanelToItem(item);
                    // Update tree text if label changed
                    if (id == IDC_CTX_LABEL) {
                        std::wstring displayText = item->label.empty() ? L"(Unnamed)" : item->label;
                        if (item->type == ContextMenuItemType::kSubmenu)
                            displayText = L"\u25B6 " + displayText;
                        if (!item->enabled) displayText += L" [disabled]";
                        TVITEMW tv{};
                        tv.hItem = sel;
                        tv.mask = TVIF_TEXT;
                        tv.pszText = const_cast<wchar_t*>(displayText.c_str());
                        TreeView_SetItem(tree, &tv);
                    }
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            return TRUE;
        }

        // BN_CLICKED from checkboxes and buttons
        if (code == BN_CLICKED) {
            // Check if it's a property checkbox
            if (id == IDC_CTX_ENABLED || id == IDC_CTX_RUN_ADMIN || id == IDC_CTX_WAIT ||
                id == IDC_CTX_EXPAND_ENV || id == IDC_CTX_FILES || id == IDC_CTX_FOLDERS ||
                id == IDC_CTX_MULTIPLE || id == IDC_CTX_CONFIRM) {
                if (s_ctxProps && !s_ctxProps->suppressNotify) {
                    HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                    HTREEITEM sel = TreeView_GetSelection(tree);
                    if (sel) {
                        auto* item = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
                        if (item) {
                            SavePanelToItem(item);
                            if (id == IDC_CTX_ENABLED) {
                                // Refresh tree text for enabled/disabled indicator
                                RefreshContextMenuTree(page, data);
                            }
                            PropSheet_Changed(GetParent(page), page);
                        }
                    }
                }
                return TRUE;
            }

            // Browse buttons
            if (id == IDC_CTX_ICON_BROWSE || id == IDC_CTX_COMMAND_BROWSE || id == IDC_CTX_WORKDIR_BROWSE) {
                if (id == IDC_CTX_WORKDIR_BROWSE) {
                    // Folder browse
                    BROWSEINFOW bi{};
                    bi.hwndOwner = page;
                    bi.lpszTitle = L"Select Working Directory";
                    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
                    PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
                    if (pidl) {
                        wchar_t path[MAX_PATH];
                        if (SHGetPathFromIDListW(pidl, path)) {
                            SetWindowTextW(s_ctxProps->edWorkDir, path);
                        }
                        CoTaskMemFree(pidl);
                    }
                } else {
                    // File browse
                    OPENFILENAMEW ofn{};
                    wchar_t filePath[MAX_PATH] = {};
                    ofn.lStructSize = sizeof(ofn);
                    ofn.hwndOwner = page;
                    ofn.lpstrFilter = (id == IDC_CTX_ICON_BROWSE)
                        ? L"Icon Files\0*.ico;*.exe;*.dll\0All Files\0*.*\0"
                        : L"Executables\0*.exe;*.bat;*.cmd;*.ps1\0All Files\0*.*\0";
                    ofn.lpstrFile = filePath;
                    ofn.nMaxFile = MAX_PATH;
                    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
                    if (GetOpenFileNameW(&ofn)) {
                        HWND target = (id == IDC_CTX_ICON_BROWSE) ? s_ctxProps->edIcon : s_ctxProps->edExe;
                        SetWindowTextW(target, filePath);
                    }
                }
                return TRUE;
            }

            // Additional commands add/remove
            if (id == IDC_CTX_ADDL_ADD && s_ctxProps) {
                HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                HTREEITEM sel = TreeView_GetSelection(tree);
                // Simple input: we'll reuse a small edit dialog approach
                // For simplicity, just add a placeholder that the user edits in the listbox
                SendMessageW(s_ctxProps->lbAddlCmds, LB_ADDSTRING, 0,
                    reinterpret_cast<LPARAM>(L"command.exe /args"));
                if (sel) {
                    auto* item = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
                    if (item) {
                        SavePanelToItem(item);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
                return TRUE;
            }
            if (id == IDC_CTX_ADDL_REMOVE && s_ctxProps) {
                int curSel = static_cast<int>(SendMessageW(s_ctxProps->lbAddlCmds, LB_GETCURSEL, 0, 0));
                if (curSel != LB_ERR) {
                    SendMessageW(s_ctxProps->lbAddlCmds, LB_DELETESTRING, curSel, 0);
                    HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                    HTREEITEM sel = TreeView_GetSelection(tree);
                    if (sel) {
                        auto* item = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
                        if (item) {
                            SavePanelToItem(item);
                            PropSheet_Changed(GetParent(page), page);
                        }
                    }
                }
                return TRUE;
            }
        }

        // CBN_SELCHANGE from combo boxes
        if (code == CBN_SELCHANGE) {
            if (id == IDC_CTX_WINDOW_STATE || id == IDC_CTX_ANCHOR) {
                if (s_ctxProps && !s_ctxProps->suppressNotify) {
                    HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                    HTREEITEM sel = TreeView_GetSelection(tree);
                    if (sel) {
                        auto* item = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
                        if (item) {
                            SavePanelToItem(item);
                            PropSheet_Changed(GetParent(page), page);
                        }
                    }
                }
                return TRUE;
            }

            // Template selection
            if (id == IDC_CTX_TEMPLATE) {
                HWND combo = GetDlgItem(page, IDC_CTX_TEMPLATE);
                int sel = ComboBox_GetCurSel(combo);
                ContextMenuItem newItem;
                bool addItem = false;
                switch (sel) {
                case 1: newItem = CreateCommandPromptMenuItem(); addItem = true; break;
                case 2: newItem = CreatePowerShellMenuItem(); addItem = true; break;
                case 3: newItem = CreateVSCodeMenuItem(); addItem = true; break;
                }
                if (addItem) {
                    data->workingOptions.contextMenuItems.push_back(newItem);
                    RefreshContextMenuTree(page, data);
                    ComboBox_SetCurSel(combo, 0);
                    PropSheet_Changed(GetParent(page), page);
                }
                return TRUE;
            }
        }

        // --- Tree management buttons ---
        if (id == IDC_CTX_ADD_COMMAND) {
            ContextMenuItem item{};
            item.type = ContextMenuItemType::kCommand;
            item.label = L"New Command";
            item.enabled = true;
            item.anchor = ContextMenuInsertionAnchor::kDefault;
            item.windowState = ContextMenuWindowState::kNormal;
            item.visibility.showForFiles = true;
            item.visibility.showForFolders = true;
            item.visibility.minimumSelection = 0;
            item.visibility.maximumSelection = 100;

            // Add to selected submenu or top level
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                auto* parent = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
                if (parent && parent->type == ContextMenuItemType::kSubmenu) {
                    parent->children.push_back(item);
                    RefreshContextMenuTree(page, data);
                    PropSheet_Changed(GetParent(page), page);
                    return TRUE;
                }
            }
            data->workingOptions.contextMenuItems.push_back(item);
            RefreshContextMenuTree(page, data);
            PropSheet_Changed(GetParent(page), page);
            return TRUE;
        }

        if (id == IDC_CTX_ADD_SUBMENU) {
            ContextMenuItem item{};
            item.type = ContextMenuItemType::kSubmenu;
            item.label = L"New Submenu";
            item.enabled = true;
            item.anchor = ContextMenuInsertionAnchor::kDefault;

            data->workingOptions.contextMenuItems.push_back(item);
            RefreshContextMenuTree(page, data);
            PropSheet_Changed(GetParent(page), page);
            return TRUE;
        }

        if (id == IDC_CTX_ADD_SEPARATOR) {
            ContextMenuItem item{};
            item.type = ContextMenuItemType::kSeparator;
            item.enabled = true;
            item.anchor = ContextMenuInsertionAnchor::kDefault;

            data->workingOptions.contextMenuItems.push_back(item);
            RefreshContextMenuTree(page, data);
            PropSheet_Changed(GetParent(page), page);
            return TRUE;
        }

        if (id == IDC_CTX_REMOVE) {
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                std::vector<ContextMenuItem>* parentVec = nullptr;
                size_t idx = 0;
                if (ResolveTreeItemParent(tree, sel, data->workingOptions.contextMenuItems,
                                          &parentVec, &idx)) {
                    parentVec->erase(parentVec->begin() + static_cast<ptrdiff_t>(idx));
                    RefreshContextMenuTree(page, data);
                    LoadItemToPanel(nullptr);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            return TRUE;
        }

        if (id == IDC_CTX_MOVE_UP) {
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                std::vector<ContextMenuItem>* parentVec = nullptr;
                size_t idx = 0;
                if (ResolveTreeItemParent(tree, sel, data->workingOptions.contextMenuItems,
                                          &parentVec, &idx) && idx > 0) {
                    std::swap((*parentVec)[idx], (*parentVec)[idx - 1]);
                    RefreshContextMenuTree(page, data);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            return TRUE;
        }

        if (id == IDC_CTX_MOVE_DOWN) {
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                std::vector<ContextMenuItem>* parentVec = nullptr;
                size_t idx = 0;
                if (ResolveTreeItemParent(tree, sel, data->workingOptions.contextMenuItems,
                                          &parentVec, &idx) && idx + 1 < parentVec->size()) {
                    std::swap((*parentVec)[idx], (*parentVec)[idx + 1]);
                    RefreshContextMenuTree(page, data);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            return TRUE;
        }

        // Indent (make child of previous sibling)
        if (id == IDC_CTX_INDENT) {
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                std::vector<ContextMenuItem>* parentVec = nullptr;
                size_t idx = 0;
                if (ResolveTreeItemParent(tree, sel, data->workingOptions.contextMenuItems,
                                          &parentVec, &idx) && idx > 0) {
                    auto& prevSibling = (*parentVec)[idx - 1];
                    if (prevSibling.type == ContextMenuItemType::kSubmenu) {
                        prevSibling.children.push_back(std::move((*parentVec)[idx]));
                        parentVec->erase(parentVec->begin() + static_cast<ptrdiff_t>(idx));
                        RefreshContextMenuTree(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            return TRUE;
        }

        // Outdent (move to parent's level)
        if (id == IDC_CTX_OUTDENT) {
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = TreeView_GetSelection(tree);
            if (sel) {
                HTREEITEM treeParent = TreeView_GetParent(tree, sel);
                if (treeParent) {
                    // Find the parent item and the grandparent vector
                    std::vector<ContextMenuItem>* childVec = nullptr;
                    size_t childIdx = 0;
                    std::vector<ContextMenuItem>* grandparentVec = nullptr;
                    size_t parentIdx = 0;
                    if (ResolveTreeItemParent(tree, sel, data->workingOptions.contextMenuItems,
                                              &childVec, &childIdx) &&
                        ResolveTreeItemParent(tree, treeParent, data->workingOptions.contextMenuItems,
                                              &grandparentVec, &parentIdx)) {
                        ContextMenuItem moved = std::move((*childVec)[childIdx]);
                        childVec->erase(childVec->begin() + static_cast<ptrdiff_t>(childIdx));
                        grandparentVec->insert(
                            grandparentVec->begin() + static_cast<ptrdiff_t>(parentIdx + 1),
                            std::move(moved));
                        RefreshContextMenuTree(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            return TRUE;
        }

        break;
    }

    case WM_NOTIFY: {
        NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);

        if (nmhdr->idFrom == IDC_CTX_TREE && nmhdr->code == TVN_SELCHANGEDW) {
            if (!data) break;
            NMTREEVIEWW* nmtv = reinterpret_cast<NMTREEVIEWW*>(lParam);
            HWND tree = GetDlgItem(page, IDC_CTX_TREE);
            HTREEITEM sel = nmtv->itemNew.hItem;

            ContextMenuItem* item = nullptr;
            if (sel) {
                item = ResolveTreeItem(tree, sel, data->workingOptions.contextMenuItems);
            }
            LoadItemToPanel(item);
            return TRUE;
        }

        if (nmhdr->code == PSN_APPLY && data) {
            return PSNRET_NOERROR;
        }
        break;
    }

    case WM_DESTROY: {
        if (s_ctxProps) {
            delete s_ctxProps;
            s_ctxProps = nullptr;
        }
        break;
    }
    }

    return FALSE;
}

//=============================================================================
// GROUPS PAGE
//=============================================================================

DialogTemplatePtr CreateGroupsPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    int y = kMargin;

    builder.AddStatic(-1, L"Saved Groups / Islands:",
        kMargin, y, 120, kLabelHeight);
    y += kLabelHeight + 4;

    // List box for groups
    builder.AddListBox(IDC_GRP_LIST,
        kMargin, y, 300, 250);

    // Buttons
    builder.AddPushButton(IDC_GRP_NEW, L"New Group...",
        kMargin + 310, y, 100, kButtonHeight);
    y += kButtonHeight + 6;

    builder.AddPushButton(IDC_GRP_EDIT, L"Edit Group...",
        kMargin + 310, y, 100, kButtonHeight);
    y += kButtonHeight + 6;

    builder.AddPushButton(IDC_GRP_REMOVE, L"Remove",
        kMargin + 310, y, 100, kButtonHeight);

    return builder.Build();
}

// Group Editor Dialog Data
struct GroupEditorData {
    SavedGroup* group = nullptr;
    bool isNew = false;
    OptionsDialogData* optionsData = nullptr;
};

INT_PTR CALLBACK GroupEditorDialogProc(HWND dialog, UINT msg, WPARAM wParam, LPARAM lParam) {
    GroupEditorData* data = reinterpret_cast<GroupEditorData*>(GetWindowLongPtrW(dialog, GWLP_USERDATA));

    switch (msg) {
        case WM_INITDIALOG: {
            data = reinterpret_cast<GroupEditorData*>(lParam);
            SetWindowLongPtrW(dialog, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));

            // Set dialog title
            SetWindowTextW(dialog, data->isNew ? L"New Group" : L"Edit Group");

            // Resize dialog
            SetWindowPos(dialog, nullptr, 0, 0, 500, 450, SWP_NOMOVE | SWP_NOZORDER);

            // Create controls using DialogBuilder approach but manually
            HINSTANCE hInst = GetModuleHandleInstance();
            HFONT hFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));

            int y = 10;
            int x = 10;
            int labelWidth = 480;
            int editWidth = 480;

            // Name label and edit
            HWND label = CreateWindowW(L"STATIC", L"Group Name:",
                WS_CHILD | WS_VISIBLE, x, y, labelWidth, 16,
                dialog, nullptr, hInst, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 20;

            HWND nameEdit = CreateWindowW(L"EDIT", data->group->name.c_str(),
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
                x, y, editWidth, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_NAME), hInst, nullptr);
            SendMessageW(nameEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 34;

            // Paths label and listbox
            label = CreateWindowW(L"STATIC", L"Folders / Locations:",
                WS_CHILD | WS_VISIBLE, x, y, labelWidth, 16,
                dialog, nullptr, hInst, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 20;

            HWND pathList = CreateWindowW(L"LISTBOX", nullptr,
                WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
                x, y, editWidth, 150,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_PATHS), hInst, nullptr);
            SendMessageW(pathList, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            // Populate paths
            for (const auto& path : data->group->tabPaths) {
                SendMessageW(pathList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(path.c_str()));
            }
            y += 160;

            // Path buttons
            int btnX = x;
            HWND btnAdd = CreateWindowW(L"BUTTON", L"Add...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                btnX, y, 100, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_ADD), hInst, nullptr);
            SendMessageW(btnAdd, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            btnX += 105;

            HWND btnEdit = CreateWindowW(L"BUTTON", L"Edit...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                btnX, y, 100, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_EDIT_PATH), hInst, nullptr);
            SendMessageW(btnEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            EnableWindow(btnEdit, FALSE);
            btnX += 105;

            HWND btnRemove = CreateWindowW(L"BUTTON", L"Remove",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                btnX, y, 100, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_REMOVE_PATH), hInst, nullptr);
            SendMessageW(btnRemove, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            EnableWindow(btnRemove, FALSE);
            y += 34;

            // Color
            label = CreateWindowW(L"STATIC", L"Color:",
                WS_CHILD | WS_VISIBLE, x, y, labelWidth, 16,
                dialog, nullptr, hInst, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 20;

            (void)CreateWindowW(L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | SS_OWNERDRAW | WS_BORDER,
                x, y, 50, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_COLOR_PREVIEW), hInst, nullptr);

            HWND btnColor = CreateWindowW(L"BUTTON", L"Choose Color...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                x + 60, y, 120, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_COLOR_BTN), hInst, nullptr);
            SendMessageW(btnColor, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 34;

            // Outline style
            label = CreateWindowW(L"STATIC", L"Outline Style:",
                WS_CHILD | WS_VISIBLE, x, y, labelWidth, 16,
                dialog, nullptr, hInst, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 20;

            HWND styleCombo = CreateWindowW(L"COMBOBOX", nullptr,
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
                x, y, 150, 200,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_STYLE), hInst, nullptr);
            SendMessageW(styleCombo, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            ComboBox_AddString(styleCombo, L"Solid");
            ComboBox_AddString(styleCombo, L"Dashed");
            ComboBox_AddString(styleCombo, L"Dotted");
            ComboBox_SetCurSel(styleCombo, static_cast<int>(data->group->outlineStyle));
            y += 34;

            // OK/Cancel buttons
            int btnY = y + 10;
            HWND btnOK = CreateWindowW(L"BUTTON", L"OK",
                WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                300, btnY, 80, 24,
                dialog, reinterpret_cast<HMENU>(IDOK), hInst, nullptr);
            SendMessageW(btnOK, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            HWND btnCancel = CreateWindowW(L"BUTTON", L"Cancel",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                390, btnY, 80, 24,
                dialog, reinterpret_cast<HMENU>(IDCANCEL), hInst, nullptr);
            SendMessageW(btnCancel, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            return TRUE;
        }

        case WM_DRAWITEM: {
            DRAWITEMSTRUCT* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis->CtlID == IDC_GRP_ED_COLOR_PREVIEW && data && data->group) {
                HBRUSH brush = CreateSolidBrush(data->group->color);
                FillRect(dis->hDC, &dis->rcItem, brush);
                DeleteObject(brush);
                return TRUE;
            }
            break;
        }

        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_GRP_ED_PATHS && code == LBN_SELCHANGE) {
                HWND pathList = GetDlgItem(dialog, IDC_GRP_ED_PATHS);
                bool hasSelection = (SendMessageW(pathList, LB_GETCURSEL, 0, 0) != LB_ERR);
                EnableWindow(GetDlgItem(dialog, IDC_GRP_ED_EDIT_PATH), hasSelection);
                EnableWindow(GetDlgItem(dialog, IDC_GRP_ED_REMOVE_PATH), hasSelection);
            }
            else if (id == IDC_GRP_ED_ADD && code == BN_CLICKED) {
                std::wstring path;
                if (BrowseForFolderImpl(dialog, &path, L"Select folder or shell location to add:", true)) {
                    HWND pathList = GetDlgItem(dialog, IDC_GRP_ED_PATHS);
                    data->group->tabPaths.push_back(path);
                    SendMessageW(pathList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(path.c_str()));
                }
            }
            else if (id == IDC_GRP_ED_EDIT_PATH && code == BN_CLICKED) {
                HWND pathList = GetDlgItem(dialog, IDC_GRP_ED_PATHS);
                int sel = static_cast<int>(SendMessageW(pathList, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->group->tabPaths.size())) {
                    std::wstring path = data->group->tabPaths[sel];
                    if (BrowseForFolderImpl(dialog, &path, L"Select folder or shell location:", true)) {
                        data->group->tabPaths[sel] = path;
                        SendMessageW(pathList, LB_DELETESTRING, sel, 0);
                        SendMessageW(pathList, LB_INSERTSTRING, sel, reinterpret_cast<LPARAM>(path.c_str()));
                        SendMessageW(pathList, LB_SETCURSEL, sel, 0);
                    }
                }
            }
            else if (id == IDC_GRP_ED_REMOVE_PATH && code == BN_CLICKED) {
                HWND pathList = GetDlgItem(dialog, IDC_GRP_ED_PATHS);
                int sel = static_cast<int>(SendMessageW(pathList, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->group->tabPaths.size())) {
                    data->group->tabPaths.erase(data->group->tabPaths.begin() + sel);
                    SendMessageW(pathList, LB_DELETESTRING, sel, 0);
                }
            }
            else if (id == IDC_GRP_ED_COLOR_BTN && code == BN_CLICKED) {
                CHOOSECOLORW cc{};
                static COLORREF customColors[16] = {};
                cc.lStructSize = sizeof(cc);
                cc.hwndOwner = dialog;
                cc.rgbResult = data->group->color;
                cc.lpCustColors = customColors;
                cc.Flags = CC_FULLOPEN | CC_RGBINIT;

                if (ChooseColorW(&cc)) {
                    data->group->color = cc.rgbResult;
                    InvalidateRect(GetDlgItem(dialog, IDC_GRP_ED_COLOR_PREVIEW), nullptr, TRUE);
                }
            }
            else if (id == IDOK && code == BN_CLICKED) {
                // Get name
                wchar_t name[256];
                GetDlgItemTextW(dialog, IDC_GRP_ED_NAME, name, 256);
                std::wstring newName = name;

                if (newName.empty()) {
                    MessageBoxW(dialog, L"Group name cannot be empty.", L"Validation Error", MB_OK | MB_ICONWARNING);
                    return TRUE;
                }

                // Check for duplicate names (only if renaming or new)
                if (data->optionsData && (data->isNew || newName != data->group->name)) {
                    for (const auto& g : data->optionsData->workingGroups) {
                        if (g.name == newName && &g != data->group) {
                            MessageBoxW(dialog, L"A group with this name already exists.", L"Validation Error", MB_OK | MB_ICONWARNING);
                            return TRUE;
                        }
                    }
                }

                data->group->name = newName;

                // Get outline style
                HWND styleCombo = GetDlgItem(dialog, IDC_GRP_ED_STYLE);
                int styleIdx = ComboBox_GetCurSel(styleCombo);
                if (styleIdx >= 0) {
                    data->group->outlineStyle = static_cast<TabGroupOutlineStyle>(styleIdx);
                }

                EndDialog(dialog, IDOK);
                return TRUE;
            }
            else if (id == IDCANCEL && code == BN_CLICKED) {
                EndDialog(dialog, IDCANCEL);
                return TRUE;
            }
            break;
        }
    }

    return FALSE;
}

bool ShowGroupEditorDialog(HWND parent, SavedGroup& group, bool isNew, OptionsDialogData* optionsData) {
    GroupEditorData editorData;
    editorData.group = &group;
    editorData.isNew = isNew;
    editorData.optionsData = optionsData;

    // Create a runtime dialog template
    struct {
        DLGTEMPLATE dlg;
        WORD menu;
        WORD windowClass;
        WORD title;
    } template_data = {};

    template_data.dlg.style = DS_SETFONT | DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU;
    template_data.dlg.dwExtendedStyle = 0;
    template_data.dlg.cdit = 0;  // We'll create controls manually in WM_INITDIALOG
    template_data.dlg.x = 0;
    template_data.dlg.y = 0;
    template_data.dlg.cx = 300;
    template_data.dlg.cy = 200;

    INT_PTR result = DialogBoxIndirectParamW(GetModuleHandleInstance(),
                                              reinterpret_cast<LPCDLGTEMPLATEW>(&template_data),
                                              parent, GroupEditorDialogProc,
                                              reinterpret_cast<LPARAM>(&editorData));

    return result == IDOK;
}

void InitGroupsPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    HWND list = GetDlgItem(page, IDC_GRP_LIST);
    if (list) {
        SendMessageW(list, LB_RESETCONTENT, 0, 0);
        for (const auto& group : data->workingGroups) {
            SendMessageW(list, LB_ADDSTRING, 0,
                reinterpret_cast<LPARAM>(group.name.c_str()));
        }
    }

    // Update button states
    bool hasSelection = (SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR);
    EnableWindow(GetDlgItem(page, IDC_GRP_EDIT), hasSelection);
    EnableWindow(GetDlgItem(page, IDC_GRP_REMOVE), hasSelection);

    if (!data->focusHandled && !data->focusGroupId.empty() && list) {
        for (size_t i = 0; i < data->workingGroups.size(); ++i) {
            const std::wstring& currentId =
                (i < data->workingGroupIds.size() && !data->workingGroupIds[i].empty())
                    ? data->workingGroupIds[i]
                    : data->workingGroups[i].name;
            if (_wcsicmp(currentId.c_str(), data->focusGroupId.c_str()) != 0 &&
                _wcsicmp(data->workingGroups[i].name.c_str(), data->focusGroupId.c_str()) != 0) {
                continue;
            }

            SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(i), 0);
            EnableWindow(GetDlgItem(page, IDC_GRP_EDIT), TRUE);
            EnableWindow(GetDlgItem(page, IDC_GRP_REMOVE), TRUE);
            data->focusHandled = true;
            if (data->focusGroupEdit) {
                PostMessageW(page, WM_COMMAND, MAKEWPARAM(IDC_GRP_EDIT, BN_CLICKED), 0);
            }
            break;
        }
    }
}

INT_PTR CALLBACK GroupsPageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitGroupsPage(page, data);
        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_GRP_LIST && code == LBN_SELCHANGE) {
                HWND list = GetDlgItem(page, IDC_GRP_LIST);
                bool hasSelection = (SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR);
                EnableWindow(GetDlgItem(page, IDC_GRP_EDIT), hasSelection);
                EnableWindow(GetDlgItem(page, IDC_GRP_REMOVE), hasSelection);
            }
            else if (id == IDC_GRP_NEW && code == BN_CLICKED) {
                SavedGroup newGroup;
                newGroup.name = L"New Group";
                newGroup.color = RGB(0, 120, 215);
                newGroup.outlineStyle = TabGroupOutlineStyle::kSolid;

                if (ShowGroupEditorDialog(page, newGroup, true, data)) {
                    data->workingGroups.push_back(newGroup);
                    data->workingGroupIds.push_back(std::wstring());
                    data->groupsChanged = true;
                    InitGroupsPage(page, data);
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_GRP_EDIT && code == BN_CLICKED) {
                HWND list = GetDlgItem(page, IDC_GRP_LIST);
                int sel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->workingGroups.size())) {
                    SavedGroup groupCopy = data->workingGroups[sel];
                    if (ShowGroupEditorDialog(page, groupCopy, false, data)) {
                        data->workingGroups[sel] = groupCopy;
                        data->groupsChanged = true;
                        InitGroupsPage(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            else if (id == IDC_GRP_REMOVE && code == BN_CLICKED) {
                HWND list = GetDlgItem(page, IDC_GRP_LIST);
                int sel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->workingGroups.size())) {
                    std::wstring name = data->workingGroups[sel].name;
                    std::wstring confirmMsg = L"Remove group '" + name + L"'?";
                    if (MessageBoxW(page, confirmMsg.c_str(), L"Confirm", MB_YESNO | MB_ICONQUESTION) == IDYES) {
                        if (sel < static_cast<int>(data->workingGroupIds.size())) {
                            if (!data->workingGroupIds[sel].empty()) {
                                data->removedGroupIds.push_back(data->workingGroupIds[sel]);
                            }
                            data->workingGroupIds.erase(data->workingGroupIds.begin() + sel);
                        }
                        data->workingGroups.erase(data->workingGroups.begin() + sel);
                        data->groupsChanged = true;
                        InitGroupsPage(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            break;
        }

        case WM_NOTIFY: {
            NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);
            if (nmhdr->code == PSN_APPLY && data) {
                return PSNRET_NOERROR;
            }
            break;
        }
    }

    return FALSE;
}

//=============================================================================
// WEB FOLDERS PAGE
//=============================================================================

DialogTemplatePtr CreateWebFoldersPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    int y = kMargin;

    builder.AddStatic(-1, L"Configured web directory sites:",
        kMargin, y, 200, kLabelHeight);
    y += kLabelHeight + 4;

    // List box for web folders
    builder.AddListBox(IDC_WEB_LIST,
        kMargin, y, 300, 200);

    // Buttons
    int btnX = kMargin + 310;
    builder.AddPushButton(IDC_WEB_ADD, L"Add...",
        btnX, y, 100, kButtonHeight);

    builder.AddPushButton(IDC_WEB_EDIT, L"Edit...",
        btnX, y + kButtonHeight + 6, 100, kButtonHeight);

    builder.AddPushButton(IDC_WEB_REMOVE, L"Remove",
        btnX, y + (kButtonHeight + 6) * 2, 100, kButtonHeight);

    y += 200 + kSpacing;

    // URL display label
    builder.AddStatic(IDC_WEB_URL_LABEL, L"",
        kMargin, y, 400, kLabelHeight);

    return builder.Build();
}

void InitWebFoldersPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    HWND list = GetDlgItem(page, IDC_WEB_LIST);
    if (list) {
        SendMessageW(list, LB_RESETCONTENT, 0, 0);
        for (const auto& entry : data->workingOptions.webFolderEntries) {
            std::wstring display = entry.displayName;
            if (!entry.enabled) {
                display += L" (disabled)";
            }
            SendMessageW(list, LB_ADDSTRING, 0,
                reinterpret_cast<LPARAM>(display.c_str()));
        }
    }

    // Update button states
    bool hasSelection = list && (SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR);
    EnableWindow(GetDlgItem(page, IDC_WEB_EDIT), hasSelection);
    EnableWindow(GetDlgItem(page, IDC_WEB_REMOVE), hasSelection);

    // Clear URL display
    SetWindowTextW(GetDlgItem(page, IDC_WEB_URL_LABEL), L"");
}

void UpdateWebFolderUrlDisplay(HWND page, OptionsDialogData* data) {
    if (!data) return;

    HWND list = GetDlgItem(page, IDC_WEB_LIST);
    int sel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
    if (sel >= 0 && sel < static_cast<int>(data->workingOptions.webFolderEntries.size())) {
        std::wstring label = L"URL: " + data->workingOptions.webFolderEntries[sel].url;
        SetWindowTextW(GetDlgItem(page, IDC_WEB_URL_LABEL), label.c_str());
    } else {
        SetWindowTextW(GetDlgItem(page, IDC_WEB_URL_LABEL), L"");
    }
}

// Web Folder Editor Dialog Data
struct WebFolderEditorData {
    WebFolderEntry* entry = nullptr;
    bool isNew = false;
};

INT_PTR CALLBACK WebFolderEditorDialogProc(HWND dialog, UINT msg, WPARAM wParam, LPARAM lParam) {
    WebFolderEditorData* data = reinterpret_cast<WebFolderEditorData*>(
        GetWindowLongPtrW(dialog, GWLP_USERDATA));

    switch (msg) {
        case WM_INITDIALOG: {
            data = reinterpret_cast<WebFolderEditorData*>(lParam);
            SetWindowLongPtrW(dialog, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));

            SetWindowTextW(dialog, data->isNew ? L"Add Web Folder" : L"Edit Web Folder");

            // Resize dialog
            SetWindowPos(dialog, nullptr, 0, 0, 400, 310, SWP_NOMOVE | SWP_NOZORDER);

            HFONT hFont = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));

            int y = 14;
            int x = 14;
            int editWidth = 350;

            // Display name
            HWND label = CreateWindowW(L"STATIC", L"Display name:",
                WS_CHILD | WS_VISIBLE, x, y, 120, 16, dialog, nullptr, nullptr, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 18;

            HWND nameEdit = CreateWindowW(L"EDIT", data->entry->displayName.c_str(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL,
                x, y, editWidth, 22, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_WEB_ED_NAME)),
                nullptr, nullptr);
            SendMessageW(nameEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 30;

            // URL
            label = CreateWindowW(L"STATIC", L"URL:",
                WS_CHILD | WS_VISIBLE, x, y, 120, 16, dialog, nullptr, nullptr, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 18;

            HWND urlEdit = CreateWindowW(L"EDIT", data->entry->url.c_str(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL,
                x, y, editWidth, 22, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_WEB_ED_URL)),
                nullptr, nullptr);
            SendMessageW(urlEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 30;

            // Parallel downloads checkbox
            HWND parallelCheck = CreateWindowW(L"BUTTON", L"Enable parallel downloads",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                x, y, 200, 20, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_WEB_ED_PARALLEL)),
                nullptr, nullptr);
            SendMessageW(parallelCheck, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            SendMessageW(parallelCheck, BM_SETCHECK,
                data->entry->parallelDownloads ? BST_CHECKED : BST_UNCHECKED, 0);
            y += 26;

            // Max concurrent downloads
            label = CreateWindowW(L"STATIC", L"Max concurrent:",
                WS_CHILD | WS_VISIBLE, x, y + 2, 110, 16, dialog, nullptr, nullptr, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            HWND maxEdit = CreateWindowW(L"EDIT",
                std::to_wstring(data->entry->maxParallelDownloads).c_str(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_NUMBER,
                x + 114, y, 50, 22, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_WEB_ED_MAX_CONCURRENT)),
                nullptr, nullptr);
            SendMessageW(maxEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            // Speed limit
            label = CreateWindowW(L"STATIC", L"Speed limit KB/s (0=unlimited):",
                WS_CHILD | WS_VISIBLE, x + 180, y + 2, 200, 16, dialog, nullptr, nullptr, nullptr);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 26;

            HWND speedEdit = CreateWindowW(L"EDIT",
                std::to_wstring(data->entry->downloadSpeedLimitKBps).c_str(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_NUMBER,
                x, y, 80, 22, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_WEB_ED_SPEED_LIMIT)),
                nullptr, nullptr);
            SendMessageW(speedEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            y += 34;

            // Enable/disable download settings based on parallel checkbox
            BOOL parallelEnabled = data->entry->parallelDownloads ? TRUE : FALSE;
            EnableWindow(maxEdit, parallelEnabled);
            EnableWindow(speedEdit, parallelEnabled);

            // OK / Cancel buttons
            int btnWidth = 80;
            int btnSpacing = 10;
            int totalBtnWidth = btnWidth * 2 + btnSpacing;
            int btnX = (400 - totalBtnWidth) / 2 - 8;  // account for window border

            HWND okBtn = CreateWindowW(L"BUTTON", L"OK",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                btnX, y, btnWidth, 26, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDOK)),
                nullptr, nullptr);
            SendMessageW(okBtn, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            HWND cancelBtn = CreateWindowW(L"BUTTON", L"Cancel",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                btnX + btnWidth + btnSpacing, y, btnWidth, 26, dialog,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDCANCEL)),
                nullptr, nullptr);
            SendMessageW(cancelBtn, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

            // Focus the name field
            SetFocus(nameEdit);
            return FALSE;  // We set focus manually
        }

        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            if (id == IDOK) {
                std::wstring name = GetControlText(GetDlgItem(dialog, IDC_WEB_ED_NAME));
                std::wstring url = GetControlText(GetDlgItem(dialog, IDC_WEB_ED_URL));

                // Validate
                if (name.empty()) {
                    MessageBoxW(dialog, L"Display name cannot be empty.",
                        L"Validation Error", MB_OK | MB_ICONWARNING);
                    SetFocus(GetDlgItem(dialog, IDC_WEB_ED_NAME));
                    return TRUE;
                }

                if (url.empty()) {
                    MessageBoxW(dialog, L"URL cannot be empty.",
                        L"Validation Error", MB_OK | MB_ICONWARNING);
                    SetFocus(GetDlgItem(dialog, IDC_WEB_ED_URL));
                    return TRUE;
                }

                // URL must start with http:// or https://
                bool validScheme = false;
                if (url.size() >= 7) {
                    std::wstring lower = url.substr(0, 8);
                    for (auto& ch : lower) ch = towlower(ch);
                    if (lower.substr(0, 7) == L"http://" || lower == L"https://") {
                        validScheme = true;
                    }
                }
                if (!validScheme) {
                    MessageBoxW(dialog,
                        L"URL must start with http:// or https://.",
                        L"Validation Error", MB_OK | MB_ICONWARNING);
                    SetFocus(GetDlgItem(dialog, IDC_WEB_ED_URL));
                    return TRUE;
                }

                data->entry->displayName = name;
                data->entry->url = url;
                data->entry->parallelDownloads =
                    (SendMessageW(GetDlgItem(dialog, IDC_WEB_ED_PARALLEL), BM_GETCHECK, 0, 0) == BST_CHECKED);
                data->entry->maxParallelDownloads = std::clamp(
                    GetDlgItemInt(dialog, IDC_WEB_ED_MAX_CONCURRENT, nullptr, FALSE), 1u, 16u);
                data->entry->downloadSpeedLimitKBps = std::max(
                    0, static_cast<int>(GetDlgItemInt(dialog, IDC_WEB_ED_SPEED_LIMIT, nullptr, FALSE)));
                EndDialog(dialog, IDOK);
                return TRUE;
            }
            else if (id == IDC_WEB_ED_PARALLEL && HIWORD(wParam) == BN_CLICKED) {
                BOOL checked = (SendMessageW(GetDlgItem(dialog, IDC_WEB_ED_PARALLEL),
                    BM_GETCHECK, 0, 0) == BST_CHECKED);
                EnableWindow(GetDlgItem(dialog, IDC_WEB_ED_MAX_CONCURRENT), checked);
                EnableWindow(GetDlgItem(dialog, IDC_WEB_ED_SPEED_LIMIT), checked);
                return TRUE;
            }
            else if (id == IDCANCEL) {
                EndDialog(dialog, IDCANCEL);
                return TRUE;
            }
            break;
        }

        case WM_CLOSE:
            EndDialog(dialog, IDCANCEL);
            return TRUE;
    }

    return FALSE;
}

bool ShowWebFolderEditorDialog(HWND parent, WebFolderEntry& entry, bool isNew) {
    WebFolderEditorData editorData;
    editorData.entry = &entry;
    editorData.isNew = isNew;

    struct {
        DLGTEMPLATE dlg;
        WORD menu;
        WORD windowClass;
        WORD title;
    } template_data = {};

    template_data.dlg.style = DS_SETFONT | DS_MODALFRAME | DS_CENTER |
                               WS_POPUP | WS_CAPTION | WS_SYSMENU;
    template_data.dlg.dwExtendedStyle = 0;
    template_data.dlg.cdit = 0;
    template_data.dlg.x = 0;
    template_data.dlg.y = 0;
    template_data.dlg.cx = 250;
    template_data.dlg.cy = 130;

    INT_PTR result = DialogBoxIndirectParamW(GetModuleHandleInstance(),
                                              reinterpret_cast<LPCDLGTEMPLATEW>(&template_data),
                                              parent, WebFolderEditorDialogProc,
                                              reinterpret_cast<LPARAM>(&editorData));
    return result == IDOK;
}

INT_PTR CALLBACK WebFoldersPageProc(HWND page, UINT msg, WPARAM wParam, LPARAM lParam) {
    OptionsDialogData* data = nullptr;

    if (msg == WM_INITDIALOG) {
        PROPSHEETPAGEW* psp = reinterpret_cast<PROPSHEETPAGEW*>(lParam);
        data = reinterpret_cast<OptionsDialogData*>(psp->lParam);
        SetWindowLongPtrW(page, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(data));
        InitWebFoldersPage(page, data);
        return TRUE;
    }

    data = reinterpret_cast<OptionsDialogData*>(GetWindowLongPtrW(page, GWLP_USERDATA));

    switch (msg) {
        case WM_COMMAND: {
            if (!data) break;

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_WEB_LIST && code == LBN_SELCHANGE) {
                HWND list = GetDlgItem(page, IDC_WEB_LIST);
                bool hasSelection = (SendMessageW(list, LB_GETCURSEL, 0, 0) != LB_ERR);
                EnableWindow(GetDlgItem(page, IDC_WEB_EDIT), hasSelection);
                EnableWindow(GetDlgItem(page, IDC_WEB_REMOVE), hasSelection);
                UpdateWebFolderUrlDisplay(page, data);
            }
            else if (id == IDC_WEB_ADD && code == BN_CLICKED) {
                WebFolderEntry newEntry;
                newEntry.displayName = L"New Site";
                newEntry.url = L"https://";
                newEntry.enabled = true;

                if (ShowWebFolderEditorDialog(page, newEntry, true)) {
                    data->workingOptions.webFolderEntries.push_back(newEntry);
                    InitWebFoldersPage(page, data);

                    // Select the newly added item
                    HWND list = GetDlgItem(page, IDC_WEB_LIST);
                    int lastIndex = static_cast<int>(data->workingOptions.webFolderEntries.size()) - 1;
                    SendMessageW(list, LB_SETCURSEL, lastIndex, 0);
                    UpdateWebFolderUrlDisplay(page, data);
                    EnableWindow(GetDlgItem(page, IDC_WEB_EDIT), TRUE);
                    EnableWindow(GetDlgItem(page, IDC_WEB_REMOVE), TRUE);

                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_WEB_EDIT && code == BN_CLICKED) {
                HWND list = GetDlgItem(page, IDC_WEB_LIST);
                int sel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->workingOptions.webFolderEntries.size())) {
                    WebFolderEntry entryCopy = data->workingOptions.webFolderEntries[sel];
                    if (ShowWebFolderEditorDialog(page, entryCopy, false)) {
                        data->workingOptions.webFolderEntries[sel] = entryCopy;
                        InitWebFoldersPage(page, data);
                        SendMessageW(list, LB_SETCURSEL, sel, 0);
                        UpdateWebFolderUrlDisplay(page, data);
                        EnableWindow(GetDlgItem(page, IDC_WEB_EDIT), TRUE);
                        EnableWindow(GetDlgItem(page, IDC_WEB_REMOVE), TRUE);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            else if (id == IDC_WEB_REMOVE && code == BN_CLICKED) {
                HWND list = GetDlgItem(page, IDC_WEB_LIST);
                int sel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->workingOptions.webFolderEntries.size())) {
                    std::wstring name = data->workingOptions.webFolderEntries[sel].displayName;
                    std::wstring confirmMsg = L"Remove web folder '" + name + L"'?";
                    if (MessageBoxW(page, confirmMsg.c_str(), L"Confirm",
                                    MB_YESNO | MB_ICONQUESTION) == IDYES) {
                        data->workingOptions.webFolderEntries.erase(
                            data->workingOptions.webFolderEntries.begin() + sel);
                        InitWebFoldersPage(page, data);
                        PropSheet_Changed(GetParent(page), page);
                    }
                }
            }
            break;
        }

        case WM_NOTIFY: {
            NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);
            if (nmhdr->code == PSN_APPLY && data) {
                return PSNRET_NOERROR;
            }
            break;
        }
    }

    return FALSE;
}

//=============================================================================
// MAIN DIALOG ENTRY POINT
//=============================================================================

}  // namespace

OptionsDialogResult ShowOptionsDialog(HWND parent, OptionsDialogPage initialPage,
                                      const wchar_t* focusSavedGroupId,
                                      bool editFocusedGroup) {
    OptionsDialogResult result{};

    // Load options
    OptionsStore& store = OptionsStore::Instance();
    if (!store.Load()) {
        MessageBoxW(parent, L"Failed to load options.", L"Error", MB_OK | MB_ICONERROR);
        return result;
    }

    // Initialize data
    auto data = std::make_unique<OptionsDialogData>();
    data->originalOptions = store.Get();
    data->workingOptions = data->originalOptions;
    data->initialTab = static_cast<int>(initialPage);

    // Load groups
    GroupStore& groupStore = GroupStore::Instance();
    if (groupStore.Load()) {
        data->originalGroups = groupStore.Groups();
        data->workingGroups = data->originalGroups;
        for (const auto& group : data->originalGroups) {
            data->workingGroupIds.push_back(group.name);
        }
    }

    if (focusSavedGroupId) {
        data->focusGroupId = focusSavedGroupId;
        data->focusGroupEdit = editFocusedGroup;
    }

    // Build dialog templates
    auto generalTemplate = CreateGeneralPageTemplate();
    auto appearanceTemplate = CreateAppearancePageTemplate();
    auto glowTemplate = CreateGlowEffectsPageTemplate();
    auto backgroundsTemplate = CreateBackgroundsPageTemplate();
    auto contextTemplate = CreateContextMenusPageTemplate();
    auto groupsTemplate = CreateGroupsPageTemplate();
    auto webFoldersTemplate = CreateWebFoldersPageTemplate();

    if (!generalTemplate || !appearanceTemplate || !glowTemplate ||
        !backgroundsTemplate || !contextTemplate || !groupsTemplate ||
        !webFoldersTemplate) {
        MessageBoxW(parent, L"Failed to create dialog templates.", L"Error", MB_OK | MB_ICONERROR);
        return result;
    }

    // Create property sheet pages
    std::array<PROPSHEETPAGEW, 7> pages{};
    HINSTANCE hInst = GetModuleHandleW(nullptr);

    pages[0].dwSize = sizeof(PROPSHEETPAGEW);
    pages[0].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[0].hInstance = hInst;
    pages[0].pResource = generalTemplate.get();
    pages[0].pszTitle = L"General";
    pages[0].pfnDlgProc = GeneralPageProc;
    pages[0].lParam = reinterpret_cast<LPARAM>(data.get());

    pages[1].dwSize = sizeof(PROPSHEETPAGEW);
    pages[1].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[1].hInstance = hInst;
    pages[1].pResource = appearanceTemplate.get();
    pages[1].pszTitle = L"Appearance";
    pages[1].pfnDlgProc = AppearancePageProc;
    pages[1].lParam = reinterpret_cast<LPARAM>(data.get());

    pages[2].dwSize = sizeof(PROPSHEETPAGEW);
    pages[2].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[2].hInstance = hInst;
    pages[2].pResource = glowTemplate.get();
    pages[2].pszTitle = L"Glow Effects";
    pages[2].pfnDlgProc = GlowEffectsPageProc;
    pages[2].lParam = reinterpret_cast<LPARAM>(data.get());

    pages[3].dwSize = sizeof(PROPSHEETPAGEW);
    pages[3].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[3].hInstance = hInst;
    pages[3].pResource = backgroundsTemplate.get();
    pages[3].pszTitle = L"Backgrounds";
    pages[3].pfnDlgProc = BackgroundsPageProc;
    pages[3].lParam = reinterpret_cast<LPARAM>(data.get());

    pages[4].dwSize = sizeof(PROPSHEETPAGEW);
    pages[4].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[4].hInstance = hInst;
    pages[4].pResource = contextTemplate.get();
    pages[4].pszTitle = L"Context Menus";
    pages[4].pfnDlgProc = ContextMenusPageProc;
    pages[4].lParam = reinterpret_cast<LPARAM>(data.get());

    pages[5].dwSize = sizeof(PROPSHEETPAGEW);
    pages[5].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[5].hInstance = hInst;
    pages[5].pResource = groupsTemplate.get();
    pages[5].pszTitle = L"Groups";
    pages[5].pfnDlgProc = GroupsPageProc;
    pages[5].lParam = reinterpret_cast<LPARAM>(data.get());

    pages[6].dwSize = sizeof(PROPSHEETPAGEW);
    pages[6].dwFlags = PSP_USETITLE | PSP_DLGINDIRECT;
    pages[6].hInstance = hInst;
    pages[6].pResource = webFoldersTemplate.get();
    pages[6].pszTitle = L"Web Folders";
    pages[6].pfnDlgProc = WebFoldersPageProc;
    pages[6].lParam = reinterpret_cast<LPARAM>(data.get());

    // Create property sheet
    PROPSHEETHEADERW psh{};
    psh.dwSize = sizeof(psh);
    psh.dwFlags = PSH_PROPSHEETPAGE | PSH_NOCONTEXTHELP;
    psh.hwndParent = parent;
    psh.hInstance = hInst;
    psh.pszCaption = L"ShellTabs Options";
    psh.nPages = static_cast<UINT>(pages.size());
    psh.nStartPage = static_cast<UINT>(data->initialTab);
    psh.ppsp = pages.data();

    // Show dialog
    INT_PTR dlgResult = PropertySheetW(&psh);

    // Process result
    if (dlgResult > 0) {
        result.saved = true;
        result.optionsChanged = (data->workingOptions != data->originalOptions);
        result.groupsChanged = data->groupsChanged;

        if (result.optionsChanged) {
            store.Set(data->workingOptions);
            store.Save();
        }

        if (result.groupsChanged) {
            result.savedGroups = data->workingGroups;
            for (size_t i = 0; i < data->workingGroups.size() && i < data->workingGroupIds.size(); ++i) {
                const std::wstring& originalId = data->workingGroupIds[i];
                const std::wstring& newName = data->workingGroups[i].name;
                if (!originalId.empty() && _wcsicmp(originalId.c_str(), newName.c_str()) != 0) {
                    result.renamedGroups.emplace_back(originalId, newName);
                }
            }
            result.removedGroupIds = data->removedGroupIds;
        }
    }

    // Cleanup brushes
    if (data->breadcrumbStartBrush) DeleteObject(data->breadcrumbStartBrush);
    if (data->breadcrumbEndBrush) DeleteObject(data->breadcrumbEndBrush);
    if (data->fontStartBrush) DeleteObject(data->fontStartBrush);
    if (data->fontEndBrush) DeleteObject(data->fontEndBrush);
    if (data->tabSelectedBrush) DeleteObject(data->tabSelectedBrush);
    if (data->tabUnselectedBrush) DeleteObject(data->tabUnselectedBrush);
    if (data->progressStartBrush) DeleteObject(data->progressStartBrush);
    if (data->progressEndBrush) DeleteObject(data->progressEndBrush);
    if (data->glowPrimaryBrush) DeleteObject(data->glowPrimaryBrush);
    if (data->glowSecondaryBrush) DeleteObject(data->glowSecondaryBrush);
    if (data->groupColorBrush) DeleteObject(data->groupColorBrush);

    if (data->universalPreview) DeleteObject(data->universalPreview);
    if (data->folderPreview) DeleteObject(data->folderPreview);

    return result;
}

}  // namespace shelltabs
