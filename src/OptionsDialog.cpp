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
constexpr int kPageHeight = 380;
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

bool BrowseForFolder(HWND parent, std::wstring* path, const wchar_t* title = nullptr) {
    if (!path) return false;

    Microsoft::WRL::ComPtr<IFileDialog> dialog;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&dialog));
    if (FAILED(hr)) return false;

    DWORD options = 0;
    dialog->GetOptions(&options);
    dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

    if (title) dialog->SetTitle(title);

    if (SUCCEEDED(dialog->Show(parent))) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (SUCCEEDED(dialog->GetResult(&item))) {
            wchar_t* filePath = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath))) {
                *path = filePath;
                CoTaskMemFree(filePath);
                return true;
            }
        }
    }
    return false;
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
        WS_CHILD | WS_VISIBLE | SS_NOTIFY);
    builder.AddPushButton(IDC_APP_TAB_SEL_BTN, L"Choose...",
        kMargin + kGroupMargin + 200, y - 2, 70, kButtonHeight);
    y += kButtonHeight + kSpacing;

    builder.AddCheckbox(IDC_APP_TAB_UNSEL_CHECK,
        L"Custom unselected color:",
        kMargin + kGroupMargin, y, 150, kCheckHeight);
    builder.AddStatic(IDC_APP_TAB_UNSEL_PREVIEW, L"",
        kMargin + kGroupMargin + 160, y - 2, kColorBoxSize, kColorBoxSize,
        WS_CHILD | WS_VISIBLE | SS_NOTIFY);
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
            if (dis->CtlID == IDC_APP_TAB_SEL_PREVIEW && data->tabSelectedBrush) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.customTabSelectedColor);
                return TRUE;
            }
            else if (dis->CtlID == IDC_APP_TAB_UNSEL_PREVIEW && data->tabUnselectedBrush) {
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
    DialogBuilder builder(kPageWidth, kPageHeight);

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

    builder.AddCheckbox(IDC_GLOW_SURF_DIRECTUI,
        L"Enable DirectUI glow",
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
        WS_CHILD | WS_VISIBLE | SS_NOTIFY);
    builder.AddPushButton(IDC_GLOW_PRIMARY_BTN, L"Choose...",
        kMargin + 145, y - 2, 80, kButtonHeight);
    y += kButtonHeight + kSpacing;

    builder.AddStatic(IDC_GLOW_SECONDARY_LABEL,
        L"Secondary color:",
        kMargin, y, 100, kLabelHeight);
    builder.AddStatic(IDC_GLOW_SECONDARY_PREVIEW, L"",
        kMargin + 105, y - 2, kColorBoxSize, kColorBoxSize,
        WS_CHILD | WS_VISIBLE | SS_NOTIFY);
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
            if (dis->CtlID == IDC_GLOW_PRIMARY_PREVIEW && data->glowPrimaryBrush) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.neonGlowPrimaryColor);
                return TRUE;
            }
            else if (dis->CtlID == IDC_GLOW_SECONDARY_PREVIEW && data->glowSecondaryBrush) {
                DrawColorBox(dis->hDC, dis->rcItem, data->workingOptions.neonGlowSecondaryColor);
                return TRUE;
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
// BACKGROUNDS PAGE
//=============================================================================

// Helper function to generate unique token for async operations
static UINT64 GenerateToken() {
    static std::atomic<UINT64> counter{1};
    return counter.fetch_add(1, std::memory_order_relaxed);
}

DialogTemplatePtr CreateBackgroundsPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    int y = kMargin;

    // Enable checkbox
    builder.AddCheckbox(IDC_BG_ENABLE,
        L"Enable custom folder backgrounds",
        kMargin, y, 280, kCheckHeight);
    y += kCheckHeight + kSpacing * 2;

    // Universal background group
    builder.AddGroupBox(IDC_BG_UNIVERSAL_GROUP, L"Universal Background",
        kMargin, y, kPageWidth - 2 * kMargin, 100);
    y += 18;

    builder.AddStatic(-1, L"Image:",
        kMargin + kGroupMargin, y, 60, kLabelHeight);
    builder.AddStatic(IDC_BG_UNIVERSAL_PREVIEW, L"",
        kMargin + kGroupMargin + 65, y - 2, kPreviewSize, kPreviewSize,
        WS_CHILD | WS_VISIBLE | SS_NOTIFY);
    y += kPreviewSize + kSpacing;

    builder.AddPushButton(IDC_BG_UNIVERSAL_BROWSE, L"Browse...",
        kMargin + kGroupMargin, y, 80, kButtonHeight);
    builder.AddPushButton(IDC_BG_UNIVERSAL_CLEAR, L"Clear",
        kMargin + kGroupMargin + 90, y, 60, kButtonHeight);
    y += kButtonHeight + 6;

    builder.AddStatic(IDC_BG_UNIVERSAL_NAME, L"(no image selected)",
        kMargin + kGroupMargin, y, kPageWidth - 2 * kMargin - 2 * kGroupMargin, kLabelHeight);
    y += kLabelHeight + kSpacing * 2;

    // Folder overrides group
    builder.AddGroupBox(IDC_BG_FOLDER_GROUP, L"Folder-Specific Backgrounds",
        kMargin, y, kPageWidth - 2 * kMargin, 150);
    y += 18;

    builder.AddListBox(IDC_BG_FOLDER_LIST,
        kMargin + kGroupMargin, y, 200, 80);

    builder.AddStatic(IDC_BG_FOLDER_PREVIEW, L"",
        kMargin + kGroupMargin + 210, y, kPreviewSize, kPreviewSize,
        WS_CHILD | WS_VISIBLE | SS_NOTIFY);
    y += 86;

    builder.AddStatic(IDC_BG_FOLDER_NAME, L"",
        kMargin + kGroupMargin, y, kPageWidth - 2 * kMargin - 2 * kGroupMargin, kLabelHeight);
    y += kLabelHeight + 6;

    builder.AddPushButton(IDC_BG_FOLDER_ADD, L"Add...",
        kMargin + kGroupMargin, y, 60, kButtonHeight);
    builder.AddPushButton(IDC_BG_FOLDER_EDIT, L"Edit...",
        kMargin + kGroupMargin + 70, y, 60, kButtonHeight);
    builder.AddPushButton(IDC_BG_FOLDER_REMOVE, L"Remove",
        kMargin + kGroupMargin + 140, y, 60, kButtonHeight);
    builder.AddPushButton(IDC_BG_FOLDER_CLEAN, L"Clean Up...",
        kMargin + kGroupMargin + 210, y, 80, kButtonHeight);

    return builder.Build();
}

void InitBackgroundsPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    CheckDlgButton(page, IDC_BG_ENABLE,
        data->workingOptions.enableFolderBackgrounds ? BST_CHECKED : BST_UNCHECKED);

    // Set universal background name
    HWND nameLabel = GetDlgItem(page, IDC_BG_UNIVERSAL_NAME);
    if (nameLabel) {
        if (data->workingOptions.universalFolderBackgroundImage.displayName.empty()) {
            SetWindowTextW(nameLabel, L"(no image selected)");
        } else {
            SetWindowTextW(nameLabel, data->workingOptions.universalFolderBackgroundImage.displayName.c_str());
        }
    }

    // Populate folder list
    HWND folderList = GetDlgItem(page, IDC_BG_FOLDER_LIST);
    if (folderList) {
        SendMessageW(folderList, LB_RESETCONTENT, 0, 0);
        for (const auto& entry : data->workingOptions.folderBackgroundEntries) {
            SendMessageW(folderList, LB_ADDSTRING, 0,
                reinterpret_cast<LPARAM>(entry.folderPath.c_str()));
        }
    }

    // Update control states
    bool enabled = data->workingOptions.enableFolderBackgrounds;
    EnableWindow(GetDlgItem(page, IDC_BG_UNIVERSAL_BROWSE), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_UNIVERSAL_CLEAR), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_LIST), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_ADD), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_EDIT), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_REMOVE), enabled);
    EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_CLEAN), enabled);
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

            int id = LOWORD(wParam);
            int code = HIWORD(wParam);

            if (id == IDC_BG_ENABLE && code == BN_CLICKED) {
                data->workingOptions.enableFolderBackgrounds =
                    (IsDlgButtonChecked(page, IDC_BG_ENABLE) == BST_CHECKED);

                bool enabled = data->workingOptions.enableFolderBackgrounds;
                EnableWindow(GetDlgItem(page, IDC_BG_UNIVERSAL_BROWSE), enabled);
                EnableWindow(GetDlgItem(page, IDC_BG_UNIVERSAL_CLEAR), enabled);
                EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_LIST), enabled);
                EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_ADD), enabled);
                EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_EDIT), enabled);
                EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_REMOVE), enabled);
                EnableWindow(GetDlgItem(page, IDC_BG_FOLDER_CLEAN), enabled);

                PropSheet_Changed(GetParent(page), page);
            }
            else if (id == IDC_BG_UNIVERSAL_BROWSE && code == BN_CLICKED) {
                std::wstring path;
                if (BrowseForImage(page, &path, &data->lastImageDir)) {
                    data->workingOptions.universalFolderBackgroundImage.displayName = path;
                    SetWindowTextW(GetDlgItem(page, IDC_BG_UNIVERSAL_NAME), path.c_str());
                    PropSheet_Changed(GetParent(page), page);
                }
            }
            else if (id == IDC_BG_UNIVERSAL_CLEAR && code == BN_CLICKED) {
                data->workingOptions.universalFolderBackgroundImage.displayName.clear();
                data->workingOptions.universalFolderBackgroundImage.cachedImagePath.clear();
                SetWindowTextW(GetDlgItem(page, IDC_BG_UNIVERSAL_NAME), L"(no image selected)");
                PropSheet_Changed(GetParent(page), page);
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
// CONTEXT MENUS PAGE - Simplified
//=============================================================================

DialogTemplatePtr CreateContextMenusPageTemplate() {
    DialogBuilder builder(kPageWidth, kPageHeight);

    int y = kMargin;

    builder.AddStatic(-1,
        L"Context menu customization is a complex feature.",
        kMargin, y, kPageWidth - 2 * kMargin, kLabelHeight);
    y += kLabelHeight + 6;

    builder.AddStatic(-1,
        L"This simplified version shows basic settings.",
        kMargin, y, kPageWidth - 2 * kMargin, kLabelHeight);
    y += kLabelHeight + kSpacing * 2;

    builder.AddStatic(-1,
        L"Template:",
        kMargin, y, 100, kLabelHeight);
    builder.AddComboBox(IDC_CTX_TEMPLATE,
        kMargin + 105, y - 2, 200, kComboHeight);
    y += kEditHeight + kSpacing * 2;

    // Tree view for menu items
    builder.AddTreeView(IDC_CTX_TREE,
        kMargin, y, kPageWidth - 2 * kMargin, 200);
    y += 205;

    // Control buttons
    builder.AddPushButton(IDC_CTX_ADD_COMMAND, L"Add Command",
        kMargin, y, 90, kButtonHeight);
    builder.AddPushButton(IDC_CTX_ADD_SUBMENU, L"Add Submenu",
        kMargin + 100, y, 90, kButtonHeight);
    builder.AddPushButton(IDC_CTX_ADD_SEPARATOR, L"Add Separator",
        kMargin + 200, y, 90, kButtonHeight);
    y += kButtonHeight + 6;

    builder.AddPushButton(IDC_CTX_REMOVE, L"Remove",
        kMargin, y, 90, kButtonHeight);
    builder.AddPushButton(IDC_CTX_MOVE_UP, L"Move Up",
        kMargin + 100, y, 90, kButtonHeight);
    builder.AddPushButton(IDC_CTX_MOVE_DOWN, L"Move Down",
        kMargin + 200, y, 90, kButtonHeight);

    return builder.Build();
}

// Helper: Store ContextMenuItem pointer in tree item's LPARAM
void PopulateContextMenuTree(HWND tree, const std::vector<ContextMenuItem>& items, HTREEITEM parent = TVI_ROOT) {
    for (size_t i = 0; i < items.size(); ++i) {
        const auto& item = items[i];

        TVINSERTSTRUCTW insert{};
        insert.hParent = parent;
        insert.hInsertAfter = TVI_LAST;
        insert.item.mask = TVIF_TEXT | TVIF_PARAM;

        std::wstring displayText;
        if (item.type == ContextMenuItemType::kSeparator) {
            displayText = L"─────────────";
        } else {
            displayText = item.label.empty() ? L"(Unnamed)" : item.label;
        }

        insert.item.pszText = const_cast<wchar_t*>(displayText.c_str());
        // Store index in LPARAM for retrieval
        insert.item.lParam = static_cast<LPARAM>(i);

        HTREEITEM hItem = TreeView_InsertItem(tree, &insert);

        // Recursively add children for submenus
        if (item.type == ContextMenuItemType::kSubmenu && !item.children.empty()) {
            PopulateContextMenuTree(tree, item.children, hItem);
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
        std::wstring text = L"(No custom menu items - use buttons below to add)";
        insert.item.pszText = text.data();
        TreeView_InsertItem(tree, &insert);
    } else {
        PopulateContextMenuTree(tree, data->workingOptions.contextMenuItems);
    }
}

void InitContextMenusPage(HWND page, OptionsDialogData* data) {
    if (!data) return;

    HWND combo = GetDlgItem(page, IDC_CTX_TEMPLATE);
    if (combo) {
        ComboBox_AddString(combo, L"Empty");
        ComboBox_AddString(combo, L"Command Prompt Here");
        ComboBox_AddString(combo, L"PowerShell Here");
        ComboBox_AddString(combo, L"Open with VS Code");
        ComboBox_SetCurSel(combo, 0);
    }

    // Initialize window state combo
    HWND windowStateCombo = GetDlgItem(page, IDC_CTX_WINDOW_STATE);
    if (windowStateCombo) {
        ComboBox_AddString(windowStateCombo, L"Normal");
        ComboBox_AddString(windowStateCombo, L"Minimized");
        ComboBox_AddString(windowStateCombo, L"Maximized");
        ComboBox_AddString(windowStateCombo, L"Hidden");
        ComboBox_SetCurSel(windowStateCombo, 0);
    }

    // Initialize anchor combo
    HWND anchorCombo = GetDlgItem(page, IDC_CTX_ANCHOR);
    if (anchorCombo) {
        ComboBox_AddString(anchorCombo, L"Default");
        ComboBox_AddString(anchorCombo, L"Top");
        ComboBox_AddString(anchorCombo, L"Bottom");
        ComboBox_AddString(anchorCombo, L"Before Shell Items");
        ComboBox_AddString(anchorCombo, L"After Shell Items");
        ComboBox_SetCurSel(anchorCombo, 0);
    }

    RefreshContextMenuTree(page, data);
}

// Helper to create template menu items
ContextMenuItem CreateCommandPromptMenuItem() {
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

ContextMenuItem CreatePowerShellMenuItem() {
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

ContextMenuItem CreateVSCodeMenuItem() {
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

            // Handle template selection
            if (id == IDC_CTX_TEMPLATE && code == CBN_SELCHANGE) {
                HWND combo = GetDlgItem(page, IDC_CTX_TEMPLATE);
                int sel = ComboBox_GetCurSel(combo);

                ContextMenuItem newItem;
                bool addItem = false;

                switch (sel) {
                    case 1: // Command Prompt
                        newItem = CreateCommandPromptMenuItem();
                        addItem = true;
                        break;
                    case 2: // PowerShell
                        newItem = CreatePowerShellMenuItem();
                        addItem = true;
                        break;
                    case 3: // VS Code
                        newItem = CreateVSCodeMenuItem();
                        addItem = true;
                        break;
                }

                if (addItem) {
                    data->workingOptions.contextMenuItems.push_back(newItem);
                    RefreshContextMenuTree(page, data);
                    ComboBox_SetCurSel(combo, 0); // Reset to "Empty"
                    PropSheet_Changed(GetParent(page), page);
                }
                return TRUE;
            }

            // Add Command button
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

                data->workingOptions.contextMenuItems.push_back(item);
                RefreshContextMenuTree(page, data);
                PropSheet_Changed(GetParent(page), page);
                return TRUE;
            }

            // Add Submenu button
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

            // Add Separator button
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

            // Remove button
            if (id == IDC_CTX_REMOVE) {
                HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                HTREEITEM selected = TreeView_GetSelection(tree);

                if (selected && !data->workingOptions.contextMenuItems.empty()) {
                    TVITEMW item{};
                    item.hItem = selected;
                    item.mask = TVIF_PARAM;

                    if (TreeView_GetItem(tree, &item)) {
                        size_t index = static_cast<size_t>(item.lParam);
                        if (index < data->workingOptions.contextMenuItems.size()) {
                            data->workingOptions.contextMenuItems.erase(
                                data->workingOptions.contextMenuItems.begin() + index
                            );
                            RefreshContextMenuTree(page, data);
                            PropSheet_Changed(GetParent(page), page);
                        }
                    }
                }
                return TRUE;
            }

            // Move Up button
            if (id == IDC_CTX_MOVE_UP) {
                HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                HTREEITEM selected = TreeView_GetSelection(tree);

                if (selected && !data->workingOptions.contextMenuItems.empty()) {
                    TVITEMW item{};
                    item.hItem = selected;
                    item.mask = TVIF_PARAM;

                    if (TreeView_GetItem(tree, &item)) {
                        size_t index = static_cast<size_t>(item.lParam);
                        if (index > 0 && index < data->workingOptions.contextMenuItems.size()) {
                            std::swap(data->workingOptions.contextMenuItems[index],
                                    data->workingOptions.contextMenuItems[index - 1]);
                            RefreshContextMenuTree(page, data);
                            PropSheet_Changed(GetParent(page), page);
                        }
                    }
                }
                return TRUE;
            }

            // Move Down button
            if (id == IDC_CTX_MOVE_DOWN) {
                HWND tree = GetDlgItem(page, IDC_CTX_TREE);
                HTREEITEM selected = TreeView_GetSelection(tree);

                if (selected && !data->workingOptions.contextMenuItems.empty()) {
                    TVITEMW item{};
                    item.hItem = selected;
                    item.mask = TVIF_PARAM;

                    if (TreeView_GetItem(tree, &item)) {
                        size_t index = static_cast<size_t>(item.lParam);
                        if (index < data->workingOptions.contextMenuItems.size() - 1) {
                            std::swap(data->workingOptions.contextMenuItems[index],
                                    data->workingOptions.contextMenuItems[index + 1]);
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

            // Handle tree view selection changes
            if (nmhdr->idFrom == IDC_CTX_TREE && nmhdr->code == TVN_SELCHANGED) {
                // TODO: Load selected item properties into edit controls
                return TRUE;
            }

            if (nmhdr->code == PSN_APPLY && data) {
                // Changes are already in workingOptions, which will be saved by the main dialog
                return PSNRET_NOERROR;
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

    builder.AddStatic(-1, L"Saved Groups:",
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
            label = CreateWindowW(L"STATIC", L"Folder Paths:",
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
            HWND btnAdd = CreateWindowW(L"BUTTON", L"Add Path...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                btnX, y, 100, 24,
                dialog, reinterpret_cast<HMENU>(IDC_GRP_ED_ADD), hInst, nullptr);
            SendMessageW(btnAdd, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            btnX += 105;

            HWND btnEdit = CreateWindowW(L"BUTTON", L"Edit Path...",
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

            HWND colorPreview = CreateWindowW(L"STATIC", L"",
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
                // Browse for folder
                BROWSEINFOW bi{};
                bi.hwndOwner = dialog;
                bi.lpszTitle = L"Select folder to add to group:";
                bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;

                PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
                if (pidl) {
                    wchar_t path[MAX_PATH];
                    if (SHGetPathFromIDListW(pidl, path)) {
                        HWND pathList = GetDlgItem(dialog, IDC_GRP_ED_PATHS);
                        data->group->tabPaths.push_back(path);
                        SendMessageW(pathList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(path));
                    }
                    CoTaskMemFree(pidl);
                }
            }
            else if (id == IDC_GRP_ED_EDIT_PATH && code == BN_CLICKED) {
                HWND pathList = GetDlgItem(dialog, IDC_GRP_ED_PATHS);
                int sel = static_cast<int>(SendMessageW(pathList, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < static_cast<int>(data->group->tabPaths.size())) {
                    BROWSEINFOW bi{};
                    bi.hwndOwner = dialog;
                    bi.lpszTitle = L"Select folder:";
                    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;

                    PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
                    if (pidl) {
                        wchar_t path[MAX_PATH];
                        if (SHGetPathFromIDListW(pidl, path)) {
                            data->group->tabPaths[sel] = path;
                            SendMessageW(pathList, LB_DELETESTRING, sel, 0);
                            SendMessageW(pathList, LB_INSERTSTRING, sel, reinterpret_cast<LPARAM>(path));
                            SendMessageW(pathList, LB_SETCURSEL, sel, 0);
                        }
                        CoTaskMemFree(pidl);
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
                    std::wstring msg = L"Remove group '" + name + L"'?";
                    if (MessageBoxW(page, msg.c_str(), L"Confirm", MB_YESNO | MB_ICONQUESTION) == IDYES) {
                        data->removedGroupIds.push_back(name);
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

    if (!generalTemplate || !appearanceTemplate || !glowTemplate ||
        !backgroundsTemplate || !contextTemplate || !groupsTemplate) {
        MessageBoxW(parent, L"Failed to create dialog templates.", L"Error", MB_OK | MB_ICONERROR);
        return result;
    }

    // Create property sheet pages
    std::array<PROPSHEETPAGEW, 6> pages{};
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
