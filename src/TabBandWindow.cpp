#include "TabBandWindow.h"

#include <algorithm>
#include <atomic>
#include <cwchar>
#include <cstdlib>
#include <memory>
#include <new>
#include <string_view>
#include <vector>

#include <CommCtrl.h>
#include <Ole2.h>
#include <ShlObj.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <windowsx.h>
#include <winreg.h>
#include <vssym32.h>

#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

#ifndef BTNS_FIXEDSIZE
#define BTNS_FIXEDSIZE 0x0020
#endif

#include "Module.h"
#include "TabBand.h"

namespace shelltabs {

namespace {
const wchar_t kWindowClassName[] = L"ShellTabsNativeToolbarHost";
constexpr int kNewTabCommandId = 40000;
constexpr int kMaxTooltip = 512;

enum class PreferredAppMode {
    Default,
    AllowDark,
    ForceDark,
    ForceLight,
    Max,
};

using ShouldAppsUseDarkModeFunc = BOOL(WINAPI*)();
using AllowDarkModeForWindowFunc = BOOL(WINAPI*)(HWND, BOOL);
using AllowDarkModeForAppFunc = BOOL(WINAPI*)(BOOL);
using RefreshImmersiveColorPolicyStateFunc = void(WINAPI*)();
using SetPreferredAppModeFunc = PreferredAppMode(WINAPI*)(PreferredAppMode);

struct ThemeApi {
    HMODULE module = nullptr;
    ShouldAppsUseDarkModeFunc shouldAppsUseDarkMode = nullptr;
    AllowDarkModeForWindowFunc allowDarkModeForWindow = nullptr;
    AllowDarkModeForAppFunc allowDarkModeForApp = nullptr;
    RefreshImmersiveColorPolicyStateFunc refreshImmersiveColorPolicyState = nullptr;
    SetPreferredAppModeFunc setPreferredAppMode = nullptr;
    bool preferredAppModeInitialized = false;
};

FARPROC LoadThemeProcedure(HMODULE module, const char* name, WORD ordinal) {
    if (!module) {
        return nullptr;
    }
    FARPROC proc = nullptr;
    if (name) {
        proc = GetProcAddress(module, name);
    }
    if (!proc && ordinal != 0) {
        proc = GetProcAddress(module, MAKEINTRESOURCEA(ordinal));
    }
    return proc;
}

ThemeApi& GetThemeApi() {
    static ThemeApi api;
    if (!api.module) {
        api.module = LoadLibraryW(L"uxtheme.dll");
        if (api.module) {
            api.shouldAppsUseDarkMode = reinterpret_cast<ShouldAppsUseDarkModeFunc>(
                LoadThemeProcedure(api.module, "ShouldAppsUseDarkMode", 132));
            api.allowDarkModeForWindow = reinterpret_cast<AllowDarkModeForWindowFunc>(
                LoadThemeProcedure(api.module, "AllowDarkModeForWindow", 133));
            api.refreshImmersiveColorPolicyState = reinterpret_cast<RefreshImmersiveColorPolicyStateFunc>(
                LoadThemeProcedure(api.module, "RefreshImmersiveColorPolicyState", 104));
            api.setPreferredAppMode = reinterpret_cast<SetPreferredAppModeFunc>(
                LoadThemeProcedure(api.module, "SetPreferredAppMode", 135));
            if (!api.setPreferredAppMode) {
                api.allowDarkModeForApp = reinterpret_cast<AllowDarkModeForAppFunc>(
                    LoadThemeProcedure(api.module, "AllowDarkModeForApp", 135));
            }
        }
    }
    return api;
}

void EnsurePreferredAppMode() {
    auto& api = GetThemeApi();
    if (api.preferredAppModeInitialized) {
        return;
    }
    if (api.setPreferredAppMode) {
        api.setPreferredAppMode(PreferredAppMode::AllowDark);
        api.preferredAppModeInitialized = true;
        return;
    }
    if (api.allowDarkModeForApp) {
        api.allowDarkModeForApp(TRUE);
        api.preferredAppModeInitialized = true;
    }
}

COLORREF LightenColor(COLORREF color, double factor) {
    factor = std::clamp(factor, 0.0, 1.0);
    const auto adjust = [factor](BYTE component) -> BYTE {
        const double result = static_cast<double>(component) + (255.0 - static_cast<double>(component)) * factor;
        return static_cast<BYTE>(std::clamp(result, 0.0, 255.0));
    };
    return RGB(adjust(GetRValue(color)), adjust(GetGValue(color)), adjust(GetBValue(color)));
}

COLORREF DarkenColor(COLORREF color, double factor) {
    factor = std::clamp(factor, 0.0, 1.0);
    const auto adjust = [factor](BYTE component) -> BYTE {
        const double result = static_cast<double>(component) * (1.0 - factor);
        return static_cast<BYTE>(std::clamp(result, 0.0, 255.0));
    };
    return RGB(adjust(GetRValue(color)), adjust(GetGValue(color)), adjust(GetBValue(color)));
}

COLORREF AdjustIndicatorColorForState(COLORREF base, bool hot, bool pressed) {
    if (pressed) {
        return DarkenColor(base, 0.2);
    }
    if (hot) {
        return LightenColor(base, 0.15);
    }
    return base;
}

int ToolbarHitTest(HWND toolbar, POINT pt) {
    if (!toolbar) {
        return -1;
    }

    const LRESULT buttonCount = SendMessageW(toolbar, TB_BUTTONCOUNT, 0, 0);
    for (LRESULT index = 0; index < buttonCount; ++index) {
        RECT rect{};
        if (SendMessageW(toolbar, TB_GETITEMRECT, static_cast<WPARAM>(index),
                         reinterpret_cast<LPARAM>(&rect))) {
            if (PtInRect(&rect, pt)) {
                return static_cast<int>(index);
            }
        }
    }

    return -1;
}

HICON LoadItemIcon(const TabViewItem& item, int iconSize) {
    if (!item.pidl) {
        return nullptr;
    }

    SHFILEINFOW info{};
    UINT flags = SHGFI_PIDL | SHGFI_ICON;
    if (iconSize <= GetSystemMetrics(SM_CXSMICON)) {
        flags |= SHGFI_SMALLICON;
    } else {
        flags |= SHGFI_LARGEICON;
    }

    if (SHGetFileInfoW(reinterpret_cast<LPCWSTR>(item.pidl), 0, &info, sizeof(info), flags) == 0) {
        return nullptr;
    }
    return info.hIcon;
}

void RegisterWindowClass() {
    static bool registered = false;
    if (registered) {
        return;
    }

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.style = CS_DBLCLKS | CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = DefWindowProcW;
    wc.hInstance = GetModuleHandleInstance();
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kWindowClassName;

    if (RegisterClassExW(&wc)) {
        registered = true;
    }
}

int ToolbarIconSize() {
    const int smallIconSize = GetSystemMetrics(SM_CXSMICON);
    if (smallIconSize > 0) {
        return smallIconSize;
    }
    return 16;
}

bool SupportsFileDrop(IDataObject* dataObject) {
    if (!dataObject) {
        return false;
    }
    FORMATETC format{};
    format.cfFormat = CF_HDROP;
    format.dwAspect = DVASPECT_CONTENT;
    format.lindex = -1;
    format.tymed = TYMED_HGLOBAL;
    return SUCCEEDED(dataObject->QueryGetData(&format));
}

}  // namespace

class TabToolbarDropTarget : public IDropTarget {
public:
    explicit TabToolbarDropTarget(TabBandWindow* window) : m_window(window) {}

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDropTarget) {
            *object = static_cast<IDropTarget*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return m_refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        const ULONG remaining = m_refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (remaining == 0) {
            delete this;
            return 0;
        }
        return remaining;
    }

    // IDropTarget
    IFACEMETHODIMP DragEnter(IDataObject* dataObject, DWORD keyState, POINTL point, DWORD* effect) override {
        m_hasFileData = SupportsFileDrop(dataObject);
        return UpdateEffect(keyState, point, effect);
    }

    IFACEMETHODIMP DragOver(DWORD keyState, POINTL point, DWORD* effect) override {
        return UpdateEffect(keyState, point, effect);
    }

    IFACEMETHODIMP DragLeave() override {
        m_hasFileData = false;
        return S_OK;
    }

    IFACEMETHODIMP Drop(IDataObject* dataObject, DWORD keyState, POINTL point, DWORD* effect) override {
        if (!effect) {
            return E_INVALIDARG;
        }
        *effect = DROPEFFECT_NONE;
        if (!m_window || !m_hasFileData) {
            return S_OK;
        }

        POINT screenPt{static_cast<LONG>(point.x), static_cast<LONG>(point.y)};
        TabLocation location = m_window->TabLocationFromPoint(screenPt);
        if (!location.IsValid()) {
            return S_OK;
        }

        std::vector<std::wstring> paths;
        if (!ExtractPaths(dataObject, &paths) || paths.empty()) {
            return S_OK;
        }

        const bool move = (keyState & MK_SHIFT) != 0;
        m_window->HandleFilesDropped(location, paths, move);
        *effect = move ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
        return S_OK;
    }

private:
    HRESULT UpdateEffect(DWORD keyState, const POINTL& point, DWORD* effect) {
        if (!effect) {
            return E_INVALIDARG;
        }
        *effect = DROPEFFECT_NONE;
        if (!m_window || !m_hasFileData) {
            return S_OK;
        }

        POINT screenPt{static_cast<LONG>(point.x), static_cast<LONG>(point.y)};
        const TabViewItem* item = m_window->ItemFromPoint(screenPt);
        if (!item || item->type != TabViewItemType::kTab) {
            return S_OK;
        }

        *effect = (keyState & MK_SHIFT) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
        return S_OK;
    }

    bool ExtractPaths(IDataObject* dataObject, std::vector<std::wstring>* paths) const {
        if (!dataObject || !paths) {
            return false;
        }

        FORMATETC format{};
        format.cfFormat = CF_HDROP;
        format.dwAspect = DVASPECT_CONTENT;
        format.lindex = -1;
        format.tymed = TYMED_HGLOBAL;

        STGMEDIUM medium{};
        if (FAILED(dataObject->GetData(&format, &medium))) {
            return false;
        }

        bool success = false;
        if (medium.tymed == TYMED_HGLOBAL && medium.hGlobal) {
            HDROP drop = static_cast<HDROP>(medium.hGlobal);
            const UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
            for (UINT i = 0; i < count; ++i) {
                const UINT required = DragQueryFileW(drop, i, nullptr, 0);
                if (required == 0) {
                    continue;
                }
                std::wstring path;
                path.resize(required + 1);
                if (DragQueryFileW(drop, i, path.data(), static_cast<UINT>(path.size()))) {
                    path.resize(wcsnlen_s(path.c_str(), path.size()));
                    if (!path.empty()) {
                        paths->emplace_back(std::move(path));
                    }
                }
            }
            success = !paths->empty();
        }

        ReleaseStgMedium(&medium);
        return success;
    }

    std::atomic<ULONG> m_refCount{1};
    TabBandWindow* m_window = nullptr;
    bool m_hasFileData = false;
};

TabBandWindow::TabBandWindow(TabBand* owner) : m_owner(owner) {}

TabBandWindow::~TabBandWindow() { Destroy(); }

HWND TabBandWindow::Create(HWND parent) {
    RegisterWindowClass();

    if (m_hwnd) {
        return m_hwnd;
    }

    HWND hwnd = CreateWindowExW(0, kWindowClassName, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_TABSTOP,
                                0, 0, 0, 0, parent, nullptr, GetModuleHandleInstance(), nullptr);
    if (!hwnd) {
        return nullptr;
    }

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
    SetWindowSubclass(hwnd, WndProc, reinterpret_cast<UINT_PTR>(this), reinterpret_cast<DWORD_PTR>(this));

    m_hwnd = hwnd;
    UpdateTheme();
    EnsureToolbar();
    return m_hwnd;
}

void TabBandWindow::Destroy() {
    DestroyToolbar();
    if (m_hwnd) {
        RemoveWindowSubclass(m_hwnd, WndProc, reinterpret_cast<UINT_PTR>(this));
        DestroyWindow(m_hwnd);
        m_hwnd = nullptr;
    }
}

void TabBandWindow::Show(bool show) {
    if (!m_hwnd) {
        return;
    }
    ShowWindow(m_hwnd, show ? SW_SHOW : SW_HIDE);
}

void TabBandWindow::SetTabs(const std::vector<TabViewItem>& items) {
    m_tabData = items;
    RebuildToolbar();
}

bool TabBandWindow::HasFocus() const {
    HWND focus = GetFocus();
    return focus == m_toolbar || focus == m_hwnd;
}

void TabBandWindow::FocusTab() {
    if (!m_hwnd) {
        return;
    }
    if (m_toolbar) {
        SetFocus(m_toolbar);
    } else {
        SetFocus(m_hwnd);
    }
}

void TabBandWindow::EnsureToolbar() {
    if (m_toolbar || !m_hwnd) {
        return;
    }

    DWORD style = WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_LIST | TBSTYLE_TOOLTIPS | CCS_NOPARENTALIGN |
                  CCS_NORESIZE | CCS_NODIVIDER | CCS_ADJUSTABLE;
    DWORD exStyle = TBSTYLE_EX_MIXEDBUTTONS | TBSTYLE_EX_HIDECLIPPEDBUTTONS | TBSTYLE_EX_DOUBLEBUFFER;

    HWND toolbar = CreateWindowExW(0, TOOLBARCLASSNAMEW, L"", style, 0, 0, 0, 0, m_hwnd, nullptr,
                                   GetModuleHandleInstance(), nullptr);
    if (!toolbar) {
        return;
    }

    SendMessageW(toolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
    SendMessageW(toolbar, TB_SETEXTENDEDSTYLE, 0, exStyle);

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(m_hwnd, WM_GETFONT, 0, 0));
    if (!font) {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    if (font) {
        SendMessageW(toolbar, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }

    SetWindowSubclass(toolbar, ToolbarWndProc, reinterpret_cast<UINT_PTR>(this), reinterpret_cast<DWORD_PTR>(this));

    m_toolbar = toolbar;
    ConfigureToolbarMetrics();
    RegisterDropTarget();
    UpdateTheme();
    RebuildToolbar();
}

void TabBandWindow::DestroyToolbar() {
    if (!m_toolbar) {
        return;
    }

    ResetCloseTracking();
    RevokeDropTarget();
    RemoveWindowSubclass(m_toolbar, ToolbarWndProc, reinterpret_cast<UINT_PTR>(this));
    ClearImageList();
    DestroyWindow(m_toolbar);
    m_toolbar = nullptr;
}

void TabBandWindow::ClearToolbar() {
    if (!m_toolbar) {
        return;
    }

    const int count = static_cast<int>(SendMessageW(m_toolbar, TB_BUTTONCOUNT, 0, 0));
    for (int i = count - 1; i >= 0; --i) {
        SendMessageW(m_toolbar, TB_DELETEBUTTON, i, 0);
    }
    m_commandMap.clear();
    m_commandToIndex.clear();
    m_nextCommandId = 41000;
}

void TabBandWindow::ClearImageList() {
    if (m_imageList) {
        ImageList_Destroy(m_imageList);
        m_imageList = nullptr;
    }
    if (m_toolbar) {
        SendMessageW(m_toolbar, TB_SETIMAGELIST, 0, 0);
    }
}

void TabBandWindow::ConfigureToolbarMetrics() {
    if (!m_toolbar) {
        return;
    }
    const UINT dpi = CurrentDpi();
    const int horizontalPadding = 0;
    const int verticalPadding = std::max(0, MulDiv(3, static_cast<int>(dpi), 96));
    SendMessageW(m_toolbar, TB_SETPADDING, 0, MAKELPARAM(horizontalPadding, verticalPadding));
    SendMessageW(m_toolbar, TB_SETINDENT, 0, 0);
}

int TabBandWindow::AppendImage(HICON icon) {
    if (!icon) {
        return I_IMAGENONE;
    }

    const int iconSize = ToolbarIconSize();
    if (!m_imageList) {
        m_imageList = ImageList_Create(iconSize, iconSize, ILC_COLOR32 | ILC_MASK, 16, 16);
        if (m_toolbar) {
            SendMessageW(m_toolbar, TB_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(m_imageList));
        }
    }
    if (!m_imageList) {
        DestroyIcon(icon);
        return I_IMAGENONE;
    }

    const int index = ImageList_AddIcon(m_imageList, icon);
    DestroyIcon(icon);
    if (index < 0) {
        return I_IMAGENONE;
    }
    return index;
}

void TabBandWindow::RebuildToolbar() {
    if (!m_hwnd) {
        return;
    }
    EnsureToolbar();
    if (!m_toolbar) {
        return;
    }

    ResetCloseTracking();
    ConfigureToolbarMetrics();

    SendMessageW(m_toolbar, WM_SETREDRAW, FALSE, 0);
    ClearToolbar();
    ClearImageList();

    std::vector<TBBUTTON> buttons;
    buttons.reserve(m_tabData.size() + 1);
    struct GroupIndicatorCommand {
        int commandId = -1;
        bool collapsed = false;
    };
    std::vector<GroupIndicatorCommand> groupIndicatorCommands;
    struct TabButtonWidth {
        int commandId = -1;
        int width = 0;
    };
    std::vector<TabButtonWidth> tabButtonWidths;

    for (size_t index = 0; index < m_tabData.size(); ++index) {
        const auto& item = m_tabData[index];
        if (item.type == TabViewItemType::kGroupHeader) {
            if (!item.headerVisible) {
                continue;
            }
            TBBUTTON button{};
            const int commandId = m_nextCommandId++;
            button.idCommand = commandId;
            button.fsStyle = BTNS_BUTTON | BTNS_FIXEDSIZE;
            button.fsState = TBSTATE_ENABLED;
            button.iBitmap = I_IMAGENONE;
            button.dwData = static_cast<DWORD_PTR>(index);
            button.iString = static_cast<INT_PTR>(SendMessageW(m_toolbar, TB_ADDSTRINGW, 0,
                                                               reinterpret_cast<LPARAM>(item.name.c_str())));
            buttons.push_back(button);
            m_commandToIndex[commandId] = index;
            groupIndicatorCommands.push_back(GroupIndicatorCommand{commandId, item.collapsed});
        } else {
            TBBUTTON button{};
            const int commandId = m_nextCommandId++;
            button.idCommand = commandId;
            button.fsStyle = BTNS_BUTTON | BTNS_SHOWTEXT | BTNS_CHECKGROUP;
            button.fsState = TBSTATE_ENABLED;
            if (item.selected) {
                button.fsState |= TBSTATE_CHECKED;
            }
            button.dwData = static_cast<DWORD_PTR>(index);
            button.iString = static_cast<INT_PTR>(SendMessageW(m_toolbar, TB_ADDSTRINGW, 0,
                                                               reinterpret_cast<LPARAM>(item.name.c_str())));
            if (item.pidl) {
                HICON icon = LoadItemIcon(item, ToolbarIconSize());
                button.iBitmap = AppendImage(icon);
            } else {
                button.iBitmap = I_IMAGENONE;
            }
            buttons.push_back(button);
            m_commandMap.emplace(commandId, item.location);
            m_commandToIndex[commandId] = index;
            tabButtonWidths.push_back(TabButtonWidth{commandId, CalculateTabButtonWidth(item)});
        }
    }

    if (!buttons.empty()) {
        SendMessageW(m_toolbar, TB_ADDBUTTONS, static_cast<WPARAM>(buttons.size()),
                     reinterpret_cast<LPARAM>(buttons.data()));
    }

    if (!groupIndicatorCommands.empty()) {
        const int indicatorWidth = GroupIndicatorWidth();
        for (const auto& entry : groupIndicatorCommands) {
            const int hitWidth = GroupIndicatorHitWidth(entry.collapsed);
            TBBUTTONINFO info{};
            info.cbSize = sizeof(info);
            info.dwMask = TBIF_SIZE | TBIF_STYLE | TBIF_STATE;
            info.fsStyle = BTNS_BUTTON | BTNS_FIXEDSIZE;
            info.fsState = TBSTATE_ENABLED;
            info.cx = static_cast<WORD>(entry.collapsed ? hitWidth : indicatorWidth);
            SendMessageW(m_toolbar, TB_SETBUTTONINFO, entry.commandId, reinterpret_cast<LPARAM>(&info));
        }
    }

    if (!tabButtonWidths.empty()) {
        for (const auto& entry : tabButtonWidths) {
            if (entry.commandId == -1 || entry.width <= 0) {
                continue;
            }
            TBBUTTONINFO info{};
            info.cbSize = sizeof(info);
            info.dwMask = TBIF_SIZE;
            const int clampedWidth = std::clamp(entry.width, 0, 0xFFFF);
            info.cx = static_cast<WORD>(clampedWidth);
            SendMessageW(m_toolbar, TB_SETBUTTONINFO, entry.commandId, reinterpret_cast<LPARAM>(&info));
        }
    }

    // Add new tab button at the end.
    TBBUTTON newTab{};
    newTab.idCommand = kNewTabCommandId;
    newTab.fsStyle = BTNS_BUTTON | BTNS_AUTOSIZE;
    newTab.fsState = TBSTATE_ENABLED;
    newTab.iBitmap = I_IMAGENONE;
    newTab.iString = static_cast<INT_PTR>(SendMessageW(m_toolbar, TB_ADDSTRINGW, 0,
                                                       reinterpret_cast<LPARAM>(L"+")));
    SendMessageW(m_toolbar, TB_ADDBUTTONS, 1, reinterpret_cast<LPARAM>(&newTab));

    SendMessageW(m_toolbar, TB_AUTOSIZE, 0, 0);
    SendMessageW(m_toolbar, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(m_toolbar, nullptr, TRUE);
    UpdateCheckedState();
}

void TabBandWindow::UpdateCheckedState() {
    if (!m_toolbar) {
        return;
    }
    for (const auto& entry : m_commandMap) {
        const int commandId = entry.first;
        const TabLocation location = entry.second;
        bool checked = false;
        if (location.IsValid()) {
            if (location.groupIndex < static_cast<int>(m_tabData.size())) {
                const TabViewItem* item = ItemForCommand(commandId);
                if (item) {
                    checked = item->selected;
                }
            }
        }
        SendMessageW(m_toolbar, TB_CHECKBUTTON, commandId, MAKELPARAM(checked ? TRUE : FALSE, 0));
    }
}

TabLocation TabBandWindow::LocationForCommand(int commandId) const {
    auto it = m_commandMap.find(commandId);
    if (it != m_commandMap.end()) {
        return it->second;
    }
    return {};
}

const TabViewItem* TabBandWindow::ItemForCommand(int commandId) const {
    auto it = m_commandToIndex.find(commandId);
    if (it == m_commandToIndex.end()) {
        return nullptr;
    }
    size_t index = it->second;
    if (index >= m_tabData.size()) {
        return nullptr;
    }
    return &m_tabData[index];
}

void TabBandWindow::HandleToolbarCommand(int commandId) {
    if (!m_owner) {
        return;
    }
    if (commandId == kNewTabCommandId) {
        m_owner->OnNewTabRequested();
        return;
    }

    const TabLocation location = LocationForCommand(commandId);
    if (!location.IsValid()) {
        const TabViewItem* item = ItemForCommand(commandId);
        if (item && item->type == TabViewItemType::kGroupHeader) {
            m_owner->OnToggleGroupCollapsed(item->location.groupIndex);
        }
        return;
    }

    m_owner->OnTabSelected(location);
}

void TabBandWindow::HandleContextMenu(int commandId, const POINT& screenPt) {
    if (!m_owner) {
        return;
    }

    const TabViewItem* item = ItemForCommand(commandId);
    HMENU menu = CreatePopupMenu();
    if (!menu) {
        return;
    }

    enum MenuId {
        kMenuNewTab = 1,
        kMenuCloseTab,
        kMenuDetachTab,
        kMenuHideTab,
        kMenuCloneTab,
        kMenuToggleGroup,
        kMenuDetachGroup,
        kMenuCreateGroupAfter,
        kMenuUnhideAll,
        kMenuToggleSplit,
        kMenuOpenTerminal,
        kMenuOpenVSCode,
        kMenuCopyPath,
    };

    if (item && item->type == TabViewItemType::kTab) {
        AppendMenuW(menu, MF_STRING, kMenuCloseTab, L"Close Tab");
        AppendMenuW(menu, MF_STRING, kMenuDetachTab, L"Detach Tab");
        AppendMenuW(menu, MF_STRING, kMenuCloneTab, L"Duplicate Tab");
        AppendMenuW(menu, MF_STRING, kMenuHideTab, L"Hide Tab");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, kMenuOpenTerminal, L"Open Terminal Here");
        AppendMenuW(menu, MF_STRING, kMenuOpenVSCode, L"Open in VS Code");
        AppendMenuW(menu, MF_STRING, kMenuCopyPath, L"Copy Path");
    } else if (item && item->type == TabViewItemType::kGroupHeader) {
        AppendMenuW(menu, MF_STRING, kMenuToggleGroup, item->collapsed ? L"Expand Group" : L"Collapse Group");
        AppendMenuW(menu, MF_STRING, kMenuCreateGroupAfter, L"Create Group After");
        AppendMenuW(menu, MF_STRING, kMenuDetachGroup, L"Detach Group");
        AppendMenuW(menu, MF_STRING, kMenuUnhideAll, L"Unhide All Tabs");
        if (item->splitAvailable) {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(menu, MF_STRING, kMenuToggleSplit, item->splitActive ? L"Disable Split View" : L"Enable Split View");
        }
    }

    if (GetMenuItemCount(menu) == 0) {
        DestroyMenu(menu);
        return;
    }

    const UINT command = TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD, screenPt.x, screenPt.y, 0, m_hwnd, nullptr);
    DestroyMenu(menu);
    if (command == 0) {
        return;
    }

    if (item && item->type == TabViewItemType::kTab) {
        const TabLocation location = item->location;
        switch (command) {
            case kMenuCloseTab:
                m_owner->OnCloseTabRequested(location);
                break;
            case kMenuDetachTab:
                m_owner->OnDetachTabRequested(location);
                break;
            case kMenuHideTab:
                m_owner->OnHideTabRequested(location);
                break;
            case kMenuCloneTab:
                m_owner->OnCloneTabRequested(location);
                break;
            case kMenuOpenTerminal:
                m_owner->OnOpenTerminal(location);
                break;
            case kMenuOpenVSCode:
                m_owner->OnOpenVSCode(location);
                break;
            case kMenuCopyPath:
                m_owner->OnCopyPath(location);
                break;
            default:
                break;
        }
    } else if (item && item->type == TabViewItemType::kGroupHeader) {
        const int groupIndex = item->location.groupIndex;
        switch (command) {
            case kMenuToggleGroup:
                m_owner->OnToggleGroupCollapsed(groupIndex);
                break;
            case kMenuDetachGroup:
                m_owner->OnDetachGroupRequested(groupIndex);
                break;
            case kMenuCreateGroupAfter:
                m_owner->OnCreateIslandAfter(groupIndex);
                break;
            case kMenuUnhideAll:
                m_owner->OnUnhideAllInGroup(groupIndex);
                break;
            case kMenuToggleSplit:
                m_owner->OnToggleSplitView(groupIndex);
                break;
            default:
                break;
        }
    }
}

void TabBandWindow::HandleMiddleClick(int commandId) {
    if (!m_owner) {
        return;
    }
    const TabLocation location = LocationForCommand(commandId);
    if (location.IsValid()) {
        m_owner->OnCloseTabRequested(location);
    }
}

void TabBandWindow::HandleLButtonDown(int commandId) {
    RelayFocusToToolbar();
    if (commandId == kNewTabCommandId) {
        HandleToolbarCommand(commandId);
        return;
    }

    const TabViewItem* item = ItemForCommand(commandId);
    if (!item) {
        return;
    }

    POINT screenPt{};
    GetCursorPos(&screenPt);

    if (item->type == TabViewItemType::kTab) {
        RECT closeRect{};
        if (IsPointInCloseButton(commandId, screenPt, &closeRect)) {
            m_closeState.tracking = true;
            m_closeState.hot = true;
            m_closeState.commandId = commandId;
            m_closeState.rect = closeRect;
            m_ignoreNextCommand = true;
            m_ignoredCommandId = commandId;
            if (m_toolbar) {
                SetCapture(m_toolbar);
            }
            InvalidateButton(commandId);
            return;
        }
    }

    BeginDrag(commandId, screenPt);
}

void TabBandWindow::HandleFilesDropped(TabLocation location, const std::vector<std::wstring>& paths, bool move) {
    if (!m_owner || !location.IsValid() || paths.empty()) {
        return;
    }
    m_owner->OnFilesDropped(location, paths, move);
}

void TabBandWindow::HandleTooltipRequest(NMTTDISPINFOW* info) {
    if (!info) {
        return;
    }
    const int commandId = static_cast<int>(info->hdr.idFrom);
    const TabViewItem* item = ItemForCommand(commandId);
    if (!item) {
        return;
    }
    const std::wstring& tooltip = item->tooltip.empty() ? item->name : item->tooltip;
    static wchar_t buffer[kMaxTooltip];
    wcsncpy_s(buffer, tooltip.c_str(), _TRUNCATE);
    info->lpszText = buffer;
    info->hinst = nullptr;
}

void TabBandWindow::RelayFocusToToolbar() {
    if (m_toolbar) {
        SetFocus(m_toolbar);
    }
}

void TabBandWindow::HandleMouseMove(const POINT& screenPt) {
    if (m_closeState.tracking) {
        bool wasHot = m_closeState.hot;
        m_closeState.hot = PtInRect(&m_closeState.rect, screenPt) != FALSE;
        if (wasHot != m_closeState.hot) {
            InvalidateButton(m_closeState.commandId);
        }
        return;
    }
    if (!m_dragState.tracking) {
        return;
    }
    if (!m_dragState.dragging) {
        const int thresholdX = std::max(GetSystemMetrics(SM_CXDRAG), 1);
        const int thresholdY = std::max(GetSystemMetrics(SM_CYDRAG), 1);
        const int deltaX = std::abs(screenPt.x - m_dragState.startPoint.x);
        const int deltaY = std::abs(screenPt.y - m_dragState.startPoint.y);
        if (deltaX >= thresholdX || deltaY >= thresholdY) {
            m_dragState.dragging = true;
            m_ignoreNextCommand = true;
            m_ignoredCommandId = m_dragState.commandId;
            if (m_toolbar) {
                SendMessageW(m_toolbar, WM_CANCELMODE, 0, 0);
                SendMessageW(m_toolbar, TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
            }
        }
    }
    if (m_dragState.dragging) {
        UpdateDrag(screenPt);
    }
}

void TabBandWindow::HandleLButtonUp(const POINT& screenPt) {
    if (m_closeState.tracking) {
        const int commandId = m_closeState.commandId;
        const bool inside = PtInRect(&m_closeState.rect, screenPt) != FALSE;
        ResetCloseTracking();
        if (inside) {
            if (m_owner) {
                const TabLocation location = LocationForCommand(commandId);
                if (location.IsValid()) {
                    m_owner->OnCloseTabRequested(location);
                }
            }
            return;
        }
        m_ignoreNextCommand = false;
        m_ignoredCommandId = -1;
        if (commandId != -1) {
            HandleToolbarCommand(commandId);
            UpdateCheckedState();
        }
        return;
    }
    if (!m_dragState.tracking) {
        return;
    }
    EndDrag(screenPt, false);
}

int TabBandWindow::CommandIdFromButtonIndex(int index) const {
    if (!m_toolbar || index < 0) {
        return -1;
    }

    TBBUTTON button{};
    const LRESULT result = SendMessageW(m_toolbar, TB_GETBUTTON, static_cast<WPARAM>(index),
                                        reinterpret_cast<LPARAM>(&button));
    if (result == FALSE) {
        return -1;
    }

    return button.idCommand;
}

void TabBandWindow::BeginDrag(int commandId, const POINT& screenPt) {
    if (!m_toolbar) {
        return;
    }
    const TabViewItem* item = ItemForCommand(commandId);
    if (!item) {
        return;
    }

    DragState state{};
    state.tracking = true;
    state.commandId = commandId;
    state.startPoint = screenPt;

    if (item->type == TabViewItemType::kGroupHeader) {
        state.isGroup = true;
        state.groupIndex = item->location.groupIndex;
    } else if (item->type == TabViewItemType::kTab && item->location.IsValid()) {
        state.isGroup = false;
        state.tabLocation = item->location;
    } else {
        return;
    }

    m_dragState = state;
    SetCapture(m_toolbar);
}

void TabBandWindow::UpdateDrag(const POINT& /*screenPt*/) {
    // No visual feedback for now beyond cancelling the default hot item.
}

void TabBandWindow::EndDrag(const POINT& screenPt, bool canceled) {
    if (!m_toolbar) {
        m_dragState = {};
        m_ignoreNextCommand = false;
        m_ignoredCommandId = -1;
        return;
    }

    if (GetCapture() == m_toolbar) {
        ReleaseCapture();
    }

    DragState state = m_dragState;
    m_dragState = {};

    if (canceled || !state.tracking) {
        m_ignoreNextCommand = false;
        m_ignoredCommandId = -1;
        return;
    }

    if (!state.dragging) {
        m_ignoreNextCommand = false;
        m_ignoredCommandId = -1;
        return;
    }

    POINT clientPt = screenPt;
    ScreenToClient(m_toolbar, &clientPt);

    bool performedAction = false;

    if (state.isGroup) {
        bool detachGroup = false;
        if (m_toolbar) {
            RECT toolbarRect{};
            if (GetWindowRect(m_toolbar, &toolbarRect)) {
                detachGroup = !PtInRect(&toolbarRect, screenPt);
            }
        }

        if (detachGroup) {
            if (m_owner && state.groupIndex >= 0) {
                m_owner->OnDetachGroupRequested(state.groupIndex);
                performedAction = true;
            }
        } else {
            const int insertIndex = ComputeGroupInsertIndex(clientPt);
            if (insertIndex >= 0 && state.groupIndex >= 0 && m_owner) {
                if (insertIndex != state.groupIndex && insertIndex != state.groupIndex + 1) {
                    m_owner->OnMoveGroupRequested(state.groupIndex, insertIndex);
                    performedAction = true;
                }
            }
        }
    } else {
        TabLocation target = ComputeTabInsertLocation(clientPt);
        if (target.IsValid() && state.tabLocation.IsValid() && m_owner) {
            if (target.groupIndex == state.tabLocation.groupIndex &&
                target.tabIndex > state.tabLocation.tabIndex) {
                --target.tabIndex;
            }
            if (target.groupIndex != state.tabLocation.groupIndex ||
                target.tabIndex != state.tabLocation.tabIndex) {
                m_owner->OnMoveTabRequested(state.tabLocation, target);
                performedAction = true;
            }
        }
    }

    if (!performedAction && state.commandId != -1) {
        m_ignoreNextCommand = true;
        m_ignoredCommandId = state.commandId;
        HandleToolbarCommand(state.commandId);
        UpdateCheckedState();
    }

    m_ignoreNextCommand = false;
    m_ignoredCommandId = -1;
}

void TabBandWindow::CancelDrag() {
    if (!m_toolbar) {
        ResetCloseTracking();
        return;
    }
    if (GetCapture() == m_toolbar) {
        ReleaseCapture();
    }
    m_dragState = {};
    m_ignoreNextCommand = false;
    m_ignoredCommandId = -1;
    ResetCloseTracking();
}

TabLocation TabBandWindow::ComputeTabInsertLocation(const POINT& clientPt) const {
    TabLocation invalid{};
    invalid.groupIndex = -1;
    invalid.tabIndex = -1;

    if (!m_toolbar) {
        return invalid;
    }

    const LRESULT buttonCount = SendMessageW(m_toolbar, TB_BUTTONCOUNT, 0, 0);
    if (buttonCount <= 0) {
        return invalid;
    }

    struct ButtonInfo {
        RECT rect{};
        const TabViewItem* item = nullptr;
    };

    std::vector<ButtonInfo> items;
    items.reserve(static_cast<size_t>(buttonCount));
    RECT newTabRect{};
    bool hasNewTab = false;

    for (LRESULT index = 0; index < buttonCount; ++index) {
        RECT rect{};
        if (!SendMessageW(m_toolbar, TB_GETITEMRECT, static_cast<WPARAM>(index),
                          reinterpret_cast<LPARAM>(&rect))) {
            continue;
        }
        const int commandId = CommandIdFromButtonIndex(static_cast<int>(index));
        if (commandId == -1) {
            continue;
        }
        const TabViewItem* item = ItemForCommand(commandId);
        if (item) {
            items.push_back(ButtonInfo{rect, item});
        } else if (commandId == kNewTabCommandId) {
            newTabRect = rect;
            hasNewTab = true;
        }
    }

    if (items.empty()) {
        return invalid;
    }

    auto headerLocation = [](const TabViewItem* header, bool atEnd) {
        TabLocation location{};
        location.groupIndex = header ? header->location.groupIndex : -1;
        if (location.groupIndex < 0) {
            location.tabIndex = -1;
        } else {
            if (atEnd) {
                location.tabIndex = static_cast<int>(header->totalTabs);
            } else {
                location.tabIndex = 0;
            }
        }
        return location;
    };

    auto tabBeforeLocation = [](const TabViewItem* tab) {
        if (!tab) {
            return TabLocation{};
        }
        return tab->location;
    };

    if (clientPt.x < items.front().rect.left) {
        const auto& info = items.front();
        if (info.item->type == TabViewItemType::kGroupHeader) {
            return headerLocation(info.item, false);
        }
        return tabBeforeLocation(info.item);
    }

    for (size_t i = 0; i < items.size(); ++i) {
        const auto& info = items[i];
        if (!info.item) {
            continue;
        }
        if (info.item->type == TabViewItemType::kGroupHeader) {
            const LONG mid = info.rect.left + (info.rect.right - info.rect.left) / 2;
            if (clientPt.x < mid) {
                return headerLocation(info.item, false);
            }
            if (clientPt.x <= info.rect.right) {
                return headerLocation(info.item, true);
            }
            continue;
        }

        const LONG mid = info.rect.left + (info.rect.right - info.rect.left) / 2;
        if (clientPt.x < mid) {
            return tabBeforeLocation(info.item);
        }
    }

    const auto& last = items.back();
    if (last.item && last.item->type == TabViewItemType::kGroupHeader) {
        return headerLocation(last.item, true);
    }
    if (last.item && last.item->type == TabViewItemType::kTab) {
        TabLocation location = last.item->location;
        if (location.IsValid()) {
            ++location.tabIndex;
        }
        return location;
    }

    if (hasNewTab) {
        POINT adjusted = clientPt;
        if (adjusted.x >= newTabRect.left) {
            const auto& info = items.back();
            if (info.item && info.item->type == TabViewItemType::kGroupHeader) {
                return headerLocation(info.item, true);
            }
            if (info.item && info.item->type == TabViewItemType::kTab) {
                TabLocation location = info.item->location;
                if (location.IsValid()) {
                    ++location.tabIndex;
                }
                return location;
            }
        }
    }

    return invalid;
}

int TabBandWindow::ComputeGroupInsertIndex(const POINT& clientPt) const {
    if (!m_toolbar) {
        return -1;
    }

    const LRESULT buttonCount = SendMessageW(m_toolbar, TB_BUTTONCOUNT, 0, 0);
    if (buttonCount <= 0) {
        return -1;
    }

    struct GroupButton {
        RECT rect{};
        int index = -1;
    };

    std::vector<GroupButton> groups;
    groups.reserve(static_cast<size_t>(buttonCount));

    for (LRESULT i = 0; i < buttonCount; ++i) {
        RECT rect{};
        if (!SendMessageW(m_toolbar, TB_GETITEMRECT, static_cast<WPARAM>(i),
                          reinterpret_cast<LPARAM>(&rect))) {
            continue;
        }
        const int commandId = CommandIdFromButtonIndex(static_cast<int>(i));
        if (commandId == -1) {
            continue;
        }
        const TabViewItem* item = ItemForCommand(commandId);
        if (item && item->type == TabViewItemType::kGroupHeader && item->location.groupIndex >= 0) {
            groups.push_back(GroupButton{rect, item->location.groupIndex});
        }
    }

    if (groups.empty()) {
        return -1;
    }

    if (clientPt.x < groups.front().rect.left) {
        return groups.front().index;
    }

    for (const auto& group : groups) {
        const LONG mid = group.rect.left + (group.rect.right - group.rect.left) / 2;
        if (clientPt.x < mid) {
            return group.index;
        }
        if (clientPt.x <= group.rect.right) {
            return group.index + 1;
        }
    }

    const auto& last = groups.back();
    return last.index + 1;
}

const TabViewItem* TabBandWindow::ItemFromPoint(const POINT& screenPt) const {
    if (!m_toolbar) {
        return nullptr;
    }
    POINT clientPt = screenPt;
    ScreenToClient(m_toolbar, &clientPt);
    const int hit = ToolbarHitTest(m_toolbar, clientPt);
    if (hit < 0) {
        return nullptr;
    }
    const int commandId = CommandIdFromButtonIndex(hit);
    if (commandId == -1) {
        return nullptr;
    }
    return ItemForCommand(commandId);
}

TabLocation TabBandWindow::TabLocationFromPoint(const POINT& screenPt) const {
    TabLocation location{};
    location.groupIndex = -1;
    location.tabIndex = -1;
    const TabViewItem* item = ItemFromPoint(screenPt);
    if (!item || item->type != TabViewItemType::kTab) {
        return location;
    }
    if (!item->location.IsValid()) {
        return location;
    }
    return item->location;
}

UINT TabBandWindow::CurrentDpi() const {
    UINT dpi = 96;
    HWND reference = m_toolbar ? m_toolbar : m_hwnd;
    if (reference) {
        HDC dc = GetDC(reference);
        if (dc) {
            dpi = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
            ReleaseDC(reference, dc);
        }
    } else {
        HDC dc = GetDC(nullptr);
        if (dc) {
            dpi = static_cast<UINT>(GetDeviceCaps(dc, LOGPIXELSX));
            ReleaseDC(nullptr, dc);
        }
    }
    return dpi == 0 ? 96u : dpi;
}

int TabBandWindow::GroupIndicatorWidth() const {
    const UINT dpi = CurrentDpi();
    return std::max(5, MulDiv(5, static_cast<int>(dpi), 96));
}

int TabBandWindow::GroupIndicatorHitWidth(bool collapsed) const {
    const UINT dpi = CurrentDpi();
    const int base = std::max(5, MulDiv(5, static_cast<int>(dpi), 96));
    if (!collapsed) {
        return base;
    }
    const int minHit = MulDiv(12, static_cast<int>(dpi), 96);
    return std::max(base, minHit);
}

COLORREF TabBandWindow::GroupIndicatorColor(const TabViewItem& item) const {
    if (item.hasCustomOutline) {
        return item.outlineColor;
    }
    if (item.hasTagColor) {
        return item.tagColor;
    }
    return m_theme.highlight;
}

int TabBandWindow::TabHorizontalPadding() const {
    const UINT dpi = CurrentDpi();
    return std::max(4, MulDiv(4, static_cast<int>(dpi), 96));
}

int TabBandWindow::IconTextSpacing() const {
    const UINT dpi = CurrentDpi();
    return std::max(2, MulDiv(4, static_cast<int>(dpi), 96));
}

int TabBandWindow::CloseButtonSpacing() const {
    const UINT dpi = CurrentDpi();
    return std::max(4, MulDiv(6, static_cast<int>(dpi), 96));
}

int TabBandWindow::CloseButtonSize() const {
    const UINT dpi = CurrentDpi();
    return std::max(10, MulDiv(12, static_cast<int>(dpi), 96));
}

RECT TabBandWindow::CloseButtonRect(const RECT& buttonRect) const {
    RECT closeRect = buttonRect;
    const int padding = TabHorizontalPadding();
    const int size = CloseButtonSize();
    closeRect.right = std::max(closeRect.left, closeRect.right - padding);
    closeRect.left = std::max(closeRect.left, closeRect.right - size);
    const int height = closeRect.bottom - closeRect.top;
    const int verticalOffset = std::max(0, (height - size) / 2);
    closeRect.top += verticalOffset;
    closeRect.bottom = closeRect.top + size;
    return closeRect;
}

bool TabBandWindow::GetButtonRect(int commandId, RECT* rect) const {
    if (!m_toolbar || commandId == -1 || !rect) {
        return false;
    }
    LRESULT index = SendMessageW(m_toolbar, TB_COMMANDTOINDEX, commandId, 0);
    if (index < 0) {
        return false;
    }
    RECT buttonRect{};
    if (!SendMessageW(m_toolbar, TB_GETITEMRECT, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&buttonRect))) {
        return false;
    }
    *rect = buttonRect;
    return true;
}

void TabBandWindow::InvalidateButton(int commandId) const {
    if (!m_toolbar || commandId == -1) {
        return;
    }
    RECT rect{};
    if (GetButtonRect(commandId, &rect)) {
        InvalidateRect(m_toolbar, &rect, TRUE);
    }
}

bool TabBandWindow::IsPointInCloseButton(int commandId, const POINT& screenPt, RECT* closeRectOut) const {
    if (!m_toolbar) {
        return false;
    }
    RECT buttonRect{};
    if (!GetButtonRect(commandId, &buttonRect)) {
        return false;
    }
    RECT closeRect = CloseButtonRect(buttonRect);
    RECT screenRect = closeRect;
    MapWindowPoints(m_toolbar, nullptr, reinterpret_cast<LPPOINT>(&screenRect), 2);
    if (PtInRect(&screenRect, screenPt)) {
        if (closeRectOut) {
            *closeRectOut = screenRect;
        }
        return true;
    }
    return false;
}

void TabBandWindow::ResetCloseTracking() {
    if (!m_closeState.tracking && !m_closeState.hot) {
        m_closeState = {};
        return;
    }
    if (m_closeState.tracking && m_toolbar && GetCapture() == m_toolbar) {
        ReleaseCapture();
    }
    const int commandId = m_closeState.commandId;
    m_closeState = {};
    InvalidateButton(commandId);
}

int TabBandWindow::CalculateTabButtonWidth(const TabViewItem& item) const {
    const UINT dpi = CurrentDpi();
    const int padding = TabHorizontalPadding();
    const int iconSpacing = IconTextSpacing();
    const int closeSpacing = CloseButtonSpacing();
    const int closeSize = CloseButtonSize();

    int textWidth = 0;
    if (m_toolbar && !item.name.empty()) {
        HDC dc = GetDC(m_toolbar);
        if (dc) {
            HFONT font = reinterpret_cast<HFONT>(SendMessageW(m_toolbar, WM_GETFONT, 0, 0));
            HFONT oldFont = nullptr;
            if (font) {
                oldFont = static_cast<HFONT>(SelectObject(dc, font));
            }
            SIZE extent{};
            if (GetTextExtentPoint32W(dc, item.name.c_str(), static_cast<int>(item.name.size()), &extent)) {
                textWidth = extent.cx;
            }
            if (oldFont) {
                SelectObject(dc, oldFont);
            }
            ReleaseDC(m_toolbar, dc);
        }
    }
    if (textWidth <= 0) {
        textWidth = MulDiv(40, static_cast<int>(dpi), 96);
    }

    int width = padding;  // left padding
    if (item.pidl) {
        width += ToolbarIconSize();
        width += iconSpacing;
    }
    width += textWidth;
    width += closeSpacing + closeSize;
    width += padding;  // right padding

    const int minWidth = MulDiv(80, static_cast<int>(dpi), 96);
    return std::max(width, minWidth);
}

void TabBandWindow::RegisterDropTarget() {
    if (m_dropTarget || !m_toolbar) {
        return;
    }
    auto* target = new (std::nothrow) TabToolbarDropTarget(this);
    if (!target) {
        return;
    }
    const HRESULT hr = RegisterDragDrop(m_toolbar, target);
    if (SUCCEEDED(hr)) {
        m_dropTarget = target;
    } else {
        target->Release();
    }
}

void TabBandWindow::RevokeDropTarget() {
    if (!m_dropTarget || !m_toolbar) {
        return;
    }
    RevokeDragDrop(m_toolbar);
    m_dropTarget->Release();
    m_dropTarget = nullptr;
}

LRESULT CALLBACK TabBandWindow::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                        DWORD_PTR refData) {
    (void)id;
    auto* self = reinterpret_cast<TabBandWindow*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_CREATE:
            return 0;
        case WM_SIZE: {
            if (self->m_toolbar) {
                MoveWindow(self->m_toolbar, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
            }
            return 0;
        }
        case WM_THEMECHANGED:
        case WM_SYSCOLORCHANGE:
            self->UpdateTheme();
            break;
        case WM_SETTINGCHANGE: {
            if (self->ShouldUpdateThemeForSettingChange(lParam)) {
                self->UpdateTheme();
            }
            break;
        }
        case WM_ERASEBKGND: {
            if (self->PaintHostBackground(reinterpret_cast<HDC>(wParam))) {
                return 1;
            }
            break;
        }
        case WM_COMMAND: {
            if (HIWORD(wParam) == 0) {
                if (self->m_ignoreNextCommand && LOWORD(wParam) == self->m_ignoredCommandId) {
                    self->m_ignoreNextCommand = false;
                    self->m_ignoredCommandId = -1;
                    return 0;
                }
                self->m_ignoreNextCommand = false;
                self->m_ignoredCommandId = -1;
                self->HandleToolbarCommand(static_cast<int>(LOWORD(wParam)));
                self->UpdateCheckedState();
                return 0;
            }
            break;
        }
        case WM_SHELLTABS_DEFER_NAVIGATE: {
            if (self->m_owner) {
                self->m_owner->OnDeferredNavigate();
            }
            return 0;
        }
        case WM_SHELLTABS_REFRESH_COLORIZER: {
            if (self->m_owner) {
                self->m_owner->OnColorizerRefresh();
            }
            return 0;
        }
        case WM_SHELLTABS_ENABLE_GIT_STATUS: {
            if (self->m_owner) {
                self->m_owner->OnEnableGitStatus();
            }
            return 0;
        }
        case WM_SHELLTABS_REFRESH_GIT_STATUS: {
            if (self->m_owner) {
                self->m_owner->OnGitStatusUpdated();
            }
            return 0;
        }
        case WM_NOTIFY: {
            const NMHDR* header = reinterpret_cast<const NMHDR*>(lParam);
            if (!header) {
                break;
            }
            if (header->code == TTN_GETDISPINFOW) {
                auto* tip = reinterpret_cast<NMTTDISPINFOW*>(lParam);
                self->HandleTooltipRequest(tip);
                return TRUE;
            }
            if (header->hwndFrom == self->m_toolbar) {
                switch (header->code) {
                    case NM_CUSTOMDRAW: {
                        auto* customDraw = reinterpret_cast<NMTBCUSTOMDRAW*>(lParam);
                        return self->HandleToolbarCustomDraw(customDraw);
                    }
                    case NM_RCLICK: {
                        const auto* mouse = reinterpret_cast<const NMMOUSE*>(header);
                        POINT pt = mouse->pt;
                        ClientToScreen(self->m_toolbar, &pt);
                        int commandId = static_cast<int>(mouse->dwItemSpec);
                        if (commandId == -1) {
                            commandId = kNewTabCommandId;
                        }
                        self->HandleContextMenu(commandId, pt);
                        return TRUE;
                    }
                    case NM_CLICK: {
                        const auto* mouse = reinterpret_cast<const NMMOUSE*>(header);
                        if (mouse->dwItemSpec != static_cast<DWORD_PTR>(-1) && mouse->dwItemSpec == kNewTabCommandId) {
                            self->HandleToolbarCommand(kNewTabCommandId);
                            self->UpdateCheckedState();
                        }
                        return FALSE;
                    }
                    default:
                        break;
                }
            }
            break;
        }
        case WM_SETFOCUS: {
            self->RelayFocusToToolbar();
            return 0;
        }
        case WM_DESTROY:
            self->DestroyToolbar();
            return 0;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK TabBandWindow::ToolbarWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                               DWORD_PTR refData) {
    auto* self = reinterpret_cast<TabBandWindow*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    UNREFERENCED_PARAMETER(id);

    switch (msg) {
        case WM_LBUTTONDOWN: {
            const POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            const int hit = ToolbarHitTest(hwnd, pt);
            if (hit >= 0) {
                const int commandId = self->CommandIdFromButtonIndex(hit);
                if (commandId != -1) {
                    self->HandleLButtonDown(commandId);
                }
            }
            break;
        }
        case WM_LBUTTONUP: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            ClientToScreen(hwnd, &pt);
            self->HandleLButtonUp(pt);
            break;
        }
        case WM_MOUSEMOVE: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            ClientToScreen(hwnd, &pt);
            self->HandleMouseMove(pt);
            break;
        }
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
            self->CancelDrag();
            break;
        case WM_MBUTTONUP: {
            const POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            const int hit = ToolbarHitTest(hwnd, pt);
            if (hit >= 0) {
                const int commandId = self->CommandIdFromButtonIndex(hit);
                if (commandId != -1) {
                    self->HandleMiddleClick(commandId);
                }
            }
            return 0;
        }
        case WM_CONTEXTMENU: {
            POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            if (pt.x == -1 && pt.y == -1) {
                pt = {0, 0};
                ClientToScreen(hwnd, &pt);
            }
            POINT clientPt = pt;
            ScreenToClient(hwnd, &clientPt);
            const int hit = ToolbarHitTest(hwnd, clientPt);
            int commandId = kNewTabCommandId;
            if (hit >= 0) {
                const int hitCommandId = self->CommandIdFromButtonIndex(hit);
                if (hitCommandId != -1) {
                    commandId = hitCommandId;
                }
            }
            self->HandleContextMenu(commandId, pt);
            return 0;
        }
        case WM_THEMECHANGED:
        case WM_SYSCOLORCHANGE:
            self->UpdateTheme();
            break;
        case WM_SETTINGCHANGE:
            if (self->ShouldUpdateThemeForSettingChange(lParam)) {
                self->UpdateTheme();
            }
            break;
        case WM_ERASEBKGND: {
            if (self->PaintToolbarBackground(hwnd, reinterpret_cast<HDC>(wParam))) {
                return 1;
            }
            break;
        }
        case WM_PRINTCLIENT: {
            if (self->PaintToolbarBackground(hwnd, reinterpret_cast<HDC>(wParam))) {
                break;
            }
            break;
        }
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT TabBandWindow::HandleToolbarCustomDraw(NMTBCUSTOMDRAW* customDraw) {
    if (!customDraw) {
        return CDRF_DODEFAULT;
    }

    switch (customDraw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT: {
            customDraw->clrBtnFace = m_theme.background;
            customDraw->clrBtnHighlight = m_theme.hover;
            customDraw->clrHighlightHotTrack = m_theme.hover;
            customDraw->clrText = m_theme.text;
            return TBCDRF_USECDCOLORS | CDRF_NOTIFYITEMDRAW;
        }
        case CDDS_ITEMPREPAINT: {
            const int commandId = static_cast<int>(customDraw->nmcd.dwItemSpec);
            const TabViewItem* item = ItemForCommand(commandId);
            const bool isGroupHeader = item && item->type == TabViewItemType::kGroupHeader;
            const bool isTab = item && item->type == TabViewItemType::kTab;
            const bool isChecked = (customDraw->nmcd.uItemState & CDIS_CHECKED) != 0;
            const bool isHot = (customDraw->nmcd.uItemState & CDIS_HOT) != 0;
            const bool isPressed = (customDraw->nmcd.uItemState & CDIS_SELECTED) != 0;
            const bool isDisabled = (customDraw->nmcd.uItemState & CDIS_DISABLED) != 0;

            if (isGroupHeader && item) {
                RECT rect = customDraw->nmcd.rc;
                FillRectColor(customDraw->nmcd.hdc, rect, m_theme.background);

                const int indicatorWidth = GroupIndicatorWidth();
                RECT indicatorRect = rect;
                indicatorRect.right = std::min(indicatorRect.left + indicatorWidth, rect.right);
                COLORREF indicatorColor = AdjustIndicatorColorForState(GroupIndicatorColor(*item), isHot, isPressed);
                FillRectColor(customDraw->nmcd.hdc, indicatorRect, indicatorColor);

                return CDRF_SKIPDEFAULT;
            }

            if (isTab && item) {
                RECT rect = customDraw->nmcd.rc;
                COLORREF fill = m_theme.background;
                if (isDisabled) {
                    fill = DarkenColor(m_theme.background, 0.05);
                }
                if (isChecked) {
                    fill = m_theme.checked;
                } else if (isPressed) {
                    fill = m_theme.pressed;
                } else if (isHot) {
                    fill = m_theme.hover;
                }

                FillRectColor(customDraw->nmcd.hdc, rect, fill);

                if (isChecked) {
                    RECT borderRect = rect;
                    InflateRect(&borderRect, -1, -1);
                    FrameRectColor(customDraw->nmcd.hdc, borderRect, m_theme.highlight);
                }

                RECT closeRect = CloseButtonRect(rect);
                const UINT dpi = CurrentDpi();
                const int padding = TabHorizontalPadding();
                const int iconSpacing = IconTextSpacing();
                const int closeSpacing = CloseButtonSpacing();
                const int iconSize = ToolbarIconSize();
                int contentLeft = rect.left + padding;

                if (customDraw->iImage >= 0 && m_imageList) {
                    const int iconTop = rect.top + std::max(0, (rect.bottom - rect.top - iconSize) / 2);
                    ImageList_DrawEx(m_imageList, customDraw->iImage, customDraw->nmcd.hdc, contentLeft, iconTop, iconSize,
                                     iconSize, CLR_NONE, CLR_NONE, ILD_NORMAL);
                    contentLeft += iconSize + iconSpacing;
                }

                RECT textRect = rect;
                textRect.left = contentLeft;
                textRect.right = std::max(textRect.left, closeRect.left - closeSpacing);

                HFONT font = reinterpret_cast<HFONT>(SendMessageW(m_toolbar, WM_GETFONT, 0, 0));
                HFONT oldFont = nullptr;
                if (font) {
                    oldFont = static_cast<HFONT>(SelectObject(customDraw->nmcd.hdc, font));
                }

                SetBkMode(customDraw->nmcd.hdc, TRANSPARENT);
                COLORREF textColor = isDisabled ? m_theme.textDisabled : m_theme.text;
                SetTextColor(customDraw->nmcd.hdc, textColor);
                const int textLength = static_cast<int>(item->name.size());
                if (textLength > 0) {
                    DrawTextW(customDraw->nmcd.hdc, item->name.c_str(), textLength, &textRect,
                              DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
                }

                if (oldFont) {
                    SelectObject(customDraw->nmcd.hdc, oldFont);
                }

                COLORREF closeColor = m_darkModeEnabled ? RGB(232, 72, 72) : RGB(200, 56, 56);
                COLORREF closeBorder = DarkenColor(closeColor, 0.35);
                COLORREF glyphColor = RGB(255, 255, 255);
                const bool closeTracking = m_closeState.commandId == commandId && m_closeState.tracking;
                const bool closeHot = closeTracking && m_closeState.hot;
                if (isDisabled) {
                    closeColor = DarkenColor(m_theme.background, 0.2);
                    closeBorder = DarkenColor(closeColor, 0.25);
                    glyphColor = m_theme.textDisabled;
                } else if (closeTracking) {
                    if (closeHot) {
                        closeColor = DarkenColor(closeColor, 0.2);
                    } else {
                        closeColor = LightenColor(closeColor, 0.15);
                    }
                    closeBorder = DarkenColor(closeColor, 0.3);
                }

                FillRectColor(customDraw->nmcd.hdc, closeRect, closeColor);
                FrameRectColor(customDraw->nmcd.hdc, closeRect, closeBorder);

                const int penWidth = std::max(1, MulDiv(1, static_cast<int>(dpi), 96));
                HPEN pen = CreatePen(PS_SOLID, penWidth, glyphColor);
                if (pen) {
                    HGDIOBJ oldPen = SelectObject(customDraw->nmcd.hdc, pen);
                    const int inset = std::max(2, CloseButtonSize() / 4);
                    const int left = closeRect.left + inset;
                    const int right = closeRect.right - inset;
                    const int top = closeRect.top + inset;
                    const int bottom = closeRect.bottom - inset;
                    MoveToEx(customDraw->nmcd.hdc, left, top, nullptr);
                    LineTo(customDraw->nmcd.hdc, right, bottom);
                    MoveToEx(customDraw->nmcd.hdc, left, bottom, nullptr);
                    LineTo(customDraw->nmcd.hdc, right, top);
                    if (oldPen) {
                        SelectObject(customDraw->nmcd.hdc, oldPen);
                    }
                    DeleteObject(pen);
                }

                return CDRF_SKIPDEFAULT;
            }

            COLORREF fill = m_theme.background;
            COLORREF text = m_theme.text;
            if (isGroupHeader) {
                fill = m_theme.groupHeaderBackground;
                text = m_theme.groupHeaderText;
            }
            if (isDisabled) {
                text = m_theme.textDisabled;
            }
            if (isChecked) {
                fill = m_theme.checked;
            } else if (isPressed) {
                fill = m_theme.pressed;
            } else if (isHot) {
                fill = isGroupHeader ? m_theme.groupHeaderHover : m_theme.hover;
            }

            RECT rect = customDraw->nmcd.rc;
            FillRectColor(customDraw->nmcd.hdc, rect, fill);

            if (isChecked) {
                RECT borderRect = rect;
                InflateRect(&borderRect, -1, -1);
                FrameRectColor(customDraw->nmcd.hdc, borderRect, m_theme.highlight);
            }

            customDraw->clrText = text;
            customDraw->clrBtnFace = fill;
            customDraw->clrBtnHighlight = fill;
            customDraw->clrHighlightHotTrack = isGroupHeader ? m_theme.groupHeaderHover : m_theme.hover;
            return TBCDRF_USECDCOLORS;
        }
        default:
            break;
    }

    return CDRF_DODEFAULT;
}

void TabBandWindow::UpdateTheme() {
    EnsurePreferredAppMode();
    auto& themeApi = GetThemeApi();
    if (themeApi.refreshImmersiveColorPolicyState) {
        themeApi.refreshImmersiveColorPolicyState();
    }
    const bool darkPreferred = IsDarkModePreferred();
    const ToolbarTheme newTheme = CalculateTheme(darkPreferred);
    if (newTheme == m_theme && darkPreferred == m_darkModeEnabled) {
        return;
    }

    m_theme = newTheme;
    m_darkModeEnabled = darkPreferred;

    ApplyThemeToToolbar();

    if (m_hwnd) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
    if (m_toolbar) {
        InvalidateRect(m_toolbar, nullptr, TRUE);
    }
}

void TabBandWindow::ApplyThemeToToolbar() {
    const wchar_t* themeName = m_darkModeEnabled ? L"DarkMode_Explorer" : L"Explorer";
    auto& themeApi = GetThemeApi();
    if (m_hwnd) {
        SetWindowTheme(m_hwnd, themeName, nullptr);
        if (themeApi.allowDarkModeForWindow) {
            themeApi.allowDarkModeForWindow(m_hwnd, m_darkModeEnabled ? TRUE : FALSE);
        }
        const BOOL useDark = m_darkModeEnabled ? TRUE : FALSE;
        DwmSetWindowAttribute(m_hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDark, sizeof(useDark));
    }
    if (m_toolbar) {
        SetWindowTheme(m_toolbar, themeName, nullptr);
        if (themeApi.allowDarkModeForWindow) {
            themeApi.allowDarkModeForWindow(m_toolbar, m_darkModeEnabled ? TRUE : FALSE);
        }
        const BOOL useDark = m_darkModeEnabled ? TRUE : FALSE;
        DwmSetWindowAttribute(m_toolbar, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDark, sizeof(useDark));
    }
}

bool TabBandWindow::PaintHostBackground(HDC dc) const {
    if (!dc || !m_hwnd) {
        return false;
    }
    RECT rect{};
    if (!GetClientRect(m_hwnd, &rect)) {
        return false;
    }
    if (m_darkModeEnabled) {
        FillRectColor(dc, rect, m_theme.background);
        if (m_theme.separator != CLR_INVALID) {
            RECT border = rect;
            border.top = border.bottom - 1;
            if (border.top < border.bottom) {
                FillRectColor(dc, border, m_theme.separator);
            }
        }
        return true;
    }
    if (IsThemeActive()) {
        HTHEME theme = OpenThemeData(m_hwnd, L"Rebar");
        if (theme) {
            if (IsThemeBackgroundPartiallyTransparent(theme, RP_BAND, 0)) {
                DrawThemeParentBackground(m_hwnd, dc, &rect);
            }
            DrawThemeBackground(theme, dc, RP_BAND, 0, &rect, nullptr);
            CloseThemeData(theme);
            if (m_theme.separator != CLR_INVALID) {
                RECT border = rect;
                border.top = border.bottom - 1;
                if (border.top < border.bottom) {
                    FillRectColor(dc, border, m_theme.separator);
                }
            }
            return true;
        }
    }
    FillRectColor(dc, rect, m_theme.background);
    return true;
}

bool TabBandWindow::PaintToolbarBackground(HWND hwnd, HDC dc) const {
    if (!dc || !hwnd) {
        return false;
    }
    RECT rect{};
    if (!GetClientRect(hwnd, &rect)) {
        return false;
    }
    if (m_darkModeEnabled) {
        FillRectColor(dc, rect, m_theme.background);
        return true;
    }
    if (IsThemeActive()) {
        DrawThemeParentBackground(hwnd, dc, &rect);
        return true;
    }
    FillRectColor(dc, rect, m_theme.background);
    return true;
}

bool TabBandWindow::ShouldUpdateThemeForSettingChange(LPARAM lParam) const {
    if (!lParam) {
        return true;
    }
    std::wstring_view setting(reinterpret_cast<PCWSTR>(lParam));
    if (setting.empty()) {
        return true;
    }
    static constexpr std::wstring_view kTargets[] = {L"ImmersiveColorSet", L"ImmersiveColorSetChanged", L"WindowsThemeElement",
                                                      L"SystemUsesLightTheme", L"AppsUseLightTheme"};
    for (const auto& target : kTargets) {
        if (setting == target) {
            return true;
        }
    }
    return false;
}

bool TabBandWindow::IsDarkModePreferred() const {
    auto& themeApi = GetThemeApi();
    if (themeApi.shouldAppsUseDarkMode) {
        return themeApi.shouldAppsUseDarkMode() != FALSE;
    }
    DWORD value = 1;
    DWORD size = sizeof(value);
    if (RegGetValueW(HKEY_CURRENT_USER,
                     L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                     L"AppsUseLightTheme", RRF_RT_REG_DWORD, nullptr, &value, &size) == ERROR_SUCCESS) {
        return value == 0;
    }
    return false;
}

TabBandWindow::ToolbarTheme TabBandWindow::CalculateTheme(bool darkMode) const {
    ToolbarTheme theme{};
    theme.border = CLR_INVALID;
    theme.separator = CLR_INVALID;
    theme.highlight = GetSysColor(COLOR_HIGHLIGHT);
    if (darkMode) {
        theme.background = RGB(32, 32, 32);
        theme.hover = RGB(52, 52, 52);
        theme.pressed = RGB(62, 62, 62);
        theme.checked = RGB(72, 72, 72);
        theme.text = RGB(241, 241, 241);
        theme.textDisabled = RGB(150, 150, 150);
        theme.groupHeaderBackground = RGB(26, 26, 26);
        theme.groupHeaderHover = RGB(46, 46, 46);
        theme.groupHeaderText = RGB(189, 189, 189);
        theme.border = RGB(63, 63, 63);
        theme.separator = RGB(58, 58, 58);
    } else {
        theme.border = GetSysColor(COLOR_3DSHADOW);
        theme.separator = GetSysColor(COLOR_3DSHADOW);
    }

    if (!darkMode && IsThemeActive()) {
        if (m_toolbar) {
            HTHEME toolbarTheme = OpenThemeData(m_toolbar, L"Toolbar");
            if (toolbarTheme) {
                COLORREF color = 0;
                if (SUCCEEDED(GetThemeColor(toolbarTheme, TP_BUTTON, TS_NORMAL, TMT_FILLCOLOR, &color))) {
                    theme.background = color;
                    theme.groupHeaderBackground = color;
                }
                if (SUCCEEDED(GetThemeColor(toolbarTheme, TP_BUTTON, TS_HOT, TMT_FILLCOLOR, &color))) {
                    theme.hover = color;
                    theme.groupHeaderHover = color;
                }
                if (SUCCEEDED(GetThemeColor(toolbarTheme, TP_BUTTON, TS_PRESSED, TMT_FILLCOLOR, &color))) {
                    theme.pressed = color;
                }
                if (SUCCEEDED(GetThemeColor(toolbarTheme, TP_BUTTON, TS_CHECKED, TMT_FILLCOLOR, &color))) {
                    theme.checked = color;
                }
                if (SUCCEEDED(GetThemeColor(toolbarTheme, TP_BUTTON, TS_NORMAL, TMT_TEXTCOLOR, &color))) {
                    theme.text = color;
                    theme.groupHeaderText = color;
                }
                if (SUCCEEDED(GetThemeColor(toolbarTheme, TP_BUTTON, TS_DISABLED, TMT_TEXTCOLOR, &color))) {
                    theme.textDisabled = color;
                }
                CloseThemeData(toolbarTheme);
            }
        }

        if (m_hwnd) {
            HTHEME rebarTheme = OpenThemeData(m_hwnd, L"Rebar");
            if (rebarTheme) {
                COLORREF color = 0;
                if (SUCCEEDED(GetThemeColor(rebarTheme, RP_BAND, 0, TMT_FILLCOLOR, &color))) {
                    theme.groupHeaderBackground = color;
                    theme.background = color;
                }
                if (SUCCEEDED(GetThemeColor(rebarTheme, RP_BAND, 0, TMT_TEXTCOLOR, &color))) {
                    theme.groupHeaderText = color;
                }
                if (SUCCEEDED(GetThemeColor(rebarTheme, RP_BAND, 0, TMT_BORDERCOLOR, &color))) {
                    theme.border = color;
                    theme.separator = color;
                }
                CloseThemeData(rebarTheme);
            }
        }
    }

    theme.groupHeaderHover = theme.groupHeaderHover == RGB(225, 225, 225) ? theme.hover : theme.groupHeaderHover;
    if (theme.separator == CLR_INVALID) {
        theme.separator = theme.border;
    }

    return theme;
}

void TabBandWindow::FillRectColor(HDC dc, const RECT& rect, COLORREF color) {
    if (!dc) {
        return;
    }
    HBRUSH brush = static_cast<HBRUSH>(GetStockObject(DC_BRUSH));
    if (!brush) {
        return;
    }
    const COLORREF previous = SetDCBrushColor(dc, color);
    FillRect(dc, &rect, brush);
    SetDCBrushColor(dc, previous);
}

void TabBandWindow::FrameRectColor(HDC dc, const RECT& rect, COLORREF color) {
    if (!dc) {
        return;
    }
    HBRUSH brush = static_cast<HBRUSH>(GetStockObject(DC_BRUSH));
    if (!brush) {
        return;
    }
    const COLORREF previous = SetDCBrushColor(dc, color);
    FrameRect(dc, &rect, brush);
    SetDCBrushColor(dc, previous);
}

}  // namespace shelltabs
