#include "TabBandWindow.h"

#include <algorithm>
#include <cwchar>
#include <memory>
#include <string_view>

#include <CommCtrl.h>
#include <ShlObj.h>
#include <shellapi.h>
#include <windowsx.h>
#include <winreg.h>

#include "Module.h"
#include "TabBand.h"

namespace shelltabs {

namespace {
const wchar_t kWindowClassName[] = L"ShellTabsNativeToolbarHost";
constexpr int kNewTabCommandId = 40000;
constexpr int kMaxTooltip = 512;

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

}  // namespace

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
    UpdateTheme();
    RebuildToolbar();
}

void TabBandWindow::DestroyToolbar() {
    if (!m_toolbar) {
        return;
    }

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

    SendMessageW(m_toolbar, WM_SETREDRAW, FALSE, 0);
    ClearToolbar();
    ClearImageList();

    std::vector<TBBUTTON> buttons;
    buttons.reserve(m_tabData.size() + 1);

    for (size_t index = 0; index < m_tabData.size(); ++index) {
        const auto& item = m_tabData[index];
        if (item.type == TabViewItemType::kGroupHeader) {
            if (!item.headerVisible) {
                continue;
            }
            TBBUTTON button{};
            const int commandId = m_nextCommandId++;
            button.idCommand = commandId;
            button.fsStyle = BTNS_BUTTON | BTNS_AUTOSIZE | BTNS_SHOWTEXT;
            button.fsState = 0;
            button.iBitmap = I_IMAGENONE;
            button.dwData = static_cast<DWORD_PTR>(index);
            button.iString = static_cast<INT_PTR>(SendMessageW(m_toolbar, TB_ADDSTRINGW, 0,
                                                               reinterpret_cast<LPARAM>(item.name.c_str())));
            buttons.push_back(button);
            m_commandToIndex[commandId] = index;
        } else {
            TBBUTTON button{};
            const int commandId = m_nextCommandId++;
            button.idCommand = commandId;
            button.fsStyle = BTNS_BUTTON | BTNS_AUTOSIZE | BTNS_SHOWTEXT | BTNS_CHECKGROUP;
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
        }
    }

    if (!buttons.empty()) {
        SendMessageW(m_toolbar, TB_ADDBUTTONS, static_cast<WPARAM>(buttons.size()),
                     reinterpret_cast<LPARAM>(buttons.data()));
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
    if (item && item->type == TabViewItemType::kGroupHeader) {
        if (m_owner) {
            m_owner->OnToggleGroupCollapsed(item->location.groupIndex);
        }
        return;
    }
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

LRESULT CALLBACK TabBandWindow::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                        DWORD_PTR refData) {
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
            customDraw->clrBtnShadow = m_theme.border;
            customDraw->clrHighlightHotTrack = m_theme.hover;
            customDraw->clrText = m_theme.text;
            return CDRF_NOTIFYITEMDRAW;
        }
        case CDDS_ITEMPREPAINT: {
            const int commandId = static_cast<int>(customDraw->nmcd.dwItemSpec);
            const TabViewItem* item = ItemForCommand(commandId);
            const bool isGroupHeader = item && item->type == TabViewItemType::kGroupHeader;
            const bool isChecked = (customDraw->nmcd.uItemState & CDIS_CHECKED) != 0;
            const bool isHot = (customDraw->nmcd.uItemState & CDIS_HOT) != 0;
            const bool isPressed = (customDraw->nmcd.uItemState & CDIS_SELECTED) != 0;

            COLORREF fill = m_theme.background;
            COLORREF text = m_theme.text;
            if (isGroupHeader) {
                fill = m_theme.groupHeaderBackground;
                text = m_theme.groupHeaderText;
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
            customDraw->clrBtnShadow = m_theme.border;
            return CDRF_DODEFAULT;
        }
        default:
            break;
    }

    return CDRF_DODEFAULT;
}

void TabBandWindow::UpdateTheme() {
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
    if (!m_toolbar) {
        return;
    }

    SendMessageW(m_toolbar, TB_SETBKCOLOR, 0, static_cast<LPARAM>(m_theme.background));
    SendMessageW(m_toolbar, TB_SETTEXTCOLOR, 0, static_cast<LPARAM>(m_theme.text));
    SetWindowTheme(m_toolbar, m_darkModeEnabled ? L"DarkMode_Explorer" : L"Explorer", nullptr);
}

bool TabBandWindow::PaintHostBackground(HDC dc) const {
    if (!dc || !m_hwnd) {
        return false;
    }
    RECT rect{};
    if (!GetClientRect(m_hwnd, &rect)) {
        return false;
    }
    FillRectColor(dc, rect, m_theme.background);
    RECT bottom = rect;
    bottom.top = bottom.bottom - 1;
    if (bottom.top < bottom.bottom) {
        FillRectColor(dc, bottom, m_theme.separator);
    }
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
        theme.highlight = RGB(45, 137, 255);
        theme.border = RGB(63, 63, 63);
        theme.separator = RGB(58, 58, 58);
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
