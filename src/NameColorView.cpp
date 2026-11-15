#include "NameColorView.h"

#include <commdlg.h>
#include <shlobj.h>
#include <windowsx.h>
#include <uxtheme.h>

#include <algorithm>
#include <cwchar>

#include "ExplorerThemeUtils.h"
#include "Logging.h"
#include "NameColorStore.h"
#include "ThemeHooks.h"
#include "Utilities.h"

namespace shelltabs {

namespace {

constexpr UINT_PTR kSubclassId = 0x53544E43;  // 'STNC'

// Context menu command IDs. Kept well away from DefView's IContextMenu range and the BHO's
// 0xE170 / 0xE200+ commands.
constexpr UINT kCmdSetColor = 0xE3A0;
constexpr UINT kCmdClearColor = 0xE3A1;
constexpr UINT kCmdSetExtensionColor = 0xE3A2;
constexpr UINT kCmdClearExtensionColor = 0xE3A3;
constexpr UINT kCmdFirst = kCmdSetColor;
constexpr UINT kCmdLast = kCmdClearExtensionColor;

// The view that injected the currently open context menu (menus are modal per thread).
thread_local NameColorView* t_menuOwner = nullptr;

COLORREF g_customColors[16] = {};

bool IsClass(HWND hwnd, const wchar_t* className) {
    wchar_t buffer[64] = {};
    return hwnd && GetClassNameW(hwnd, buffer, ARRAYSIZE(buffer)) > 0 && _wcsicmp(buffer, className) == 0;
}

std::wstring ParsingName(PCIDLIST_ABSOLUTE pidl) {
    PWSTR name = nullptr;
    if (FAILED(SHGetNameFromIDList(pidl, SIGDN_DESKTOPABSOLUTEPARSING, &name)) || !name) {
        return {};
    }
    std::wstring result(name);
    CoTaskMemFree(name);
    return result;
}

bool IsFolderAttributes(SFGAOF attributes) {
    // Archives (.zip, .cab) report SFGAO_FOLDER | SFGAO_STREAM; treat them as files so extension
    // colors apply.
    return (attributes & SFGAO_FOLDER) != 0 && (attributes & SFGAO_STREAM) == 0;
}

// Reads the private per-item PIDL that NamespaceTreeControl stores in TVITEM::lParam
// (a pointer to a pointer to the child ID list, same layout Open-Shell relies on).
PCUIDLIST_RELATIVE ReadTreeItemChildPidl(LPARAM param) {
    PCUIDLIST_RELATIVE result = nullptr;
    __try {
        auto** indirect = reinterpret_cast<PCUIDLIST_RELATIVE**>(param);
        if (indirect && *indirect && **indirect) {
            result = **indirect;
        }
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        result = nullptr;
    }
    return result;
}

bool PickColor(HWND owner, std::optional<COLORREF> initial, COLORREF* result) {
    CHOOSECOLORW choose{};
    choose.lStructSize = sizeof(choose);
    choose.hwndOwner = owner;
    choose.lpCustColors = g_customColors;
    choose.rgbResult = initial.value_or(RGB(255, 128, 0));
    choose.Flags = CC_FULLOPEN | CC_RGBINIT | CC_ANYCOLOR;
    if (!ChooseColorW(&choose)) {
        return false;
    }
    *result = choose.rgbResult;
    return true;
}

}  // namespace

NameColorView::~NameColorView() { Detach(); }

void NameColorView::ApplyFolderViewMode(IShellBrowser* browser, bool classicListView) {
    if (!browser) {
        return;
    }
    Microsoft::WRL::ComPtr<IFolderViewOptions> options;
    if (FAILED(browser->QueryInterface(IID_PPV_ARGS(&options))) || !options) {
        return;
    }
    const HRESULT hr =
        options->SetFolderViewOptions(FVO_VISTALAYOUT, classicListView ? FVO_VISTALAYOUT : FVO_DEFAULT);
    LogMessage(LogLevel::Info, L"NameColorView: SetFolderViewOptions(classic=%d) hr=0x%08X", classicListView ? 1 : 0,
               static_cast<unsigned int>(hr));
}

void NameColorView::Attach(IShellBrowser* browser, HWND frame, bool enabled) {
    m_enabled = enabled;
    m_frame = frame;
    m_browser = browser;
    if (!enabled || !browser) {
        Detach();
        return;
    }

    static bool s_filterInstalled = false;
    if (!s_filterInstalled) {
        SetPopupMenuCommandFilter(&NameColorView::HandlePopupCommand);
        s_filterInstalled = true;
    }

    m_folderView.Reset();
    m_folder.Reset();
    Microsoft::WRL::ComPtr<IShellView> view;
    HWND defView = nullptr;
    if (SUCCEEDED(browser->QueryActiveShellView(&view)) && view) {
        view.As(&m_folderView);
        view->GetWindow(&defView);
    }
    if (m_folderView) {
        m_folderView->GetFolder(IID_PPV_ARGS(&m_folder));
    }

    AttachDefView(defView);
    AttachTree(frame);
}

void NameColorView::Detach() {
    DetachDefView();
    DetachTree();
    m_folderView.Reset();
    m_folder.Reset();
    if (t_menuOwner == this) {
        t_menuOwner = nullptr;
    }
}

void NameColorView::OnColorsChanged() {
    NameColorStore::Instance().Reload();
    m_treeCache.clear();
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
    if (m_tree && IsWindow(m_tree)) {
        InvalidateRect(m_tree, nullptr, FALSE);
    }
}

void NameColorView::AttachDefView(HWND defView) {
    if (defView && defView == m_defView && IsWindow(defView)) {
        // Same view; the list view may have been recreated by a view mode change.
        m_listView = FindWindowExW(defView, nullptr, WC_LISTVIEWW, nullptr);
        ApplyListViewTheme();
        return;
    }
    DetachDefView();
    if (!defView || !IsWindow(defView)) {
        return;
    }
    if (!SetWindowSubclass(defView, &NameColorView::DefViewSubclassProc, kSubclassId,
                           reinterpret_cast<DWORD_PTR>(this))) {
        return;
    }
    m_defView = defView;
    m_listView = FindWindowExW(defView, nullptr, WC_LISTVIEWW, nullptr);
    LogMessage(LogLevel::Info, L"NameColorView: attached defView=%p listView=%p", m_defView, m_listView);
    ApplyListViewTheme();
    if (m_listView) {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
}

void NameColorView::DetachDefView() {
    if (m_defView && IsWindow(m_defView)) {
        RemoveWindowSubclass(m_defView, &NameColorView::DefViewSubclassProc, kSubclassId);
    }
    UnsubclassListView();
    m_defView = nullptr;
    m_listView = nullptr;
}

void NameColorView::AttachTree(HWND frame) {
    HWND tree = nullptr;
    if (frame) {
        HWND parent = FindDescendantWindow(frame, L"NamespaceTreeControl");
        if (parent) {
            tree = FindWindowExW(parent, nullptr, WC_TREEVIEWW, nullptr);
        }
    }
    HWND treeParent = tree ? GetParent(tree) : nullptr;
    if (treeParent && treeParent == m_treeParent && IsWindow(treeParent)) {
        m_tree = tree;
        return;
    }
    DetachTree();
    if (!treeParent) {
        return;
    }
    if (SetWindowSubclass(treeParent, &NameColorView::TreeParentSubclassProc, kSubclassId,
                          reinterpret_cast<DWORD_PTR>(this))) {
        m_treeParent = treeParent;
        m_tree = tree;
        m_treeCache.clear();
        InvalidateRect(m_tree, nullptr, FALSE);
    }
}

void NameColorView::DetachTree() {
    if (m_treeParent && IsWindow(m_treeParent)) {
        RemoveWindowSubclass(m_treeParent, &NameColorView::TreeParentSubclassProc, kSubclassId);
    }
    m_treeParent = nullptr;
    m_tree = nullptr;
    m_treeCache.clear();
}

void NameColorView::ApplyListViewTheme() {
    if (!m_listView || !IsWindow(m_listView)) {
        return;
    }
    // The classic list view does not follow Explorer's dark items view on its own.
    const bool dark = IsAppDarkModePreferred();
    m_darkMode = dark;
    SetWindowTheme(m_listView, dark ? L"DarkMode_ItemsView" : L"ItemsView", nullptr);
    if (HWND header = ListView_GetHeader(m_listView)) {
        SetWindowTheme(header, dark ? L"DarkMode_ItemsView" : L"ItemsView", nullptr);
    }
    const COLORREF background = dark ? RGB(25, 25, 25) : GetSysColor(COLOR_WINDOW);
    const COLORREF text = dark ? RGB(230, 230, 230) : GetSysColor(COLOR_WINDOWTEXT);
    ListView_SetBkColor(m_listView, background);
    ListView_SetTextBkColor(m_listView, CLR_NONE);
    ListView_SetTextColor(m_listView, text);
    SubclassListView();
    if (HWND header = ListView_GetHeader(m_listView)) {
        InvalidateRect(header, nullptr, TRUE);
    }
}

// The header control is a child of the list view, so its NM_CUSTOMDRAW reaches the list view.
// The dark items-view theme paints a dark header background but leaves the text near-black;
// the list view subclass supplies a readable text color in dark mode.
void NameColorView::SubclassListView() {
    if (m_subclassedListView == m_listView) {
        return;
    }
    UnsubclassListView();
    if (m_listView && IsWindow(m_listView) &&
        SetWindowSubclass(m_listView, &NameColorView::ListViewSubclassProc, kSubclassId,
                          reinterpret_cast<DWORD_PTR>(this))) {
        m_subclassedListView = m_listView;
    }
}

void NameColorView::UnsubclassListView() {
    if (m_subclassedListView && IsWindow(m_subclassedListView)) {
        RemoveWindowSubclass(m_subclassedListView, &NameColorView::ListViewSubclassProc, kSubclassId);
    }
    m_subclassedListView = nullptr;
}

LRESULT CALLBACK NameColorView::ListViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                    UINT_PTR id, DWORD_PTR refData) {
    auto* self = reinterpret_cast<NameColorView*>(refData);
    if (msg == WM_NOTIFY && self && self->m_darkMode) {
        auto* header = reinterpret_cast<NMHDR*>(lParam);
        if (header && header->code == NM_CUSTOMDRAW && header->hwndFrom == ListView_GetHeader(hwnd)) {
            auto* draw = reinterpret_cast<NMCUSTOMDRAW*>(lParam);
            if (draw->dwDrawStage == CDDS_PREPAINT) {
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                return result | CDRF_NOTIFYITEMDRAW;
            }
            if (draw->dwDrawStage == CDDS_ITEMPREPAINT) {
                SetTextColor(draw->hdc, RGB(222, 222, 222));
                return DefSubclassProc(hwnd, msg, wParam, lParam) | CDRF_NEWFONT;
            }
        }
    }
    if (msg == WM_NCDESTROY) {
        RemoveWindowSubclass(hwnd, &NameColorView::ListViewSubclassProc, id);
        if (self && self->m_subclassedListView == hwnd) {
            self->m_subclassedListView = nullptr;
        }
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

std::optional<COLORREF> NameColorView::ResolveListItemColor(int index) {
    if (index < 0 || !m_folderView || !m_folder) {
        return std::nullopt;
    }
    PITEMID_CHILD child = nullptr;
    if (FAILED(m_folderView->Item(index, &child)) || !child) {
        return std::nullopt;
    }
    std::optional<COLORREF> color;
    STRRET name{};
    if (SUCCEEDED(m_folder->GetDisplayNameOf(child, SHGDN_FORPARSING, &name))) {
        PWSTR text = nullptr;
        if (SUCCEEDED(StrRetToStrW(&name, child, &text)) && text) {
            SFGAOF attributes = SFGAO_FOLDER | SFGAO_STREAM;
            PCUITEMID_CHILD children[] = {child};
            if (FAILED(m_folder->GetAttributesOf(1, children, &attributes))) {
                attributes = 0;
            }
            color = NameColorStore::Instance().Lookup(text, IsFolderAttributes(attributes));
            CoTaskMemFree(text);
        }
    }
    CoTaskMemFree(child);
    return color;
}

std::optional<COLORREF> NameColorView::ResolveTreeItemColor(HTREEITEM item) {
    auto cached = m_treeCache.find(item);
    if (cached != m_treeCache.end()) {
        return cached->second;
    }

    // Combine the child IDs from this item up to the root into an absolute ID list.
    PIDLIST_ABSOLUTE absolute = nullptr;
    bool valid = true;
    for (HTREEITEM current = item; current && valid; current = TreeView_GetParent(m_tree, current)) {
        TVITEMEXW info{};
        info.mask = TVIF_PARAM;
        info.hItem = current;
        if (!TreeView_GetItem(m_tree, &info)) {
            valid = false;
            break;
        }
        PCUIDLIST_RELATIVE part = ReadTreeItemChildPidl(info.lParam);
        if (!part) {
            valid = false;
            break;
        }
        PIDLIST_ABSOLUTE combined =
            absolute ? ILCombine(reinterpret_cast<PCIDLIST_ABSOLUTE>(part), absolute)
                     : reinterpret_cast<PIDLIST_ABSOLUTE>(ILClone(part));
        if (absolute) {
            ILFree(absolute);
        }
        absolute = combined;
        if (!absolute) {
            valid = false;
        }
    }

    std::optional<COLORREF> color;
    if (valid && absolute) {
        const std::wstring path = ParsingName(absolute);
        if (!path.empty()) {
            color = NameColorStore::Instance().Lookup(path, true);
        }
    }
    if (absolute) {
        ILFree(absolute);
    }
    m_treeCache[item] = color;
    return color;
}

LRESULT NameColorView::HandleListCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT defaultResult) {
    if (NameColorStore::Instance().IsEmpty()) {
        return defaultResult;
    }
    const DWORD stage = draw->nmcd.dwDrawStage;
    if (stage == CDDS_PREPAINT) {
        return defaultResult | CDRF_NOTIFYITEMDRAW;
    }
    const bool itemStage = stage == CDDS_ITEMPREPAINT;
    const bool subItemStage = stage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM);
    if (!itemStage && !(subItemStage && draw->iSubItem == 0)) {
        return defaultResult;
    }
    if ((defaultResult & CDRF_SKIPDEFAULT) != 0 || (draw->nmcd.uItemState & CDIS_SELECTED) != 0) {
        return defaultResult;
    }
    if (auto color = ResolveListItemColor(static_cast<int>(draw->nmcd.dwItemSpec))) {
        draw->clrText = *color;
        return defaultResult | CDRF_NEWFONT;
    }
    return defaultResult;
}

LRESULT NameColorView::HandleTreeCustomDraw(NMTVCUSTOMDRAW* draw, LRESULT defaultResult) {
    if (NameColorStore::Instance().IsEmpty()) {
        return defaultResult;
    }
    if (draw->nmcd.dwDrawStage == CDDS_PREPAINT) {
        return defaultResult | CDRF_NOTIFYITEMDRAW;
    }
    if (draw->nmcd.dwDrawStage != CDDS_ITEMPREPAINT || (defaultResult & CDRF_SKIPDEFAULT) != 0 ||
        (draw->nmcd.uItemState & CDIS_SELECTED) != 0) {
        return defaultResult;
    }
    if (auto color = ResolveTreeItemColor(reinterpret_cast<HTREEITEM>(draw->nmcd.dwItemSpec))) {
        draw->clrText = *color;
        return defaultResult | CDRF_NEWFONT;
    }
    return defaultResult;
}

bool NameColorView::CollectSelection(std::vector<std::wstring>* paths, std::wstring* commonExtension,
                                     bool* anyColored) {
    paths->clear();
    commonExtension->clear();
    *anyColored = false;
    if (!m_folderView) {
        return false;
    }
    Microsoft::WRL::ComPtr<IShellItemArray> items;
    if (FAILED(m_folderView->Items(SVGIO_SELECTION, IID_PPV_ARGS(&items))) || !items) {
        return false;
    }
    DWORD count = 0;
    if (FAILED(items->GetCount(&count)) || count == 0) {
        return false;
    }
    bool extensionConsistent = true;
    for (DWORD i = 0; i < count; ++i) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(i, &item)) || !item) {
            continue;
        }
        PWSTR name = nullptr;
        if (FAILED(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &name)) || !name) {
            continue;
        }
        std::wstring path(name);
        CoTaskMemFree(name);
        SFGAOF attributes = 0;
        item->GetAttributes(SFGAO_FOLDER | SFGAO_STREAM, &attributes);
        const bool isFolder = IsFolderAttributes(attributes);
        if (NameColorStore::Instance().GetPathColor(path)) {
            *anyColored = true;
        }
        if (isFolder) {
            extensionConsistent = false;
        } else if (extensionConsistent) {
            const std::wstring extension = NameColorStore::ExtensionFromPath(path);
            if (extension.empty()) {
                extensionConsistent = false;
            } else if (commonExtension->empty()) {
                *commonExtension = extension;
            } else if (_wcsicmp(commonExtension->c_str(), extension.c_str()) != 0) {
                extensionConsistent = false;
            }
        }
        paths->push_back(std::move(path));
    }
    if (!extensionConsistent || NameColorStore::NormalizeExtensionKey(*commonExtension).empty()) {
        commonExtension->clear();
    }
    return !paths->empty();
}

void NameColorView::InjectContextMenu(HMENU menu) {
    if (!menu || GetMenuState(menu, kCmdSetColor, MF_BYCOMMAND) != static_cast<UINT>(-1)) {
        return;
    }
    bool anyColored = false;
    if (!CollectSelection(&m_menuPaths, &m_menuExtension, &anyColored)) {
        return;
    }
    HMENU submenu = CreatePopupMenu();
    if (!submenu) {
        return;
    }
    AppendMenuW(submenu, MF_STRING, kCmdSetColor, L"Set Color...");
    AppendMenuW(submenu, MF_STRING | (anyColored ? 0 : MF_GRAYED), kCmdClearColor, L"Clear Color");
    if (!m_menuExtension.empty()) {
        const std::wstring setLabel = L"Color All ." + m_menuExtension + L" Files...";
        const std::wstring clearLabel = L"Clear ." + m_menuExtension + L" Color";
        const bool hasExtensionColor = NameColorStore::Instance().GetExtensionColor(m_menuExtension).has_value();
        AppendMenuW(submenu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(submenu, MF_STRING, kCmdSetExtensionColor, setLabel.c_str());
        AppendMenuW(submenu, MF_STRING | (hasExtensionColor ? 0 : MF_GRAYED), kCmdClearExtensionColor,
                    clearLabel.c_str());
    }

    // Insert just above the last separator group ("Properties") when possible.
    int position = GetMenuItemCount(menu);
    for (int i = position - 1; i >= 0; --i) {
        if (GetMenuState(menu, static_cast<UINT>(i), MF_BYPOSITION) & MF_SEPARATOR) {
            position = i;
            break;
        }
    }
    MENUITEMINFOW info{};
    info.cbSize = sizeof(info);
    info.fMask = MIIM_SUBMENU | MIIM_STRING | MIIM_ID;
    info.wID = kCmdSetColor;
    info.hSubMenu = submenu;
    info.dwTypeData = const_cast<wchar_t*>(L"Name Color");
    if (!InsertMenuItemW(menu, static_cast<UINT>(std::max(position, 0)), TRUE, &info)) {
        DestroyMenu(submenu);
        return;
    }
    t_menuOwner = this;
}

void NameColorView::ExecuteCommand(UINT command) {
    HWND owner = m_frame ? m_frame : m_defView;
    auto& store = NameColorStore::Instance();
    switch (command) {
        case kCmdSetColor: {
            std::optional<COLORREF> initial =
                m_menuPaths.empty() ? std::nullopt : store.GetPathColor(m_menuPaths.front());
            COLORREF color = 0;
            if (PickColor(owner, initial, &color)) {
                store.SetPathColors(m_menuPaths, color);
            }
            break;
        }
        case kCmdClearColor:
            store.SetPathColors(m_menuPaths, std::nullopt);
            break;
        case kCmdSetExtensionColor: {
            COLORREF color = 0;
            if (!m_menuExtension.empty() && PickColor(owner, store.GetExtensionColor(m_menuExtension), &color)) {
                store.SetExtensionColor(m_menuExtension, color);
            }
            break;
        }
        case kCmdClearExtensionColor:
            if (!m_menuExtension.empty()) {
                store.SetExtensionColor(m_menuExtension, std::nullopt);
            }
            break;
        default:
            break;
    }
    m_menuPaths.clear();
    m_menuExtension.clear();
    OnColorsChanged();
}

bool NameColorView::HandlePopupCommand(UINT command) {
    if (command < kCmdFirst || command > kCmdLast || !t_menuOwner) {
        return false;
    }
    NameColorView* owner = t_menuOwner;
    t_menuOwner = nullptr;
    owner->ExecuteCommand(command);
    return true;
}

LRESULT CALLBACK NameColorView::DefViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR id, DWORD_PTR refData) {
    auto* self = reinterpret_cast<NameColorView*>(refData);
    switch (msg) {
        case WM_NOTIFY: {
            auto* header = reinterpret_cast<NMHDR*>(lParam);
            if (self && header && header->code == NM_CUSTOMDRAW && IsClass(header->hwndFrom, WC_LISTVIEWW)) {
                self->m_listView = header->hwndFrom;
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                return self->HandleListCustomDraw(reinterpret_cast<NMLVCUSTOMDRAW*>(lParam), result);
            }
            break;
        }
        case WM_CONTEXTMENU: {
            if (!self) {
                break;
            }
            self->m_contextMenuPending = true;
            const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            self->m_contextMenuPending = false;
            return result;
        }
        case WM_INITMENUPOPUP:
            if (self && self->m_contextMenuPending && HIWORD(lParam) == 0) {
                self->m_contextMenuPending = false;  // root menu only
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                self->InjectContextMenu(reinterpret_cast<HMENU>(wParam));
                return result;
            }
            break;
        case WM_COMMAND:
            // Only reached when DefView tracks without TPM_RETURNCMD.
            if (HIWORD(wParam) == 0 && HandlePopupCommand(LOWORD(wParam))) {
                return 0;
            }
            break;
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
            if (self) {
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                self->ApplyListViewTheme();
                return result;
            }
            break;
        case WM_PARENTNOTIFY:
            // View mode switches can recreate the list view.
            if (self && LOWORD(wParam) == WM_CREATE) {
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                HWND created = reinterpret_cast<HWND>(lParam);
                if (IsClass(created, WC_LISTVIEWW)) {
                    self->m_listView = created;
                    self->ApplyListViewTheme();
                }
                return result;
            }
            break;
        case WM_NCDESTROY:
            RemoveWindowSubclass(hwnd, &NameColorView::DefViewSubclassProc, id);
            if (self && self->m_defView == hwnd) {
                self->m_defView = nullptr;
                self->m_listView = nullptr;
            }
            break;
        default:
            break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK NameColorView::TreeParentSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                      UINT_PTR id, DWORD_PTR refData) {
    auto* self = reinterpret_cast<NameColorView*>(refData);
    switch (msg) {
        case WM_NOTIFY: {
            auto* header = reinterpret_cast<NMHDR*>(lParam);
            if (!self || !header || header->hwndFrom != self->m_tree) {
                break;
            }
            if (header->code == NM_CUSTOMDRAW) {
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                return self->HandleTreeCustomDraw(reinterpret_cast<NMTVCUSTOMDRAW*>(lParam), result);
            }
            if (header->code == TVN_DELETEITEMW) {
                self->m_treeCache.erase(reinterpret_cast<NMTREEVIEWW*>(lParam)->itemOld.hItem);
            }
            break;
        }
        case WM_NCDESTROY:
            RemoveWindowSubclass(hwnd, &NameColorView::TreeParentSubclassProc, id);
            if (self && self->m_treeParent == hwnd) {
                self->m_treeParent = nullptr;
                self->m_tree = nullptr;
                self->m_treeCache.clear();
            }
            break;
        default:
            break;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

}  // namespace shelltabs
