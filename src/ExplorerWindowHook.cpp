#include "ExplorerWindowHook.h"

#include <algorithm>
#include <unordered_map>
#include <vector>

#include <optional>
#include <string>
#include <cwchar>

#include <Ole2.h>
#include <ShlObj.h>
#include <shellapi.h>

#include "NameColorProvider.h"
#include "FileColorOverrides.h"
#include "Utilities.h"

#pragma comment(lib, "Comctl32.lib")

namespace shelltabs {
namespace {
std::mutex g_hookMutex;
std::unordered_map<DWORD, std::vector<ExplorerWindowHook*>> g_hooksByThread;
std::unordered_map<HWND, ExplorerWindowHook*> g_hooksByFrame;

constexpr wchar_t kTreeClassName[] = L"SysTreeView32";
constexpr wchar_t kListClassName[] = L"SysListView32";
constexpr wchar_t kDefViewClassName[] = L"SHELLDLL_DefView";
constexpr GUID kSidDataObject = {0x000214e8, 0x0000, 0x0000,
                                 {0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
constexpr int kLineNumberMargin = 4;

std::wstring DirectoryFromPath(const std::wstring& path) {
    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring::npos) {
        return {};
    }
    return path.substr(0, separator);
}

std::wstring CombineDirectoryAndLeaf(const std::wstring& directory, const std::wstring& leaf) {
    if (leaf.empty()) {
        return {};
    }

    if (directory.empty()) {
        return leaf;
    }

    std::wstring combined = directory;
    const wchar_t last = combined.empty() ? L'\0' : combined.back();
    if (last != L'\\' && last != L'/') {
        combined.push_back(L'\\');
    }
    combined.append(leaf);
    return combined;
}

std::wstring AnsiToWide(const char* text) {
    if (!text) {
        return {};
    }

    const int required = MultiByteToWideChar(CP_ACP, 0, text, -1, nullptr, 0);
    if (required <= 0) {
        return {};
    }

    std::wstring wide(static_cast<size_t>(required - 1), L'\0');
    if (!wide.empty()) {
        MultiByteToWideChar(CP_ACP, 0, text, -1, wide.data(), required);
    }
    return wide;
}

bool IsWindowClass(HWND hwnd, const wchar_t* expected) {
    if (!hwnd || !expected) {
        return false;
    }
    wchar_t buffer[128] = {};
    if (!GetClassNameW(hwnd, buffer, ARRAYSIZE(buffer))) {
        return false;
    }
    return lstrcmpiW(buffer, expected) == 0;
}

HWND FindAncestorWithClass(HWND hwnd, const wchar_t* className) {
    HWND current = hwnd;
    while (current) {
        if (IsWindowClass(current, className)) {
            return current;
        }
        current = GetParent(current);
    }
    return nullptr;
}

HWND FindDescendantWithClass(HWND parent, const wchar_t* className) {
    if (!parent || !className) {
        return nullptr;
    }

    HWND child = FindWindowExW(parent, nullptr, className, nullptr);
    if (child) {
        return child;
    }

    struct EnumData {
        const wchar_t* targetClass = nullptr;
        HWND result = nullptr;
    } data{className, nullptr};

    EnumChildWindows(parent,
                     [](HWND hwnd, LPARAM param) -> BOOL {
                         auto* data = reinterpret_cast<EnumData*>(param);
                         if (!data || data->result) {
                             return FALSE;
                         }
                         HWND found = FindDescendantWithClass(hwnd, data->targetClass);
                         if (found) {
                             data->result = found;
                             return FALSE;
                         }
                         return TRUE;
                     },
                     reinterpret_cast<LPARAM>(&data));

    return data.result;
}

bool DrawLineNumberOverlay(HWND listView, HDC hdc, const NMLVCUSTOMDRAW* cd) {
    if (!listView || !cd) {
        return false;
    }

    if ((cd->nmcd.dwDrawStage & CDDS_SUBITEM) == 0 || cd->iSubItem != 0) {
        return false;
    }

    const int index = static_cast<int>(cd->nmcd.dwItemSpec);
    RECT bounds{};
    RECT label{};
    if (!ListView_GetItemRect(listView, index, &bounds, LVIR_BOUNDS) ||
        !ListView_GetSubItemRect(listView, index, 0, LVIR_LABEL, &label)) {
        return false;
    }

    RECT gutter = bounds;
    gutter.left += kLineNumberMargin;
    gutter.right = label.left - kLineNumberMargin;
    if (gutter.right <= gutter.left) {
        return false;
    }

    const std::wstring text = std::to_wstring(index + 1);
    const UINT format = DT_RIGHT | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS;

    const bool selected = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;
    const bool hot = (cd->nmcd.uItemState & CDIS_HOT) != 0;
    COLORREF foreground = GetSysColor(COLOR_WINDOWTEXT);
    if (selected) {
        foreground = GetSysColor(COLOR_HIGHLIGHTTEXT);
    } else if (hot) {
        foreground = GetSysColor(COLOR_HOTLIGHT);
    }

    const int previousBk = SetBkMode(hdc, TRANSPARENT);
    const COLORREF previousText = SetTextColor(hdc, foreground);
    if (!DrawTextW(hdc, text.c_str(), static_cast<int>(text.size()), &gutter, format)) {
        SetTextColor(hdc, previousText);
        SetBkMode(hdc, previousBk);
        return false;
    }
    SetTextColor(hdc, previousText);
    SetBkMode(hdc, previousBk);
    return true;
}

}  // namespace

void ExplorerWindowHook::ScopedFont::Reset() {
    if (handle) {
        DeleteObject(handle);
        handle = nullptr;
    }
}

void ExplorerWindowHook::ScopedFont::Adopt(HFONT font) {
    if (handle == font) {
        return;
    }
    Reset();
    handle = font;
}

ExplorerWindowHook::ExplorerWindowHook() = default;

ExplorerWindowHook::~ExplorerWindowHook() { Shutdown(); }

bool ExplorerWindowHook::Initialize(IUnknown* site, IWebBrowser2* browser) {
    Shutdown();

    if (site) {
        site_ = site;
        site_.As(&serviceProvider_);
    }

    if (browser) {
        browser_ = browser;
    }

    if (!serviceProvider_ && browser_) {
        browser_.As(&serviceProvider_);
    }

    if (serviceProvider_) {
        serviceProvider_->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&shellBrowser_));
    }

    if (!browser_) {
        return false;
    }

    SHANDLE_PTR handle = 0;
    if (FAILED(browser_->get_HWND(&handle)) || !handle) {
        return false;
    }

    frame_ = reinterpret_cast<HWND>(handle);
    if (!IsWindow(frame_)) {
        frame_ = nullptr;
        return false;
    }

    threadId_ = GetWindowThreadProcessId(frame_, nullptr);
    if (threadId_ == 0) {
        Shutdown();
        return false;
    }

    RegisterThreadHook();
    cbtHook_ = SetWindowsHookExW(WH_CBT, &ExplorerWindowHook::CbtHookProc, nullptr, threadId_);
    if (!cbtHook_) {
        UnregisterThreadHook();
        Shutdown();
        return false;
    }

    RegisterFrameHook();

    ApplyTreeTheme();
    ApplyListTheme();
    Attach();
    return true;
}

void ExplorerWindowHook::Shutdown() {
    UnregisterFrameHook();
    if (cbtHook_) {
        UnhookWindowsHookEx(cbtHook_);
        cbtHook_ = nullptr;
    }

    DetachTreeParent();
    DetachDefView();

    ResetFolderView();
    ResetPendingListRename();

    treeFont_.Reset();
    listFont_.Reset();

    UnregisterThreadHook();

    frame_ = nullptr;
    tree_ = nullptr;
    treeParent_ = nullptr;
    defView_ = nullptr;
    listView_ = nullptr;

    shellBrowser_.Reset();
    folderView_.Reset();
    serviceProvider_.Reset();
    browser_.Reset();
    site_.Reset();
    threadId_ = 0;
}

void ExplorerWindowHook::Attach() {
    if (!frame_) {
        return;
    }

    if (!EnsureFolderView()) {
        return;
    }

    Microsoft::WRL::ComPtr<IShellView> shellView;
    if (FAILED(folderView_.As(&shellView)) || !shellView) {
        return;
    }

    HWND defView = nullptr;
    if (FAILED(shellView->GetWindow(&defView)) || !IsWindow(defView)) {
        return;
    }

    HWND list = FindDescendantWithClass(defView, kListClassName);
    if (list) {
        OnListWindowCreated(list);
    }
}

void ExplorerWindowHook::RegisterThreadHook() {
    if (!threadId_) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_hookMutex);
    auto& hooks = g_hooksByThread[threadId_];
    if (std::find(hooks.begin(), hooks.end(), this) == hooks.end()) {
        hooks.push_back(this);
    }
}

void ExplorerWindowHook::UnregisterThreadHook() {
    if (!threadId_) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_hookMutex);
    auto it = g_hooksByThread.find(threadId_);
    if (it != g_hooksByThread.end()) {
        auto& hooks = it->second;
        hooks.erase(std::remove(hooks.begin(), hooks.end(), this), hooks.end());
        if (hooks.empty()) {
            g_hooksByThread.erase(it);
        }
    }
}

void ExplorerWindowHook::RegisterFrameHook() {
    if (!frame_) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_hookMutex);
    g_hooksByFrame[frame_] = this;
}

void ExplorerWindowHook::UnregisterFrameHook() {
    if (!frame_) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_hookMutex);
    auto it = g_hooksByFrame.find(frame_);
    if (it != g_hooksByFrame.end() && it->second == this) {
        g_hooksByFrame.erase(it);
    }
}

LRESULT CALLBACK ExplorerWindowHook::CbtHookProc(int code, WPARAM wParam, LPARAM lParam) {
    if (code == HCBT_CREATEWND) {
        std::vector<ExplorerWindowHook*> hooks;
        {
            std::lock_guard<std::mutex> guard(g_hookMutex);
            auto it = g_hooksByThread.find(GetCurrentThreadId());
            if (it != g_hooksByThread.end()) {
                hooks = it->second;
            }
        }

        if (!hooks.empty()) {
            auto* createInfo = reinterpret_cast<CBT_CREATEWND*>(lParam);
            HWND created = reinterpret_cast<HWND>(wParam);
            for (ExplorerWindowHook* hook : hooks) {
                if (hook) {
                    hook->HandleCreateWindow(created, createInfo);
                }
            }
        }
    }

    return CallNextHookEx(nullptr, code, wParam, lParam);
}

LRESULT CALLBACK ExplorerWindowHook::TreeParentSubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                                            UINT_PTR, DWORD_PTR refData) {
    auto* self = reinterpret_cast<ExplorerWindowHook*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, message, wParam, lParam);
    }

    switch (message) {
        case WM_NOTIFY: {
            LRESULT handled = 0;
            NotifyResult response =
                    self->HandleTreeNotify(reinterpret_cast<NMHDR*>(lParam), &handled);
            if (response == NotifyResult::kHandled) {
                return handled;
            }

            LRESULT forwarded = DefSubclassProc(hwnd, message, wParam, lParam);
            if (response == NotifyResult::kModify) {
                return handled | forwarded;
            }
            return forwarded;
        }
        case WM_DESTROY:
        case WM_NCDESTROY:
            self->DetachTreeParent();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

LRESULT CALLBACK ExplorerWindowHook::DefViewSubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                                         UINT_PTR, DWORD_PTR refData) {
    auto* self = reinterpret_cast<ExplorerWindowHook*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, message, wParam, lParam);
    }

    switch (message) {
        case WM_NOTIFY: {
            LRESULT handled = 0;
            NotifyResult response =
                    self->HandleListNotify(reinterpret_cast<NMHDR*>(lParam), &handled);
            if (response == NotifyResult::kHandled) {
                return handled;
            }

            LRESULT forwarded = DefSubclassProc(hwnd, message, wParam, lParam);
            if (response == NotifyResult::kModify) {
                return handled | forwarded;
            }
            return forwarded;
        }
        case WM_DESTROY:
        case WM_NCDESTROY:
            self->DetachDefView();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

void ExplorerWindowHook::HandleCreateWindow(HWND created, const CBT_CREATEWND* data) {
    UNREFERENCED_PARAMETER(data);
    if (!created || !IsWindow(created)) {
        return;
    }

    const HWND root = GetAncestor(created, GA_ROOT);
    if (!root || root != frame_) {
        return;
    }

    if (IsWindowClass(created, kTreeClassName)) {
        OnTreeWindowCreated(created);
    } else if (IsWindowClass(created, kListClassName)) {
        OnListWindowCreated(created);
    }
}

void ExplorerWindowHook::OnTreeWindowCreated(HWND tree) {
    if (tree_ == tree) {
        return;
    }

    DetachTreeParent();

    tree_ = tree;
    ApplyTreeTheme();

    HWND parent = GetParent(tree_);
    if (!parent) {
        parent = FindAncestorWithClass(tree_, L"CtrlNotifySink");
    }

    AttachTreeParent(parent);
}

void ExplorerWindowHook::OnListWindowCreated(HWND list) {
    if (listView_ == list) {
        return;
    }

    HWND defView = GetParent(list);
    if (!defView || !IsWindowClass(defView, kDefViewClassName)) {
        defView = FindAncestorWithClass(list, kDefViewClassName);
    }

    if (!defView) {
        return;
    }

    HWND cabinet = GetAncestor(defView, GA_ROOT);
    if (cabinet != frame_) {
        return;
    }

    DetachDefView();

    listView_ = list;

    ApplyListTheme();
    AttachDefView(defView);
}

void ExplorerWindowHook::AttachTreeParent(HWND parent) {
    if (!parent) {
        return;
    }

    if (treeParent_ && treeParent_ != parent && treeSubclassed_) {
        RemoveWindowSubclass(treeParent_, TreeParentSubclassProc, kTreeSubclassId);
        treeSubclassed_ = false;
    }

    treeParent_ = parent;

    if (!treeSubclassed_) {
        if (SetWindowSubclass(parent, TreeParentSubclassProc, kTreeSubclassId, reinterpret_cast<DWORD_PTR>(this))) {
            treeSubclassed_ = true;
        }
    }
}

void ExplorerWindowHook::DetachTreeParent() {
    if (treeParent_ && treeSubclassed_) {
        RemoveWindowSubclass(treeParent_, TreeParentSubclassProc, kTreeSubclassId);
    }
    treeParent_ = nullptr;
    treeSubclassed_ = false;
    tree_ = nullptr;
}

void ExplorerWindowHook::AttachDefView(HWND defView) {
    if (!defView) {
        return;
    }

    if (SetWindowSubclass(defView, DefViewSubclassProc, kListSubclassId, reinterpret_cast<DWORD_PTR>(this))) {
        defViewSubclassed_ = true;
    } else {
        defViewSubclassed_ = false;
    }

    defView_ = defView;

    if (listView_ && listFont_.Get()) {
        SendMessageW(listView_, WM_SETFONT, reinterpret_cast<WPARAM>(listFont_.Get()), TRUE);
    }
}

void ExplorerWindowHook::DetachDefView() {
    if (defView_ && defViewSubclassed_) {
        RemoveWindowSubclass(defView_, DefViewSubclassProc, kListSubclassId);
    }
    defViewSubclassed_ = false;
    defView_ = nullptr;
    listView_ = nullptr;
    ResetFolderView();
}

ExplorerWindowHook::NotifyResult ExplorerWindowHook::HandleTreeNotify(NMHDR* header, LRESULT* result) {
    if (!header || header->hwndFrom != tree_ || header->code != NM_CUSTOMDRAW) {
        return NotifyResult::kUnhandled;
    }

    auto* custom = reinterpret_cast<NMTVCUSTOMDRAW*>(header);
    return HandleTreeCustomDraw(custom, result);
}

ExplorerWindowHook::NotifyResult ExplorerWindowHook::HandleTreeCustomDraw(NMTVCUSTOMDRAW* customDraw,
                                                                          LRESULT* result) {
    if (!customDraw || !result) {
        return NotifyResult::kUnhandled;
    }

    switch (customDraw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
            *result |= CDRF_NOTIFYITEMDRAW;
            return NotifyResult::kModify;
        case CDDS_ITEMPREPAINT: {
            bool changed = false;
            const bool selected = (customDraw->nmcd.uItemState & CDIS_SELECTED) != 0;
            const bool hot = (customDraw->nmcd.uItemState & CDIS_HOT) != 0;

            if (selected) {
                if (treeTheme_.selectedTextColor != CLR_INVALID) {
                    customDraw->clrText = treeTheme_.selectedTextColor;
                    changed = true;
                }
                if (treeTheme_.selectedBackgroundColor != CLR_INVALID) {
                    customDraw->clrTextBk = treeTheme_.selectedBackgroundColor;
                    changed = true;
                }
            } else if (hot) {
                if (treeTheme_.hotTextColor != CLR_INVALID) {
                    customDraw->clrText = treeTheme_.hotTextColor;
                    changed = true;
                }
                if (treeTheme_.hotBackgroundColor != CLR_INVALID) {
                    customDraw->clrTextBk = treeTheme_.hotBackgroundColor;
                    changed = true;
                }
            } else {
                if (treeTheme_.textColor != CLR_INVALID) {
                    customDraw->clrText = treeTheme_.textColor;
                    changed = true;
                }
                if (treeTheme_.backgroundColor != CLR_INVALID) {
                    customDraw->clrTextBk = treeTheme_.backgroundColor;
                    changed = true;
                }
            }

            if (treeFont_.Get()) {
                changed = true;
            }

            if (changed) {
                *result |= CDRF_NEWFONT;
                return NotifyResult::kHandled;
            }
            return NotifyResult::kUnhandled;
        }
        default:
            break;
    }

    return NotifyResult::kUnhandled;
}

ExplorerWindowHook::NotifyResult ExplorerWindowHook::HandleListNotify(NMHDR* header, LRESULT* result) {
    if (!header || header->hwndFrom != listView_) {
        return NotifyResult::kUnhandled;
    }

    switch (header->code) {
        case NM_CUSTOMDRAW: {
            auto* custom = reinterpret_cast<NMLVCUSTOMDRAW*>(header);
            return HandleListCustomDraw(custom, result);
        }
        case LVN_BEGINLABELEDITW:
        case LVN_BEGINLABELEDITA:
        case LVN_ENDLABELEDITW:
        case LVN_ENDLABELEDITA:
            return HandleListLabelEdit(header);
        default:
            break;
    }

    return NotifyResult::kUnhandled;
}

ExplorerWindowHook::NotifyResult ExplorerWindowHook::HandleListCustomDraw(NMLVCUSTOMDRAW* customDraw,
                                                                          LRESULT* result) {
    if (!customDraw || !result) {
        return NotifyResult::kUnhandled;
    }

    switch (customDraw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
            *result |= CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW | CDRF_NOTIFYPOSTPAINT;
            return NotifyResult::kModify;
        case CDDS_ITEMPREPAINT:
        case CDDS_SUBITEM | CDDS_ITEMPREPAINT: {
            if (customDraw->iSubItem != 0) {
                return NotifyResult::kUnhandled;
            }

            const int index = static_cast<int>(customDraw->nmcd.dwItemSpec);
            const bool selected = (customDraw->nmcd.uItemState & CDIS_SELECTED) != 0;
            const bool hot = (customDraw->nmcd.uItemState & CDIS_HOT) != 0;

            bool changed = false;

            std::optional<NameColorProvider::ItemAppearance> appearance;
            std::wstring path;
            if (GetListViewItemPath(index, &path)) {
                const auto resolved = NameColorProvider::Instance().GetAppearanceForPath(path);
                if (resolved.HasOverrides() &&
                    resolved.AllowsForState(customDraw->nmcd.uItemState)) {
                    appearance = resolved;
                }
            }

            const PaneTheme& theme = listTheme_;
            if (selected) {
                if (theme.selectedTextColor != CLR_INVALID) {
                    customDraw->clrText = theme.selectedTextColor;
                    changed = true;
                }
                if (theme.selectedBackgroundColor != CLR_INVALID) {
                    customDraw->clrTextBk = theme.selectedBackgroundColor;
                    changed = true;
                }
            } else if (hot) {
                if (theme.hotTextColor != CLR_INVALID) {
                    customDraw->clrText = theme.hotTextColor;
                    changed = true;
                }
                if (theme.hotBackgroundColor != CLR_INVALID) {
                    customDraw->clrTextBk = theme.hotBackgroundColor;
                    changed = true;
                }
            } else {
                if (theme.textColor != CLR_INVALID) {
                    customDraw->clrText = theme.textColor;
                    changed = true;
                }
                if (theme.backgroundColor != CLR_INVALID) {
                    customDraw->clrTextBk = theme.backgroundColor;
                    changed = true;
                }
            }

            if (appearance.has_value()) {
                if (appearance->textColor.has_value()) {
                    customDraw->clrText = *appearance->textColor;
                    changed = true;
                }
                if (appearance->backgroundColor.has_value()) {
                    customDraw->clrTextBk = *appearance->backgroundColor;
                    changed = true;
                }
                if (appearance->font) {
                    SelectObject(customDraw->nmcd.hdc, appearance->font);
                    changed = true;
                }
            }

            if (changed || listFont_.Get()) {
                *result |= CDRF_NEWFONT;
                return NotifyResult::kHandled;
            }

            return NotifyResult::kUnhandled;
        }
        case CDDS_SUBITEM | CDDS_ITEMPOSTPAINT: {
            if (customDraw->iSubItem == 0 &&
                DrawLineNumberOverlay(listView_, customDraw->nmcd.hdc, customDraw)) {
                *result |= CDRF_DODEFAULT;
                return NotifyResult::kHandled;
            }
            return NotifyResult::kUnhandled;
        }
        default:
            break;
    }

    return NotifyResult::kUnhandled;
}

ExplorerWindowHook::NotifyResult ExplorerWindowHook::HandleListLabelEdit(NMHDR* header) {
    if (!header) {
        return NotifyResult::kUnhandled;
    }

    switch (header->code) {
        case LVN_BEGINLABELEDITW: {
            const auto* edit = reinterpret_cast<const NMLVDISPINFOW*>(header);
            if (edit) {
                RememberListItemForRename(edit->item.iItem);
            } else {
                ResetPendingListRename();
            }
            break;
        }
        case LVN_BEGINLABELEDITA: {
            const auto* edit = reinterpret_cast<const NMLVDISPINFOA*>(header);
            if (edit) {
                RememberListItemForRename(edit->item.iItem);
            } else {
                ResetPendingListRename();
            }
            break;
        }
        case LVN_ENDLABELEDITW: {
            const auto* edit = reinterpret_cast<const NMLVDISPINFOW*>(header);
            if (edit && edit->item.pszText) {
                CommitListRename(edit->item.pszText);
            } else {
                ResetPendingListRename();
            }
            break;
        }
        case LVN_ENDLABELEDITA: {
            const auto* edit = reinterpret_cast<const NMLVDISPINFOA*>(header);
            if (edit && edit->item.pszText) {
                CommitListRename(AnsiToWide(edit->item.pszText));
            } else {
                ResetPendingListRename();
            }
            break;
        }
        default:
            break;
    }

    return NotifyResult::kUnhandled;
}

void ExplorerWindowHook::RememberListItemForRename(int index) {
    ResetPendingListRename();

    if (index < 0) {
        return;
    }

    std::wstring path;
    if (!GetListViewItemPath(index, &path) || path.empty()) {
        return;
    }

    pendingListRenameOriginalPath_ = path;
    pendingListRenameDirectory_ = DirectoryFromPath(path);
}

void ExplorerWindowHook::CommitListRename(const std::wstring& newName) {
    if (pendingListRenameOriginalPath_.empty() || newName.empty()) {
        ResetPendingListRename();
        return;
    }

    std::wstring directory = pendingListRenameDirectory_;
    if (directory.empty()) {
        directory = DirectoryFromPath(pendingListRenameOriginalPath_);
    }

    const std::wstring newPath = CombineDirectoryAndLeaf(directory, newName);
    if (!newPath.empty() &&
        (_wcsicmp(pendingListRenameOriginalPath_.c_str(), newPath.c_str()) != 0)) {
        FileColorOverrides::Instance().TransferColor(pendingListRenameOriginalPath_, newPath);
    }

    ResetPendingListRename();
}

void ExplorerWindowHook::ResetPendingListRename() {
    pendingListRenameOriginalPath_.clear();
    pendingListRenameDirectory_.clear();
}

bool ExplorerWindowHook::EnsureFolderView() {
    if (folderView_) {
        return true;
    }

    if (!shellBrowser_) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellView> view;
    if (FAILED(shellBrowser_->QueryActiveShellView(&view)) || !view) {
        return false;
    }

    return SUCCEEDED(view.As(&folderView_));
}

void ExplorerWindowHook::ResetFolderView() { folderView_.Reset(); }

bool ExplorerWindowHook::GetListViewItemPath(int index, std::wstring* path) const {
    if (!path) {
        return false;
    }

    if (!const_cast<ExplorerWindowHook*>(this)->EnsureFolderView()) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItem> item;
    if (FAILED(folderView_->GetItem(index, IID_PPV_ARGS(&item))) || !item) {
        return false;
    }

    return ItemPathFromShellItem(item.Get(), path);
}

bool ExplorerWindowHook::ItemPathFromShellItem(IShellItem* item, std::wstring* path) {
    if (!item || !path) {
        return false;
    }

    return TryGetFileSystemPath(item, path);
}

Microsoft::WRL::ComPtr<IShellItemArray> ExplorerWindowHook::BuildArrayFromDataObject(IDataObject* dataObject) {
    Microsoft::WRL::ComPtr<IShellItemArray> array;
    if (!dataObject) {
        return array;
    }

    if (SUCCEEDED(SHCreateShellItemArrayFromDataObject(dataObject, IID_PPV_ARGS(&array))) && array) {
        return array;
    }

    FORMATETC format = {CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM storage = {};
    if (FAILED(dataObject->GetData(&format, &storage))) {
        return array;
    }

    HDROP drop = static_cast<HDROP>(GlobalLock(storage.hGlobal));
    if (!drop) {
        ReleaseStgMedium(&storage);
        return array;
    }

    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
    std::vector<PIDLIST_ABSOLUTE> pidls;
    pidls.reserve(count);

    for (UINT i = 0; i < count; ++i) {
        UINT length = DragQueryFileW(drop, i, nullptr, 0);
        if (!length) {
            continue;
        }

        std::wstring path(length + 1, L'\0');
        UINT copied = DragQueryFileW(drop, i, path.data(), static_cast<UINT>(path.size()));
        if (!copied) {
            continue;
        }
        path.resize(copied);

        PIDLIST_ABSOLUTE pidl = nullptr;
        if (SUCCEEDED(SHParseDisplayName(path.c_str(), nullptr, &pidl, 0, nullptr)) && pidl) {
            pidls.push_back(pidl);
        }
    }

    GlobalUnlock(storage.hGlobal);
    ReleaseStgMedium(&storage);

    if (pidls.empty()) {
        return array;
    }

    std::vector<PCIDLIST_ABSOLUTE> constPidls;
    constPidls.reserve(pidls.size());
    for (PIDLIST_ABSOLUTE pidl : pidls) {
        constPidls.push_back(pidl);
    }

    if (FAILED(SHCreateShellItemArrayFromIDLists(static_cast<UINT>(constPidls.size()), constPidls.data(), &array))) {
        array.Reset();
    }

    for (PIDLIST_ABSOLUTE pidl : pidls) {
        CoTaskMemFree(pidl);
    }

    return array;
}

bool ExplorerWindowHook::AppendPathsFromArray(IShellItemArray* array, std::vector<std::wstring>* outPaths) {
    if (!array || !outPaths) {
        return false;
    }

    DWORD count = 0;
    if (FAILED(array->GetCount(&count)) || count == 0) {
        return false;
    }

    bool any = false;
    for (DWORD i = 0; i < count; ++i) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (SUCCEEDED(array->GetItemAt(i, &item)) && item) {
            std::wstring path;
            if (ItemPathFromShellItem(item.Get(), &path)) {
                outPaths->push_back(std::move(path));
                any = true;
            }
        }
    }
    return any;
}

void ExplorerWindowHook::UpdateListTheme(const PaneTheme& theme) {
    listTheme_ = theme;
    listFont_.Adopt(CreateFontFromTheme(listTheme_));
    ApplyListTheme();
    if (listView_) {
        InvalidateRect(listView_, nullptr, TRUE);
    }
}

void ExplorerWindowHook::UpdateTreeTheme(const PaneTheme& theme) {
    treeTheme_ = theme;
    treeFont_.Adopt(CreateFontFromTheme(treeTheme_));
    ApplyTreeTheme();
    if (tree_) {
        InvalidateRect(tree_, nullptr, TRUE);
    }
}

void ExplorerWindowHook::ApplyTreeTheme() {
    if (!tree_) {
        return;
    }

    if (treeFont_.Get()) {
        SendMessageW(tree_, WM_SETFONT, reinterpret_cast<WPARAM>(treeFont_.Get()), TRUE);
    }

    if (treeTheme_.textColor != CLR_INVALID) {
        TreeView_SetTextColor(tree_, treeTheme_.textColor);
    }
    if (treeTheme_.backgroundColor != CLR_INVALID) {
        TreeView_SetBkColor(tree_, treeTheme_.backgroundColor);
    }
}

void ExplorerWindowHook::ApplyListTheme() {
    if (!listView_) {
        return;
    }

    if (listFont_.Get()) {
        SendMessageW(listView_, WM_SETFONT, reinterpret_cast<WPARAM>(listFont_.Get()), TRUE);
    }

    if (listTheme_.textColor != CLR_INVALID) {
        ListView_SetTextColor(listView_, listTheme_.textColor);
    }
    if (listTheme_.backgroundColor != CLR_INVALID) {
        ListView_SetBkColor(listView_, listTheme_.backgroundColor);
        ListView_SetTextBkColor(listView_, listTheme_.backgroundColor);
    }
}

HFONT ExplorerWindowHook::CreateFontFromTheme(const PaneTheme& theme) {
    if (!theme.useCustomFont) {
        return nullptr;
    }
    return CreateFontIndirectW(&theme.font);
}

bool ExplorerWindowHook::CollectSelection(std::vector<std::wstring>* folderViewPaths,
                                          std::vector<std::wstring>* treePaths) const {
    bool any = false;

    if (folderViewPaths) {
        folderViewPaths->clear();
        Microsoft::WRL::ComPtr<IShellItemArray> array;

        if (const_cast<ExplorerWindowHook*>(this)->EnsureFolderView()) {
            if (FAILED(folderView_->Items(SVGIO_SELECTION, IID_PPV_ARGS(&array))) || !array) {
                array.Reset();
            }
            if (!array && FAILED(folderView_->GetSelection(FALSE, &array))) {
                array.Reset();
            }
            if (!array && FAILED(folderView_->GetSelection(TRUE, &array))) {
                array.Reset();
            }
        }

        if (!array && serviceProvider_) {
            Microsoft::WRL::ComPtr<IDataObject> dataObject;
            if (SUCCEEDED(serviceProvider_->QueryService(kSidDataObject, IID_PPV_ARGS(&dataObject))) && dataObject) {
                array = BuildArrayFromDataObject(dataObject.Get());
            }
        }

        if (!array && shellBrowser_) {
            Microsoft::WRL::ComPtr<IShellView> view;
            if (SUCCEEDED(shellBrowser_->QueryActiveShellView(&view)) && view) {
                if (FAILED(view->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&array))) || !array) {
                    Microsoft::WRL::ComPtr<IDataObject> dataObject;
                    if (SUCCEEDED(view->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&dataObject))) && dataObject) {
                        array = BuildArrayFromDataObject(dataObject.Get());
                    }
                }
            }
        }

        if (AppendPathsFromArray(array.Get(), folderViewPaths)) {
            any = true;
        }
    }

    if (treePaths) {
        treePaths->clear();

        Microsoft::WRL::ComPtr<IShellBrowser> browser = shellBrowser_;
        if (!browser && serviceProvider_) {
            serviceProvider_->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser));
        }

        Microsoft::WRL::ComPtr<IServiceProvider> browserProvider;
        if (browser && SUCCEEDED(browser.As(&browserProvider)) && browserProvider) {
            Microsoft::WRL::ComPtr<INameSpaceTreeControl> treeControl;
            if (SUCCEEDED(browserProvider->QueryService(__uuidof(INameSpaceTreeControl),
                                                        IID_PPV_ARGS(&treeControl))) &&
                treeControl) {
                Microsoft::WRL::ComPtr<IShellItemArray> selection;
#if defined(NSTCGNI_SELECTION)
                if (SUCCEEDED(treeControl->GetSelectedItems(NSTCGNI_SELECTION, &selection)) && selection) {
#else
                if (SUCCEEDED(treeControl->GetSelectedItems(&selection)) && selection) {
#endif
                    if (AppendPathsFromArray(selection.Get(), treePaths)) {
                        any = true;
                    }
                }
            }
        }
    }

    return any;
}

bool ExplorerWindowHook::CollectSelectionForExplorer(HWND explorer, std::vector<std::wstring>* folderViewPaths,
                                                     std::vector<std::wstring>* treePaths) {
    ExplorerWindowHook* hook = nullptr;
    {
        std::lock_guard<std::mutex> guard(g_hookMutex);
        auto it = g_hooksByFrame.find(explorer);
        if (it != g_hooksByFrame.end()) {
            hook = it->second;
        }
    }

    if (!hook) {
        return false;
    }

    return hook->CollectSelection(folderViewPaths, treePaths);
}

void ExplorerWindowHook::AttachForExplorer(HWND explorer) {
    ExplorerWindowHook* hook = nullptr;
    {
        std::lock_guard<std::mutex> guard(g_hookMutex);
        auto it = g_hooksByFrame.find(explorer);
        if (it != g_hooksByFrame.end()) {
            hook = it->second;
        }
    }

    if (hook) {
        hook->Attach();
    }
}

}  // namespace shelltabs
