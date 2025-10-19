#include "ExplorerWindowHook.h"

#include <algorithm>
#include <unordered_map>
#include <vector>

#include <string>
#include <cwchar>

#include <Ole2.h>
#include <ShlObj.h>
#include <shellapi.h>

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
constexpr int kLineNumberMargin = 4;

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

            const bool selected = (customDraw->nmcd.uItemState & CDIS_SELECTED) != 0;
            const bool hot = (customDraw->nmcd.uItemState & CDIS_HOT) != 0;

            bool changed = false;

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
