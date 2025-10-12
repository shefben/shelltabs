#include "FolderViewColorizer.h"

#include <ShlObj.h>
#include <CommCtrl.h>
#include <string>

#include "Tagging.h"
#include "FileColorOverrides.h"

namespace shelltabs {
namespace {
constexpr UINT_PTR kSubclassId = 0x53485354;  // 'SHST'
constexpr int kLineNumberMargin = 4;

HWND FindFolderListView(HWND root) {
    if (!root) {
        return nullptr;
    }

    HWND list = FindWindowExW(root, nullptr, L"SysListView32", nullptr);
    if (list) {
        return list;
    }

    struct EnumData {
        HWND result = nullptr;
    } data;

    EnumChildWindows(
        root, [](HWND hwnd, LPARAM param) -> BOOL {
            auto* d = reinterpret_cast<EnumData*>(param);
            if (!d || d->result) {
                return FALSE;
            }
            wchar_t className[64] = {};
            if (GetClassNameW(hwnd, className, ARRAYSIZE(className)) && lstrcmpiW(className, L"SysListView32") == 0) {
                d->result = hwnd;
                return FALSE;
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&data));

    return data.result;
}

void DrawLineNumberOverlay(HWND listView, HDC hdc, const NMLVCUSTOMDRAW* cd) {
    if (!listView || !cd) {
        return;
    }

    if ((cd->nmcd.dwDrawStage & CDDS_SUBITEM) == 0 || cd->iSubItem != 0) {
        return;
    }

    const int index = static_cast<int>(cd->nmcd.dwItemSpec);
    RECT bounds{};
    RECT label{};
    if (!ListView_GetItemRect(listView, index, &bounds, LVIR_BOUNDS) ||
        !ListView_GetSubItemRect(listView, index, 0, LVIR_LABEL, &label)) {
        return;
    }

    RECT gutter = bounds;
    gutter.left += kLineNumberMargin;
    gutter.right = label.left - kLineNumberMargin;
    if (gutter.right <= gutter.left) {
        return;
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
    DrawTextW(hdc, text.c_str(), static_cast<int>(text.size()), &gutter, format);
    SetTextColor(hdc, previousText);
    SetBkMode(hdc, previousBk);
}
}

FolderViewColorizer::FolderViewColorizer() = default;

FolderViewColorizer::~FolderViewColorizer() { Detach(); }

void FolderViewColorizer::Attach(const Microsoft::WRL::ComPtr<IShellBrowser>& browser) {
    if (browser == m_shellBrowser) {
        ResetView();
        return;
    }
    Detach();
    m_shellBrowser = browser;
}

void FolderViewColorizer::Detach() {
    RemoveSubclass();
    m_folderView.Reset();
    m_shellBrowser.Reset();
    m_defView = nullptr;
    m_listView = nullptr;
}

void FolderViewColorizer::Refresh() {
    RemoveSubclass();
    m_folderView.Reset();
    m_defView = nullptr;
    m_listView = nullptr;

    if (!ResolveView()) {
        return;
    }

    EnsureSubclass();
}

bool FolderViewColorizer::ResolveView() {
    if (!m_shellBrowser) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellView> view;
    if (FAILED(m_shellBrowser->QueryActiveShellView(&view)) || !view) {
        return false;
    }

    if (FAILED(view.As(&m_folderView)) || !m_folderView) {
        return false;
    }

    HWND hwndView = nullptr;
    if (FAILED(view->GetWindow(&hwndView)) || !hwndView) {
        m_folderView.Reset();
        return false;
    }

    HWND listView = FindFolderListView(hwndView);
    if (!listView) {
        m_folderView.Reset();
        return false;
    }

    m_defView = hwndView;
    m_listView = listView;
    return true;
}

void FolderViewColorizer::ResetView() {
    RemoveSubclass();
    m_folderView.Reset();
    m_defView = nullptr;
    m_listView = nullptr;
}

bool FolderViewColorizer::EnsureSubclass() {
    if (!m_defView) {
        return false;
    }
    return SetWindowSubclass(m_defView, &FolderViewColorizer::SubclassProc, kSubclassId,
                             reinterpret_cast<DWORD_PTR>(this)) != FALSE;
}

void FolderViewColorizer::RemoveSubclass() {
    if (m_defView) {
        RemoveWindowSubclass(m_defView, &FolderViewColorizer::SubclassProc, kSubclassId);
    }
}

bool FolderViewColorizer::HandleNotify(NMHDR* header, LRESULT* result) {
    if (!header || header->hwndFrom != m_listView || header->code != NM_CUSTOMDRAW) {
        return false;
    }

    auto* customDraw = reinterpret_cast<NMLVCUSTOMDRAW*>(header);
    return HandleCustomDraw(customDraw, result);
}

bool FolderViewColorizer::HandleCustomDraw(NMLVCUSTOMDRAW* cd, LRESULT* result) {
	if (!cd || !result) return false;

        switch (cd->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
                *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW | CDRF_NOTIFYPOSTPAINT;
                return true;

        case CDDS_ITEMPREPAINT:
        case CDDS_SUBITEM | CDDS_ITEMPREPAINT: {
                const int iItem = static_cast<int>(cd->nmcd.dwItemSpec);
                std::wstring fullPath;
                if (!GetItemPath(iItem, &fullPath)) {
                        *result = CDRF_DODEFAULT;
                        return true;
                }
                COLORREF chosen;
                if (FileColorOverrides::Instance().TryGetColor(fullPath, &chosen)) {
                        if ((cd->nmcd.uItemState & (CDIS_SELECTED | CDIS_HOT)) == 0) {
                                cd->clrText = chosen;
                                *result = CDRF_NEWFONT;
                                return true;
                        }
                }
                *result = CDRF_DODEFAULT;
                return true;
        }
        case CDDS_SUBITEM | CDDS_ITEMPOSTPAINT:
                DrawLineNumberOverlay(m_listView, cd->nmcd.hdc, cd);
                *result = CDRF_DODEFAULT;
                return true;
        default:
                break;
        }
        return false;
}


bool FolderViewColorizer::GetItemPath(int index, std::wstring* path) const {
    if (!path || !m_folderView) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItem> item;
    if (FAILED(m_folderView->GetItem(index, IID_PPV_ARGS(&item))) || !item) {
        return false;
    }

    PWSTR buffer = nullptr;
    HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, &buffer);
    if (FAILED(hr) || !buffer) {
        hr = item->GetDisplayName(SIGDN_DESKTOPABSOLUTEEDITING, &buffer);
    }
    if (FAILED(hr) || !buffer) {
        return false;
    }

    path->assign(buffer);
    CoTaskMemFree(buffer);
    return true;
}

LRESULT CALLBACK FolderViewColorizer::SubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                                   UINT_PTR id, DWORD_PTR refData) {
    auto* self = reinterpret_cast<FolderViewColorizer*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, message, wParam, lParam);
    }

    switch (message) {
        case WM_NOTIFY: {
            LRESULT result = 0;
            if (self->HandleNotify(reinterpret_cast<NMHDR*>(lParam), &result)) {
                return result;
            }
            break;
        }
        case WM_DESTROY:
        case WM_NCDESTROY:
            self->ResetView();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

}  // namespace shelltabs

