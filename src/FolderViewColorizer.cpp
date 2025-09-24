#include "FolderViewColorizer.h"

#include <ShlObj.h>

#include "Tagging.h"

namespace shelltabs {
namespace {
constexpr UINT_PTR kSubclassId = 0x53485354;  // 'SHST'

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
}

FolderViewColorizer::FolderViewColorizer() = default;

FolderViewColorizer::~FolderViewColorizer() { Detach(); }

void FolderViewColorizer::Attach(const Microsoft::WRL::ComPtr<IShellBrowser>& browser) {
    if (browser == m_shellBrowser) {
        Refresh();
        return;
    }
    Detach();
    m_shellBrowser = browser;
    Refresh();
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
    if (FAILED(m_shellBrowser->QueryActiveShellView(IID_PPV_ARGS(&view))) || !view) {
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

bool FolderViewColorizer::HandleCustomDraw(NMLVCUSTOMDRAW* customDraw, LRESULT* result) {
    if (!customDraw || !result) {
        return false;
    }

    switch (customDraw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
            *result = CDRF_NOTIFYITEMDRAW;
            return true;
        case CDDS_ITEMPREPAINT: {
            const int index = static_cast<int>(customDraw->nmcd.dwItemSpec);
            std::wstring path;
            if (!GetItemPath(index, &path)) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            COLORREF color = RGB(0, 0, 0);
            if (TagStore::Instance().TryGetColorForPath(path, &color)) {
                customDraw->clrText = color;
            }
            *result = CDRF_DODEFAULT;
            return true;
        }
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

