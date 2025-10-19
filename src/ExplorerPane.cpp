#include "ExplorerPane.h"

#include <CommCtrl.h>
#include <ShlObj.h>

#include <atomic>
#include <string>

#include "NameColorProvider.h"
#include "Utilities.h"

#ifndef RETURN_IF_FAILED
#define RETURN_IF_FAILED(hrcall)                                                                                     \
    do {                                                                                                             \
        HRESULT _hr = (hrcall);                                                                                      \
        if (FAILED(_hr)) {                                                                                           \
            return _hr;                                                                                              \
        }                                                                                                            \
    } while (0)
#endif

using Microsoft::WRL::ComPtr;

namespace shelltabs {
namespace {

constexpr UINT_PTR kPaneSubclassId = 0x504E5342;  // 'PNSB'
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
    DrawTextW(hdc, text.c_str(), static_cast<int>(text.size()), &gutter, format);
    SetTextColor(hdc, previousText);
    SetBkMode(hdc, previousBk);
    return true;
}

}  // namespace

class ExplorerPane::BrowserEvents : public IExplorerBrowserEvents {
public:
    explicit BrowserEvents(ExplorerPane* owner) : m_refCount(1), m_owner(owner) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IExplorerBrowserEvents) {
            *object = static_cast<IExplorerBrowserEvents*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return static_cast<ULONG>(++m_refCount); }

    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = static_cast<ULONG>(--m_refCount);
        if (count == 0) {
            delete this;
        }
        return count;
    }

    IFACEMETHODIMP OnNavigationPending(PCIDLIST_ABSOLUTE) override { return S_OK; }

    IFACEMETHODIMP OnViewCreated(IShellView* view) override {
        if (m_owner) {
            m_owner->HandleViewCreated(view);
        }
        return S_OK;
    }

    IFACEMETHODIMP OnNavigationComplete(PCIDLIST_ABSOLUTE pidl) override {
        if (m_owner) {
            m_owner->HandleNavigationCompleted(pidl);
        }
        return S_OK;
    }

    IFACEMETHODIMP OnNavigationFailed(PCIDLIST_ABSOLUTE) override { return S_OK; }

private:
    std::atomic<ULONG> m_refCount;
    ExplorerPane* m_owner;
};

ExplorerPane::ExplorerPane() = default;

ExplorerPane::~ExplorerPane() { Destroy(); }

HRESULT ExplorerPane::Create(HWND parent, const RECT& rc) {
    Destroy();

    FOLDERSETTINGS fs{};
    fs.ViewMode = FVM_DETAILS;
    fs.fFlags = FWF_SHOWSELALWAYS;

    RETURN_IF_FAILED(CoCreateInstance(CLSID_ExplorerBrowser, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_browser)));
    RETURN_IF_FAILED(m_browser->Initialize(parent, const_cast<RECT*>(&rc), &fs));

    ComPtr<IOleWindow> ole;
    RETURN_IF_FAILED(m_browser.As(&ole));
    RETURN_IF_FAILED(ole->GetWindow(&m_hwnd));

    auto* events = new (std::nothrow) BrowserEvents(this);
    if (!events) {
        Destroy();
        return E_OUTOFMEMORY;
    }

    m_events.Attach(events);
    RETURN_IF_FAILED(m_browser->Advise(m_events.Get(), &m_adviseCookie));

    return S_OK;
}

void ExplorerPane::Destroy() {
    RemoveSubclass();
    m_folderView.Reset();

    if (m_browser && m_adviseCookie != 0) {
        m_browser->Unadvise(m_adviseCookie);
        m_adviseCookie = 0;
    }
    m_events.Reset();

    if (m_browser) {
        m_browser->Destroy();
        m_browser.Reset();
    }

    m_hwnd = nullptr;
    m_defView = nullptr;
    m_listView = nullptr;
    m_currentPath.clear();
}

HWND ExplorerPane::GetHwnd() const { return m_hwnd; }

HWND ExplorerPane::GetListViewHwnd() const { return m_listView; }

void ExplorerPane::InvalidateView() const {
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, TRUE);
        return;
    }
    if (m_defView && IsWindow(m_defView)) {
        InvalidateRect(m_defView, nullptr, TRUE);
        return;
    }
    if (m_hwnd && IsWindow(m_hwnd)) {
        InvalidateRect(m_hwnd, nullptr, TRUE);
    }
}

HRESULT ExplorerPane::SetRect(const RECT& rc) { return m_browser ? m_browser->SetRect(nullptr, rc) : E_FAIL; }

HRESULT ExplorerPane::NavigateToPIDL(PCIDLIST_ABSOLUTE pidl) {
    if (!m_browser || !pidl) {
        return E_POINTER;
    }
    ComPtr<IShellItem> item;
    RETURN_IF_FAILED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&item)));
    return m_browser->BrowseToObject(item.Get(), 0);
}

HRESULT ExplorerPane::NavigateToShellItem(IShellItem* item) {
    return m_browser ? m_browser->BrowseToObject(item, 0) : E_FAIL;
}

HRESULT ExplorerPane::NavigateToPath(const std::wstring& path) {
    if (!m_browser) {
        return E_FAIL;
    }
    auto pidl = ParseDisplayName(path);
    if (!pidl) {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }
    return NavigateToPIDL(pidl.get());
}

void ExplorerPane::SetNavigationCallback(NavigationCallback callback) { m_onNavigate = std::move(callback); }

std::wstring ExplorerPane::CurrentPath() const { return m_currentPath; }

void ExplorerPane::HandleNavigationCompleted(PCIDLIST_ABSOLUTE pidl) {
    std::wstring parsing = GetParsingName(pidl);
    if (parsing.empty()) {
        parsing = GetDisplayName(pidl);
    }
    m_currentPath = std::move(parsing);
    if (m_onNavigate) {
        m_onNavigate(m_currentPath);
    }

    if (!m_subclassed && m_browser) {
        ComPtr<IShellView> view;
        if (SUCCEEDED(m_browser->GetCurrentView(IID_PPV_ARGS(&view))) && view) {
            HandleViewCreated(view.Get());
        }
    }
}

void ExplorerPane::HandleViewCreated(IShellView* view) {
    RemoveSubclass();

    if (!view) {
        return;
    }

    m_folderView.Reset();
    view->QueryInterface(IID_PPV_ARGS(&m_folderView));

    HWND hwndView = nullptr;
    if (FAILED(view->GetWindow(&hwndView)) || !hwndView) {
        return;
    }

    HWND listView = FindFolderListView(hwndView);
    if (!listView) {
        return;
    }

    if (SetWindowSubclass(hwndView, ViewSubclassProc, kPaneSubclassId, reinterpret_cast<DWORD_PTR>(this))) {
        m_defView = hwndView;
        m_listView = listView;
        m_subclassed = true;
    }
}

void ExplorerPane::RemoveSubclass() {
    if (m_subclassed && m_defView) {
        RemoveWindowSubclass(m_defView, ViewSubclassProc, kPaneSubclassId);
    }
    m_subclassed = false;
    m_defView = nullptr;
    m_listView = nullptr;
}

bool ExplorerPane::HandleNotify(NMHDR* header, LRESULT* result) {
    if (!header || header->hwndFrom != m_listView || header->code != NM_CUSTOMDRAW) {
        return false;
    }

    auto* customDraw = reinterpret_cast<NMLVCUSTOMDRAW*>(header);
    return HandleCustomDraw(customDraw, result);
}

bool ExplorerPane::HandleCustomDraw(NMLVCUSTOMDRAW* cd, LRESULT* result) {
    if (!cd || !result) {
        return false;
    }

    switch (cd->nmcd.dwDrawStage) {
    case CDDS_PREPAINT:
        *result = CDRF_NOTIFYSUBITEMDRAW | CDRF_NOTIFYPOSTPAINT;
        return true;

    case CDDS_SUBITEM | CDDS_ITEMPOSTPAINT:
        if (DrawLineNumberOverlay(m_listView, cd->nmcd.hdc, cd)) {
            *result = CDRF_DODEFAULT;
            return true;
        }
        break;
    default:
        break;
    }

    return false;
}

bool ExplorerPane::GetItemPath(int index, std::wstring* path) const {
    if (!path || !m_folderView) {
        return false;
    }

    ComPtr<IShellItem> item;
    if (FAILED(m_folderView->GetItem(index, IID_PPV_ARGS(&item))) || !item) {
        return false;
    }

    return TryGetFileSystemPath(item.Get(), path);
}

LRESULT CALLBACK ExplorerPane::ViewSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR, DWORD_PTR refData) {
    auto* self = reinterpret_cast<ExplorerPane*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg) {
    case WM_NOTIFY: {
        LRESULT result = 0;
        if (self->HandleNotify(reinterpret_cast<NMHDR*>(lp), &result)) {
            return result;
        }
        break;
    }
    case WM_DESTROY:
    case WM_NCDESTROY:
        self->RemoveSubclass();
        break;
    default:
        break;
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

}  // namespace shelltabs
