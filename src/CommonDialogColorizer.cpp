#include "CommonDialogColorizer.h"

#include <CommCtrl.h>
#include <ShlObj.h>
#include <Shlwapi.h>
#include <ShObjIdl.h>
#include <shlguid.h>
#include <windowsx.h>

#include <atomic>
#include <mutex>
#include <unordered_set>
#include <vector>
#include <new>

#include "NameColorProvider.h"

namespace shelltabs {

class CommonDialogColorizer::DialogEvents : public IFileDialogEvents {
public:
    explicit DialogEvents(CommonDialogColorizer* owner) : m_refCount(1), m_owner(owner) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IFileDialogEvents) {
            *object = static_cast<IFileDialogEvents*>(this);
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

    IFACEMETHODIMP OnFileOk(IFileDialog*) override { return S_OK; }

    IFACEMETHODIMP OnFolderChanging(IFileDialog*, IShellItem*) override { return S_OK; }

    IFACEMETHODIMP OnFolderChange(IFileDialog* /*dialog*/) override {
        if (m_owner) {
            m_owner->UpdateCurrentFolder();
            m_owner->Refresh();
        }
        return S_OK;
    }

    IFACEMETHODIMP OnSelectionChange(IFileDialog*) override {
        if (m_owner) {
            m_owner->InvalidateView();
        }
        return S_OK;
    }

    IFACEMETHODIMP OnShareViolation(IFileDialog*, IShellItem*, FDE_SHAREVIOLATION_RESPONSE* response) override {
        if (response) {
            *response = FDESVR_DEFAULT;
        }
        return S_OK;
    }

    IFACEMETHODIMP OnTypeChange(IFileDialog*) override {
        if (m_owner) {
            m_owner->Refresh();
        }
        return S_OK;
    }

    IFACEMETHODIMP OnOverwrite(IFileDialog*, IShellItem*, FDE_OVERWRITE_RESPONSE* response) override {
        if (response) {
            *response = FDEOR_DEFAULT;
        }
        return S_OK;
    }

private:
    std::atomic<ULONG> m_refCount;
    CommonDialogColorizer* m_owner;
};

namespace {

constexpr UINT_PTR kSubclassId = 0x4344434c;  // 'CDCL'

std::mutex& RegistryMutex() {
    static std::mutex mutex;
    return mutex;
}

std::unordered_set<CommonDialogColorizer*>& Registry() {
    static std::unordered_set<CommonDialogColorizer*> registry;
    return registry;
}

void RegisterColorizer(CommonDialogColorizer* colorizer) {
    if (!colorizer) {
        return;
    }
    auto& registry = Registry();
    std::scoped_lock lock(RegistryMutex());
    registry.insert(colorizer);
}

void UnregisterColorizer(CommonDialogColorizer* colorizer) {
    if (!colorizer) {
        return;
    }
    auto& registry = Registry();
    std::scoped_lock lock(RegistryMutex());
    registry.erase(colorizer);
}

std::wstring GetShellItemDisplayName(IShellItem* item) {
    if (!item) {
        return {};
    }

    PWSTR buffer = nullptr;
    HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, &buffer);
    if (FAILED(hr) || !buffer) {
        hr = item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &buffer);
    }
    if (FAILED(hr) || !buffer) {
        return {};
    }

    std::wstring value(buffer);
    CoTaskMemFree(buffer);
    return value;
}

std::wstring NormalizeFolderPath(const std::wstring& folder) {
    if (folder.empty()) {
        return folder;
    }
    std::wstring normalized = folder;
    if (!normalized.empty() && normalized.back() == L'\\') {
        while (!normalized.empty() && normalized.back() == L'\\') {
            normalized.pop_back();
        }
    }
    return normalized;
}

}  // namespace

CommonDialogColorizer::CommonDialogColorizer() = default;

CommonDialogColorizer::~CommonDialogColorizer() { Detach(); }

bool CommonDialogColorizer::Attach(IFileDialog* dialog) {
    if (!dialog) {
        return false;
    }

    Detach();

    m_dialog = dialog;
    UpdateCurrentFolder();

    auto* events = new (std::nothrow) CommonDialogColorizer::DialogEvents(this);
    if (!events) {
        Detach();
        return false;
    }

    m_events.Attach(events);
    if (FAILED(dialog->Advise(events, &m_adviseCookie))) {
        Detach();
        return false;
    }

    Refresh();
    return true;
}

void CommonDialogColorizer::Detach() {
    RemoveSubclass();
    UnregisterColorizer(this);

    if (m_dialog && m_adviseCookie != 0) {
        m_dialog->Unadvise(m_adviseCookie);
    }
    m_adviseCookie = 0;
    m_events.Reset();
    m_folderView.Reset();
    m_dialog.Reset();
    m_dialogHwnd = nullptr;
    m_defView = nullptr;
    m_listView = nullptr;
    m_subclassed = false;
    m_currentFolder.clear();
}

void CommonDialogColorizer::Refresh() {
    RemoveSubclass();
    m_folderView.Reset();
    m_defView = nullptr;
    m_listView = nullptr;

    if (!ResolveView()) {
        return;
    }

    EnsureSubclass();
    UpdateCurrentFolder();
}

void CommonDialogColorizer::InvalidateView() const {
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, TRUE);
        UpdateWindow(m_listView);
        return;
    }
    if (m_defView && IsWindow(m_defView)) {
        InvalidateRect(m_defView, nullptr, TRUE);
        UpdateWindow(m_defView);
        return;
    }
    if (m_dialogHwnd && IsWindow(m_dialogHwnd)) {
        InvalidateRect(m_dialogHwnd, nullptr, TRUE);
    }
}

void CommonDialogColorizer::NotifyColorDataChanged() {
    std::vector<CommonDialogColorizer*> snapshot;
    {
        std::scoped_lock lock(RegistryMutex());
        auto& registry = Registry();
        snapshot.reserve(registry.size());
        for (auto* entry : registry) {
            snapshot.push_back(entry);
        }
    }

    for (auto* colorizer : snapshot) {
        if (colorizer) {
            colorizer->InvalidateView();
        }
    }
}

void CommonDialogColorizer::UpdateCurrentFolder() {
    m_currentFolder.clear();
    if (!m_dialog) {
        return;
    }

    Microsoft::WRL::ComPtr<IShellItem> folder;
    if (SUCCEEDED(m_dialog->GetFolder(&folder)) && folder) {
        m_currentFolder = NormalizeFolderPath(GetShellItemDisplayName(folder.Get()));
    }
}

bool CommonDialogColorizer::ResolveView() {
    if (!m_dialog) {
        return false;
    }

    HWND hwnd = nullptr;
    Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
    if (FAILED(m_dialog.As(&oleWindow)) || !oleWindow) {
        return false;
    }

    if (FAILED(oleWindow->GetWindow(&hwnd)) || !IsWindow(hwnd)) {
        return false;
    }

    HWND defView = FindDescendantByClass(hwnd, L"SHELLDLL_DefView");
    if (!defView || !IsWindow(defView)) {
        return false;
    }

    HWND listView = FindDescendantByClass(defView, L"SysListView32");
    if (!listView || !IsWindow(listView)) {
        return false;
    }

    m_dialogHwnd = hwnd;
    m_defView = defView;
    m_listView = listView;

    Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
    if (SUCCEEDED(m_dialog.As(&serviceProvider)) && serviceProvider) {
        Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
        if (SUCCEEDED(serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&shellBrowser))) && shellBrowser) {
            Microsoft::WRL::ComPtr<IShellView> view;
            if (SUCCEEDED(shellBrowser->QueryActiveShellView(&view)) && view) {
                view.As(&m_folderView);
            }
        }

        if (!m_folderView) {
            Microsoft::WRL::ComPtr<IExplorerBrowser> explorerBrowser;
            if (SUCCEEDED(serviceProvider->QueryService(SID_SExplorerBrowser, IID_PPV_ARGS(&explorerBrowser))) && explorerBrowser) {
                Microsoft::WRL::ComPtr<IShellView> view;
                if (SUCCEEDED(explorerBrowser->GetCurrentView(IID_PPV_ARGS(&view))) && view) {
                    view.As(&m_folderView);
                }
            }
        }
    }

    return true;
}

bool CommonDialogColorizer::EnsureSubclass() {
    if (!m_defView || !m_listView) {
        return false;
    }

    if (SetWindowSubclass(m_defView, &CommonDialogColorizer::SubclassProc, kSubclassId,
                          reinterpret_cast<DWORD_PTR>(this))) {
        RegisterColorizer(this);
        m_subclassed = true;
        return true;
    }
    return false;
}

void CommonDialogColorizer::RemoveSubclass() {
    if (m_subclassed && m_defView) {
        RemoveWindowSubclass(m_defView, &CommonDialogColorizer::SubclassProc, kSubclassId);
    }
    m_subclassed = false;
}

bool CommonDialogColorizer::HandleNotify(NMHDR* header, LRESULT* result) {
    if (!header || header->hwndFrom != m_listView || header->code != NM_CUSTOMDRAW) {
        return false;
    }

    auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(header);
    return HandleCustomDraw(draw, result);
}

bool CommonDialogColorizer::HandleCustomDraw(NMLVCUSTOMDRAW* cd, LRESULT* result) {
    if (!cd || !result) {
        return false;
    }

    switch (cd->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
            *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW;
            return true;
        case CDDS_ITEMPREPAINT:
        case CDDS_SUBITEM | CDDS_ITEMPREPAINT: {
            const int index = static_cast<int>(cd->nmcd.dwItemSpec);
            std::wstring path;
            if (!GetItemPath(index, &path)) {
                *result = CDRF_DODEFAULT;
                return true;
            }
            COLORREF colour = 0;
            if (NameColorProvider::Instance().TryGetColorForPath(path, &colour)) {
                if ((cd->nmcd.uItemState & (CDIS_SELECTED | CDIS_HOT)) == 0) {
                    cd->clrText = colour;
                    *result = CDRF_NEWFONT;
                    return true;
                }
            }
            *result = CDRF_DODEFAULT;
            return true;
        }
        default:
            break;
    }

    return false;
}

bool CommonDialogColorizer::GetItemPath(int index, std::wstring* path) const {
    if (!path) {
        return false;
    }

    if (m_folderView) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (SUCCEEDED(m_folderView->GetItem(index, IID_PPV_ARGS(&item))) && item) {
            std::wstring resolved = GetShellItemDisplayName(item.Get());
            if (!resolved.empty()) {
                *path = std::move(resolved);
                return true;
            }
        }
    }

    if (!m_listView || m_currentFolder.empty()) {
        return false;
    }

    std::wstring text(256, L'\0');
    while (true) {
        const int copied = ListView_GetItemTextW(m_listView, index, 0, text.data(),
                                                 static_cast<int>(text.size()));
        if (copied < 0) {
            return false;
        }
        if (copied < static_cast<int>(text.size()) - 1) {
            text.resize(copied);
            break;
        }
        text.resize(text.size() * 2);
    }

    if (text.empty()) {
        return false;
    }

    *path = m_currentFolder;
    if (!path->empty() && path->back() != L'\\') {
        path->push_back(L'\\');
    }
    path->append(text);
    return true;
}

HWND CommonDialogColorizer::FindDescendantByClass(HWND root, const wchar_t* className) {
    if (!root || !className) {
        return nullptr;
    }

    HWND child = FindWindowExW(root, nullptr, className, nullptr);
    if (child) {
        return child;
    }

    struct EnumData {
        const wchar_t* target;
        HWND result = nullptr;
    } data{className, nullptr};

    EnumChildWindows(root,
                     [](HWND hwnd, LPARAM lparam) -> BOOL {
                         auto* d = reinterpret_cast<EnumData*>(lparam);
                         if (!d || d->result) {
                             return FALSE;
                         }
                         wchar_t buffer[64] = {};
                         if (GetClassNameW(hwnd, buffer, ARRAYSIZE(buffer)) &&
                             lstrcmpiW(buffer, d->target) == 0) {
                             d->result = hwnd;
                             return FALSE;
                         }
                         return TRUE;
                     },
                     reinterpret_cast<LPARAM>(&data));

    return data.result;
}

LRESULT CALLBACK CommonDialogColorizer::SubclassProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam,
                                                     UINT_PTR id, DWORD_PTR refData) {
    auto* self = reinterpret_cast<CommonDialogColorizer*>(refData);
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
            UnregisterColorizer(self);
            self->m_subclassed = false;
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

}  // namespace shelltabs

