#include "NamespaceTreeColorizer.h"

#include "NameColorProvider.h"
#include "FileColorOverrides.h"
#include "Utilities.h"

#include <ShlObj.h>
#include <shlwapi.h>
#include <cwchar>

#include <string>
#include <type_traits>
#include <utility>

#pragma comment(lib, "Shlwapi.lib")

using Microsoft::WRL::ComPtr;

namespace {

template <typename T, typename = void>
struct HasNmcd : std::false_type {};

template <typename T>
struct HasNmcd<T, std::void_t<decltype(std::declval<T&>().nmcd)>> : std::true_type {};

template <typename T, typename = void>
struct HasDwItemSpec : std::false_type {};

template <typename T>
struct HasDwItemSpec<T, std::void_t<decltype(std::declval<T&>().dwItemSpec)>> : std::true_type {};

template <typename T, typename = void>
struct HasUItemState : std::false_type {};

template <typename T>
struct HasUItemState<T, std::void_t<decltype(std::declval<T&>().uItemState)>> : std::true_type {};

template <typename T, typename = void>
struct HasClrTextBk : std::false_type {};

template <typename T>
struct HasClrTextBk<T, std::void_t<decltype(std::declval<T&>().clrTextBk)>> : std::true_type {};

template <typename T, typename = void>
struct HasHFont : std::false_type {};

template <typename T>
struct HasHFont<T, std::void_t<decltype(std::declval<T&>().hFont)>> : std::true_type {};

template <typename T, typename = void>
struct HasHfont : std::false_type {};

template <typename T>
struct HasHfont<T, std::void_t<decltype(std::declval<T&>().hfont)>> : std::true_type {};

template <typename T>
DWORD_PTR ExtractItemSpec(const T* info) {
    if constexpr (HasNmcd<T>::value) {
        return info->nmcd.dwItemSpec;
    } else if constexpr (HasDwItemSpec<T>::value) {
        return info->dwItemSpec;
    }
    if (info && info->psi) {
        return reinterpret_cast<DWORD_PTR>(info->psi);
    }
    return reinterpret_cast<DWORD_PTR>(info);
}

template <typename T>
UINT ExtractItemState(const T* info) {
    if constexpr (HasNmcd<T>::value) {
        return info->nmcd.uItemState;
    } else if constexpr (HasUItemState<T>::value) {
        return info->uItemState;
    }
    return 0;
}

template <typename T>
bool ApplyBackground(T* info, COLORREF color) {
    if constexpr (HasClrTextBk<T>::value) {
        info->clrTextBk = color;
        return true;
    }
    return false;
}

template <typename T>
bool ApplyFont(T* info, HFONT font) {
    if constexpr (HasHFont<T>::value) {
        info->hFont = font;
        return true;
    } else if constexpr (HasHfont<T>::value) {
        info->hfont = font;
        return true;
    }
    return false;
}

}  // namespace

namespace shelltabs {

NamespaceTreeColorizer::NamespaceTreeColorizer() = default;

NamespaceTreeColorizer::~NamespaceTreeColorizer() { Detach(); }

bool NamespaceTreeColorizer::Attach(ComPtr<IServiceProvider> serviceProvider) {
    Detach();
    if (!ResolveNSTC(serviceProvider)) {
        return false;
    }

    ComPtr<IUnknown> sink;
    HRESULT hr = QueryInterface(IID_PPV_ARGS(&sink));
    if (FAILED(hr) || !sink) {
        return false;
    }

    hr = nstc_->TreeAdvise(sink.Get(), &cookie_);
    return SUCCEEDED(hr);
}

void NamespaceTreeColorizer::Detach() {
    if (nstc_ && cookie_ != 0) {
        nstc_->TreeUnadvise(cookie_);
    }
    cookie_ = 0;
    nstc_.Reset();
    pending_tree_rename_.clear();

    for (HFONT font : owned_fonts_) {
        if (font) {
            DeleteObject(font);
        }
    }
    owned_fonts_.clear();
    pending_paints_.clear();
}

bool NamespaceTreeColorizer::ResolveNSTC(ComPtr<IServiceProvider> serviceProvider) {
    if (!serviceProvider) {
        return false;
    }

    ComPtr<IShellBrowser> browser;
    if (FAILED(serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser))) || !browser) {
        return false;
    }

    ComPtr<IServiceProvider> browserProvider;
    if (FAILED(browser.As(&browserProvider)) || !browserProvider) {
        return false;
    }

    if (FAILED(browserProvider->QueryService(__uuidof(INameSpaceTreeControl), IID_PPV_ARGS(&nstc_))) || !nstc_) {
        return false;
    }

    return true;
}

bool NamespaceTreeColorizer::ItemPathFromShellItem(IShellItem* item, std::wstring* path) const {
    if (!item || !path) {
        return false;
    }

    return TryGetFileSystemPath(item, path);
}

// IUnknown
IFACEMETHODIMP NamespaceTreeColorizer::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == __uuidof(INameSpaceTreeControlCustomDraw)) {
        *object = static_cast<INameSpaceTreeControlCustomDraw*>(this);
        AddRef();
        return S_OK;
    }
    if (riid == __uuidof(INameSpaceTreeControlEvents)) {
        *object = static_cast<INameSpaceTreeControlEvents*>(this);
        AddRef();
        return S_OK;
    }
    *object = nullptr;
    return E_NOINTERFACE;
}

IFACEMETHODIMP_(ULONG) NamespaceTreeColorizer::AddRef() { return ++ref_; }

IFACEMETHODIMP_(ULONG) NamespaceTreeColorizer::Release() {
    ULONG remaining = --ref_;
    if (remaining == 0) {
        delete this;
    }
    return remaining;
}

// INameSpaceTreeControlCustomDraw
IFACEMETHODIMP NamespaceTreeColorizer::PrePaint(HDC, RECT*, LRESULT* result) {
    if (result) {
        *result = CDRF_NOTIFYITEMDRAW;
    }
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::PostPaint(HDC, RECT*) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::ItemPrePaint(HDC, RECT*, NSTCCUSTOMDRAW* drawInfo,
                                                    COLORREF* textColor, COLORREF* backgroundColor,
                                                    LRESULT* result) {
    if (!drawInfo || !result) {
        return S_OK;
    }

    const DWORD_PTR key = ExtractItemSpec(drawInfo);
    pending_paints_.erase(key);

    std::wstring path;
    if (!drawInfo->psi || !ItemPathFromShellItem(drawInfo->psi, &path)) {
        *result = CDRF_DODEFAULT;
        return S_OK;
    }

    const auto appearance = NameColorProvider::Instance().GetAppearanceForPath(path);
    if (!appearance.HasOverrides() || !appearance.AllowsForState(ExtractItemState(drawInfo))) {
        *result = CDRF_DODEFAULT;
        return S_OK;
    }

    bool applied = false;

    if (appearance.textColor.has_value() && textColor) {
        *textColor = *appearance.textColor;
        applied = true;
    }

    if (appearance.backgroundColor.has_value()) {
        if (ApplyBackground(drawInfo, *appearance.backgroundColor)) {
            applied = true;
        }
        if (backgroundColor) {
            *backgroundColor = *appearance.backgroundColor;
            applied = true;
        }
        pending_paints_[key] = PendingItemPaint{true, *appearance.backgroundColor};
        applied = true;
    }

    if (appearance.font) {
        applied = ApplyFont(drawInfo, appearance.font) || applied;
        if (appearance.ownsFont) {
            owned_fonts_.insert(appearance.font);
        }
    }

    if (applied) {
        *result = CDRF_NEWFONT | CDRF_NOTIFYPOSTPAINT;
    } else {
        *result = CDRF_DODEFAULT;
    }

    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::ItemPostPaint(HDC hdc, RECT* rect, NSTCCUSTOMDRAW* drawInfo) {
    if (!drawInfo) {
        return S_OK;
    }

    const DWORD_PTR key = ExtractItemSpec(drawInfo);
    auto it = pending_paints_.find(key);
    if (it == pending_paints_.end()) {
        return S_OK;
    }

    if (it->second.fillBackground && rect) {
        const RECT fillRect = *rect;
        HBRUSH brush = CreateSolidBrush(it->second.background);
        if (brush) {
            FillRect(hdc, &fillRect, brush);
            DeleteObject(brush);
        }
    }

    pending_paints_.erase(it);
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnItemClick(IShellItem*, NSTCEHITTEST, NSTCECLICKTYPE) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnPropertyItemCommit(IShellItem*) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnItemStateChanging(IShellItem*, NSTCITEMSTATE, NSTCITEMSTATE) {
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnItemStateChanged(IShellItem*, NSTCITEMSTATE, NSTCITEMSTATE) {
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnSelectionChanged(IShellItemArray*) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnKeyboardInput(UINT, WPARAM, LPARAM) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnBeforeExpand(IShellItem*) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnAfterExpand(IShellItem*) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnBeginLabelEdit(IShellItem* psi) {
    pending_tree_rename_.clear();
    if (psi) {
        ItemPathFromShellItem(psi, &pending_tree_rename_);
    }
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnEndLabelEdit(IShellItem* psi) {
    std::wstring newPath;
    if (psi) {
        ItemPathFromShellItem(psi, &newPath);
    }

    if (!pending_tree_rename_.empty() && !newPath.empty() &&
        _wcsicmp(pending_tree_rename_.c_str(), newPath.c_str()) != 0) {
        FileColorOverrides::Instance().TransferColor(pending_tree_rename_, newPath);
    }

    pending_tree_rename_.clear();
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnGetToolTip(IShellItem*, LPWSTR tip, int cch) {
    if (tip && cch > 0) {
        tip[0] = L'\0';
    }
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnBeforeItemDelete(IShellItem*) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnItemAdded(IShellItem*, BOOL) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnItemDeleted(IShellItem*, BOOL) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnBeforeContextMenu(IShellItem*, REFIID, void**) { return S_OK; }

IFACEMETHODIMP NamespaceTreeColorizer::OnAfterContextMenu(IShellItem*, IContextMenu*, REFIID, void**) {
    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::OnBeforeStateImageChange(IShellItem*) { return S_OK; }

}  // namespace shelltabs

