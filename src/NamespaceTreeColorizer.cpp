#include "NamespaceTreeColorizer.h"

#include "NameColorProvider.h"

#include <ShlObj.h>
#include <shlwapi.h>

#include <string>

#pragma comment(lib, "Shlwapi.lib")

using Microsoft::WRL::ComPtr;

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
        cookie_ = 0;
    }
    nstc_.Reset();

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

    PWSTR buffer = nullptr;
    HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, &buffer);
    if (FAILED(hr) || !buffer) {
        hr = item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &buffer);
    }
    if (FAILED(hr) || !buffer) {
        return false;
    }

    path->assign(buffer);
    CoTaskMemFree(buffer);
    return true;
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
                                                    COLORREF* textColor, COLORREF*, LRESULT* result) {
    if (!drawInfo || !result) {
        return S_OK;
    }

    const DWORD_PTR key = drawInfo->nmcd.dwItemSpec;
    pending_paints_.erase(key);

    std::wstring path;
    if (!drawInfo->psi || !ItemPathFromShellItem(drawInfo->psi, &path)) {
        *result = CDRF_DODEFAULT;
        return S_OK;
    }

    const auto appearance = NameColorProvider::Instance().GetAppearanceForPath(path);
    if (!appearance.HasOverrides() || !appearance.AllowsForState(drawInfo->nmcd.uItemState)) {
        *result = CDRF_DODEFAULT;
        return S_OK;
    }

    bool applied = false;

    if (appearance.textColor.has_value() && textColor) {
        *textColor = *appearance.textColor;
        applied = true;
    }

    if (appearance.backgroundColor.has_value()) {
        drawInfo->clrTextBk = *appearance.backgroundColor;
        pending_paints_[key] = PendingItemPaint{true, *appearance.backgroundColor};
        applied = true;
    }

    if (appearance.font) {
        drawInfo->hfont = appearance.font;
        applied = true;
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

    const DWORD_PTR key = drawInfo->nmcd.dwItemSpec;
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

}  // namespace shelltabs

