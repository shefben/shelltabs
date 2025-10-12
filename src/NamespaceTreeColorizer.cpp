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

    *result = CDRF_DODEFAULT;

    std::wstring path;
    if (drawInfo->psi && ItemPathFromShellItem(drawInfo->psi, &path)) {
        COLORREF colour = 0;
        if (NameColorProvider::Instance().TryGetColorForPath(path, &colour)) {
            if (textColor) {
                *textColor = colour;
            }
            *result = CDRF_NEWFONT;
        }
    }

    return S_OK;
}

IFACEMETHODIMP NamespaceTreeColorizer::ItemPostPaint(HDC, RECT*, NSTCCUSTOMDRAW*) { return S_OK; }

}  // namespace shelltabs

