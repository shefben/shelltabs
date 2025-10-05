#include "NamespaceTreeColorizer.h"
#include "FileColorOverrides.h"
#include <shlwapi.h>
#include <ShlObj.h>

#pragma comment(lib, "Shlwapi.lib")

using Microsoft::WRL::ComPtr;

namespace shelltabs {

	NamespaceTreeColorizer::NamespaceTreeColorizer() {}
	NamespaceTreeColorizer::~NamespaceTreeColorizer() { Detach(); }

	bool NamespaceTreeColorizer::Attach(ComPtr<IServiceProvider> sp) {
		Detach();
		if (!ResolveNSTC(sp)) return false;

		// Get IUnknown for this sink (QI returns HRESULT, the OUT param gets the pointer)
		Microsoft::WRL::ComPtr<IUnknown> unk;
		HRESULT hr = this->QueryInterface(IID_PPV_ARGS(&unk));
		if (FAILED(hr)) return false;

		// Advise the sink
		hr = nstc_->TreeAdvise(unk.Get(), &cookie_);
		return SUCCEEDED(hr);
	}

	void NamespaceTreeColorizer::Detach() {
		if (nstc_ && cookie_) {
			nstc_->TreeUnadvise(cookie_);
			cookie_ = 0;
		}
		nstc_.Reset();
	}


	bool NamespaceTreeColorizer::ResolveNSTC(ComPtr<IServiceProvider> sp) {
		if (!sp) return false;
		ComPtr<IShellBrowser> browser;
		if (FAILED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser)))) return false;

		ComPtr<IServiceProvider> bsp;
		if (FAILED(browser.As(&bsp))) return false;
		if (FAILED(bsp->QueryService(__uuidof(INameSpaceTreeControl), IID_PPV_ARGS(&nstc_)))) {
			return false;
		}
		return nstc_ != nullptr;
	}

	bool NamespaceTreeColorizer::ItemPathFromShellItem(IShellItem* psi, std::wstring* out) const {
		if (!psi || !out) return false;
		PWSTR p = nullptr;
		if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &p)) && p) {
			*out = p;
			CoTaskMemFree(p);
			return true;
		}
		return false;
	}


	// IUnknown
	IFACEMETHODIMP NamespaceTreeColorizer::QueryInterface(REFIID riid, void** ppv) {
		if (!ppv) return E_POINTER;
		if (riid == IID_IUnknown || riid == __uuidof(INameSpaceTreeControlCustomDraw)) {
			*ppv = static_cast<INameSpaceTreeControlCustomDraw*>(this);
			AddRef(); return S_OK;
		}
		*ppv = nullptr;
		return E_NOINTERFACE;
	}
	IFACEMETHODIMP_(ULONG) NamespaceTreeColorizer::AddRef() { return ++ref_; }

	IFACEMETHODIMP_(ULONG) NamespaceTreeColorizer::Release() {
		ULONG r = --ref_; if (!r) delete this; return r;
	}

	// INameSpaceTreeControlCustomDraw
	IFACEMETHODIMP NamespaceTreeColorizer::PrePaint(HDC, RECT*, LRESULT* lr) {
		if (lr) *lr = CDRF_NOTIFYITEMDRAW;
		return S_OK;
	}

	IFACEMETHODIMP NamespaceTreeColorizer::PostPaint(HDC, RECT*) {
		return S_OK;
	}

	IFACEMETHODIMP NamespaceTreeColorizer::ItemPrePaint(
		HDC, RECT*, NSTCCUSTOMDRAW* cd, COLORREF* pclrText, COLORREF* pclrTextBk, LRESULT* lr) {
		if (!cd || !lr) return S_OK;
		*lr = CDRF_DODEFAULT;

		std::wstring path;
		if (cd->psi && ItemPathFromShellItem(cd->psi, &path)) {
			COLORREF col;
			if (FileColorOverrides::Instance().TryGetColor(path, &col)) {
				if (pclrText)   *pclrText = col;   // set text color
				// leave background alone unless you want to force it:
				// if (pclrTextBk) *pclrTextBk = RGB(...);
				*lr = CDRF_NEWFONT;
			}
		}
		return S_OK;
	}

	IFACEMETHODIMP NamespaceTreeColorizer::ItemPostPaint(HDC, RECT*, NSTCCUSTOMDRAW*) {
		return S_OK;
	}


} // namespace shelltabs
