#include "ExplorerPane.h"
#include <shlobj.h>

#ifndef RETURN_IF_FAILED
#define RETURN_IF_FAILED(hrcall) do { HRESULT _hr=(hrcall); if (FAILED(_hr)) return _hr; } while(0)
#endif

using Microsoft::WRL::ComPtr;

namespace shelltabs {

	ExplorerPane::ExplorerPane() = default;
	ExplorerPane::~ExplorerPane() { Destroy(); }

	HRESULT ExplorerPane::Create(HWND parent, const RECT& rc) {
		Destroy();
		FOLDERSETTINGS fs{};
		fs.ViewMode = FVM_DETAILS;
		fs.fFlags = FWF_SHOWSELALWAYS;

		m_browser = nullptr;
		RETURN_IF_FAILED(CoCreateInstance(CLSID_ExplorerBrowser, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_browser)));

		RETURN_IF_FAILED(m_browser->Initialize(parent, const_cast<RECT*>(&rc), &fs));

		// >>> FIX: IExplorerBrowser doesn't expose GetWindow; use IOleWindow <<<
// REPLACE the "get hwnd" block in Create(...) with:
		Microsoft::WRL::ComPtr<IOleWindow> ole;
		RETURN_IF_FAILED(m_browser.As(&ole));
		RETURN_IF_FAILED(ole->GetWindow(&m_hwnd));

		// <<< FIX >>>

		return S_OK;
	}

	void ExplorerPane::Destroy() {
		if (m_browser) {
			m_browser->Destroy();
			m_browser.Reset();
		}
		m_hwnd = nullptr;
	}

	HWND ExplorerPane::GetHwnd() const { return m_hwnd; }

	HRESULT ExplorerPane::SetRect(const RECT& rc) {
		return m_browser ? m_browser->SetRect(nullptr, rc) : E_FAIL;
	}


	HRESULT ExplorerPane::NavigateToPIDL(PCIDLIST_ABSOLUTE pidl) {
		if (!m_browser) return E_FAIL;
		ComPtr<IShellItem> item;
		if (FAILED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&item)))) return E_FAIL;
		return m_browser->BrowseToObject(item.Get(), 0);
	}

	HRESULT ExplorerPane::NavigateToShellItem(IShellItem* item) {
		return m_browser ? m_browser->BrowseToObject(item, 0) : E_FAIL;
	}

} // namespace shelltabs
