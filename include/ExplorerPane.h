#pragma once
#include <shobjidl.h>
#include <wrl/client.h>

namespace shelltabs {

	class ExplorerPane {
	public:
		ExplorerPane();
		~ExplorerPane();

		HRESULT Create(HWND parent, const RECT& rc);
		void Destroy();
		HWND GetHwnd() const;

		HRESULT NavigateToPIDL(PCIDLIST_ABSOLUTE pidl);
		HRESULT NavigateToShellItem(IShellItem* item);
		HRESULT SetRect(const RECT& rc);

	private:
		Microsoft::WRL::ComPtr<IExplorerBrowser> m_browser;
		HWND m_hwnd = nullptr;
	};

} // namespace shelltabs
