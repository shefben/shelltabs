#pragma once
#include <windows.h>
#include <wrl/client.h>
#include <memory>
#include "ExplorerPane.h"

namespace shelltabs {

	class SplitHost {
	public:
		static SplitHost* CreateAndAttach(HWND contentParent);
		static SplitHost* FromHwnd(HWND hwnd);
		static void DestroyIfExistsOn(HWND contentParent);

		HWND Hwnd() const { return m_hwnd; }
		void SetFolders(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right);
		void Swap();

	private:
		SplitHost() = default;
		~SplitHost() = default;

		static LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
		LRESULT OnCreate(HWND hwnd);
		void OnSize();
		void OnLButtonDown(int x);
		void OnMouseMove(int x);
		void OnLButtonUp();
		void LayoutChildren();

		HWND m_hwnd = nullptr;
		HWND m_splitter = nullptr;
		int m_splitX = 400;      // default position
		bool m_drag = false;

		ExplorerPane m_left;
		ExplorerPane m_right;
	};

} // namespace shelltabs
