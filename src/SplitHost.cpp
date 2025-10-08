#include "SplitHost.h"
#include <commctrl.h>
#include <windowsx.h>   // >>> ADD for GET_X_LPARAM <<<
#include <algorithm>    // >>> ADD for std::min/std::max <<<


using namespace shelltabs;

namespace {
	const wchar_t* kSplitHostClass = L"ShellTabs.SplitHost";
	const int kSplitterWidth = 6;
}

SplitHost* SplitHost::CreateAndAttach(HWND parent) {
	static bool cls = false;
	if (!cls) {
		WNDCLASS wc{};
		wc.lpfnWndProc = SplitHost::WndProc;
		wc.hInstance = (HINSTANCE)GetModuleHandle(nullptr);
		wc.lpszClassName = kSplitHostClass;
		wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
		RegisterClass(&wc);
		cls = true;
	}

	RECT rc{}; GetClientRect(parent, &rc);
	HWND hwnd = CreateWindowExW(0, kSplitHostClass, L"", WS_CHILD | WS_VISIBLE,
		0, 0, rc.right, rc.bottom, parent, nullptr, (HINSTANCE)GetModuleHandle(nullptr), nullptr);
	if (!hwnd) return nullptr;
	return FromHwnd(hwnd);
}

SplitHost* SplitHost::FromHwnd(HWND hwnd) {
	return reinterpret_cast<SplitHost*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
}

void SplitHost::DestroyIfExistsOn(HWND parent) {
	for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT)) {
		wchar_t cls[128]{}; GetClassNameW(child, cls, 128);
		if (lstrcmpiW(cls, kSplitHostClass) == 0) {
			DestroyWindow(child);
			return;
		}
	}
}

LRESULT CALLBACK SplitHost::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
	if (msg == WM_NCCREATE) {
		auto self = new SplitHost();
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)self);
		self->m_hwnd = hwnd;
		return DefWindowProc(hwnd, msg, wp, lp);
	}
	auto self = FromHwnd(hwnd);
	if (!self) return DefWindowProc(hwnd, msg, wp, lp);

	switch (msg) {
	case WM_CREATE: return self->OnCreate(hwnd);
	case WM_SIZE:   self->OnSize(); break;
	case WM_LBUTTONDOWN: self->OnLButtonDown(GET_X_LPARAM(lp)); break;
	case WM_MOUSEMOVE:   self->OnMouseMove(GET_X_LPARAM(lp)); break;
	case WM_LBUTTONUP:   self->OnLButtonUp(); break;
	case WM_DESTROY:
		delete self;
		break;
	}
	return DefWindowProc(hwnd, msg, wp, lp);
}

LRESULT SplitHost::OnCreate(HWND hwnd) {
	RECT rc{}; GetClientRect(hwnd, &rc);
	m_left.Create(hwnd, rc);
	m_right.Create(hwnd, rc);
	m_splitter = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_NOTIFY, 0, 0, kSplitterWidth, rc.bottom, hwnd, nullptr, nullptr, nullptr);
	SetWindowPos(m_splitter, HWND_TOP, m_splitX, 0, kSplitterWidth, rc.bottom, SWP_NOACTIVATE);
	LayoutChildren();
	return 0;
}

void SplitHost::OnSize() { LayoutChildren(); }

void SplitHost::LayoutChildren() {
	RECT rc{}; GetClientRect(m_hwnd, &rc);
	int width = rc.right - rc.left;
	m_splitX = std::max(100, std::min(m_splitX, width - 100));  // was max/min -> use std::

	RECT leftRc{ 0, 0, m_splitX, rc.bottom };
	RECT rightRc{ m_splitX + kSplitterWidth, 0, rc.right, rc.bottom };
	m_left.SetRect(leftRc);
	m_right.SetRect(rightRc);
	SetWindowPos(m_splitter, HWND_TOP, m_splitX, 0, kSplitterWidth, rc.bottom, SWP_NOACTIVATE | SWP_SHOWWINDOW);
}

void SplitHost::OnLButtonDown(int x) {
	RECT sr{}; GetWindowRect(m_splitter, &sr);
	POINT pt{ x, 0 }; ClientToScreen(m_hwnd, &pt);
	if (pt.x >= sr.left && pt.x <= sr.right) {
		SetCapture(m_hwnd);
		m_drag = true;
	}
}
void SplitHost::OnMouseMove(int x) {
	if (!m_drag) return;
	m_splitX = x;
	LayoutChildren();
}
void SplitHost::OnLButtonUp() {
	if (m_drag) { m_drag = false; ReleaseCapture(); }
}

// Implement the simple folder setters and swap:
void SplitHost::SetFolders(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right) {
	if (left)  m_left.NavigateToPIDL(left);
	if (right) m_right.NavigateToPIDL(right);
}
void SplitHost::Swap() {
	int w; RECT rc{}; GetClientRect(m_hwnd, &rc); w = rc.right;
	m_splitX = w - m_splitX - kSplitterWidth;
	LayoutChildren();
}