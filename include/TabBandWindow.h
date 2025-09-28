#pragma once

#include <CommCtrl.h>
#include <uxtheme.h>

#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "TabManager.h"

namespace shelltabs {

class TabBand;

constexpr UINT WM_SHELLTABS_CLOSETAB = WM_APP + 42;
constexpr UINT WM_SHELLTABS_DEFER_NAVIGATE = WM_APP + 43;
constexpr UINT WM_SHELLTABS_REFRESH_COLORIZER = WM_APP + 44;
constexpr UINT WM_SHELLTABS_REFRESH_GIT_STATUS = WM_APP + 45;
constexpr UINT WM_SHELLTABS_ENABLE_GIT_STATUS = WM_APP + 46;

class TabBandWindow {
public:
    explicit TabBandWindow(TabBand* owner);
    ~TabBandWindow();

    HWND Create(HWND parent);
    void Destroy();

    HWND GetHwnd() const noexcept { return m_hwnd; }

    void Show(bool show);
    void SetTabs(const std::vector<TabViewItem>& items);
    bool HasFocus() const;
    void FocusTab();

private:
    HWND m_hwnd = nullptr;
    HWND m_toolbar = nullptr;
    TabBand* m_owner = nullptr;
    std::vector<TabViewItem> m_tabData;
    HIMAGELIST m_imageList = nullptr;
    std::unordered_map<int, TabLocation> m_commandMap;
    std::unordered_map<int, size_t> m_commandToIndex;
    int m_nextCommandId = 41000;

    void EnsureToolbar();
    void DestroyToolbar();
    void RebuildToolbar();
    void ClearToolbar();
    void ClearImageList();
    int AppendImage(HICON icon);
    void UpdateCheckedState();
    void HandleToolbarCommand(int commandId);
    void HandleContextMenu(int commandId, const POINT& screenPt);
    void HandleMiddleClick(int commandId);
    void HandleLButtonDown(int commandId);
    void HandleTooltipRequest(NMTTDISPINFOW* info);
    void RelayFocusToToolbar();
    TabLocation LocationForCommand(int commandId) const;
    const TabViewItem* ItemForCommand(int commandId) const;

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                    DWORD_PTR refData);
    static LRESULT CALLBACK ToolbarWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR id,
                                           DWORD_PTR refData);
};

}  // namespace shelltabs
