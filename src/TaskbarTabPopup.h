#pragma once

#include <windows.h>

#include <vector>

#include "PreviewOverlay.h"
#include "TabManager.h"

namespace shelltabs {

class TabBand;
class TabBandWindow;

class TaskbarTabPopup {
public:
    explicit TaskbarTabPopup(TabBand* owner);
    ~TaskbarTabPopup();

    TaskbarTabPopup(const TaskbarTabPopup&) = delete;
    TaskbarTabPopup& operator=(const TaskbarTabPopup&) = delete;

    void Show(const POINT& anchor, HWND ownerWindow, TabBandWindow* tabWindow);
    void Hide();
    void Destroy();
    bool IsVisible() const noexcept { return m_visible; }

private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static ATOM EnsurePopupWindowClass();

    void EnsureWindow(HWND ownerWindow);
    void InitializeWindow(HWND hwnd);
    void Populate(TabBandWindow* tabWindow);
    void ActivateIndex(int index);
    void HandleNotify(NMHDR* header);
    void HideInternal();
    void OnHotItemChanged(int index, const POINT& ptClient);
    void OnItemChanged(const NMLISTVIEW& info);
    void HandleHover();
    void HandleTimer(UINT_PTR timerId);
    void EnsureMouseTracking();
    void HidePreview();
    void ShowPreviewForIndex(int index);
    void StopPreviewTimer();

    TabBand* m_owner = nullptr;
    HWND m_hwnd = nullptr;
    HWND m_listView = nullptr;
    HIMAGELIST m_imageList = nullptr;
    std::vector<TabViewItem> m_items;
    bool m_visible = false;
    bool m_windowInitialized = false;
    int m_lastColumnWidth = 280;
    PreviewOverlay m_previewOverlay;
    int m_hotItem = -1;
    int m_previewItem = -1;
    bool m_mouseTracking = false;
    bool m_previewTimerActive = false;
    POINT m_lastHoverPoint{};
};

}  // namespace shelltabs
