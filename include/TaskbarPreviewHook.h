#pragma once

#include <windows.h>
#include <oleacc.h>
#include <memory>
#include <string>
#include <vector>

#include "TaskbarTabProvider.h"

namespace shelltabs {

class TaskbarTabListPopup;

class TaskbarPreviewHook {
public:
    static TaskbarPreviewHook& Instance();
    void Initialize(std::unique_ptr<ITaskbarTabProvider> provider);
    void Shutdown();

private:
    TaskbarPreviewHook() = default;
    ~TaskbarPreviewHook();
    TaskbarPreviewHook(const TaskbarPreviewHook&) = delete;
    TaskbarPreviewHook& operator=(const TaskbarPreviewHook&) = delete;

    static void CALLBACK WinEventProc(HWINEVENTHOOK hook, DWORD event, HWND hwnd,
                                       LONG idObject, LONG idChild,
                                       DWORD eventThread, DWORD eventTime);

    static void CALLBACK PollTimerProc(HWND hwnd, UINT msg, UINT_PTR id, DWORD time);
    static void CALLBACK HideDebounceProc(HWND hwnd, UINT msg, UINT_PTR id, DWORD time);

    void OnPreviewPopupShown(HWND thumbnailWnd);
    void OnPreviewPopupHidden(HWND thumbnailWnd);
    void OnPollTick();
    void OnHideDebounce();

    struct ThumbnailInfo {
        HWND explorerHwnd = nullptr;
        RECT thumbnailRect{};
        std::wstring windowTitle;
    };

    ThumbnailInfo HitTestThumbnail(HWND thumbnailWnd, POINT screenPt);
    HWND MatchTitleToExplorerWindow(const std::wstring& title);

    std::unique_ptr<ITaskbarTabProvider> m_provider;
    std::unique_ptr<TaskbarTabListPopup> m_popup;
    HWINEVENTHOOK m_eventHook = nullptr;
    HWND m_activeThumbnailWnd = nullptr;
    HWND m_currentExplorerHwnd = nullptr;
    UINT_PTR m_pollTimer = 0;
    UINT_PTR m_hideTimer = 0;
    bool m_initialized = false;

    struct ExplorerWindowEntry {
        HWND hwnd;
        std::wstring title;
    };
    std::vector<ExplorerWindowEntry> m_explorerWindows;
    void RefreshExplorerWindowList();

    static constexpr UINT_PTR kPollTimerId = 9901;
    static constexpr UINT_PTR kHideTimerId = 9902;
    static constexpr DWORD kPollIntervalMs = 50;
    static constexpr DWORD kHideDebounceMs = 150;
};

}  // namespace shelltabs
