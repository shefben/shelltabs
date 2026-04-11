/*
 * ShellTabs - Folder Background Hooks
 * Folder background rendering via window subclassing on DirectUIHWND / SHELLDLL_DefView.
 */

#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <commctrl.h>
#include <objidl.h>
#include <gdiplus.h>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <mutex>
#include <memory>

#include "OptionsStore.h"

namespace shelltabs {

// GDI-compatible bitmap wrapper for efficient AlphaBlend painting.
class BackgroundBitmap {
public:
    explicit BackgroundBitmap(const std::wstring& path);
    ~BackgroundBitmap();

    bool    IsValid()    const { return m_bitmap != nullptr; }
    HBITMAP GetBitmap()  const { return m_bitmap; }
    SIZE    GetSize()    const { return m_size; }

private:
    HBITMAP m_bitmap = nullptr;
    SIZE    m_size   = {0, 0};
};

// Main hook management singleton.
// Supports per-frame (per-Explorer-window) folder-specific backgrounds with a universal fallback.
// Uses SetWindowSubclass on SHELLDLL_DefView and DirectUIHWND to intercept WM_ERASEBKGND.
class FolderBackgroundHooks {
public:
    static FolderBackgroundHooks& Instance();

    bool Initialize();
    void Shutdown();
    bool IsInitialized() const { return m_initialized; }

    // Called from CExplorerBHO when options are applied.
    void SetImages(bool enabled,
                   BackgroundPositionMode positionMode,
                   BYTE opacity,
                   const std::wstring& universalImagePath,
                   const std::unordered_map<std::wstring, std::wstring>& folderImagePaths);

    // Called from CExplorerBHO on each folder navigation.
    void SetFrameFolderPath(HWND explorerFrame, const std::wstring& normalizedFolderPath);

    // Called from CExplorerBHO on disconnect.
    void ClearFrameFolderPath(HWND explorerFrame);

private:
    FolderBackgroundHooks() = default;
    ~FolderBackgroundHooks();
    FolderBackgroundHooks(const FolderBackgroundHooks&) = delete;
    FolderBackgroundHooks& operator=(const FolderBackgroundHooks&) = delete;

    // Window subclassing
    static LRESULT CALLBACK SubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                         UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    void TrySubclassFrame(HWND explorerFrame);
    void SubclassHwnd(HWND hwnd, HWND explorerFrame, UINT_PTR id);
    void UnsubclassAll();

    // Hook implementations
    static HWND WINAPI HookedCreateWindowExW(
        DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
        DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
        HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam);

    // Drawing
    // Returns a live shared_ptr so the caller can draw safely after releasing m_mutex.
    std::shared_ptr<const BackgroundBitmap> ResolveBitmapForFrame(HWND explorerFrame) const;
    void CalculateImagePosition(const SIZE& wndSize, const SIZE& imgSize,
                                BackgroundPositionMode mode,
                                POINT& pos, SIZE& dstSize) const;

    // Overlay window management
    void CreateOverlayForDUI(HWND duiHwnd, HWND defViewParent);
    void DestroyOverlayForDUI(HWND duiHwnd);
    void UpdateOverlayContent(HWND duiHwnd, HWND explorerFrame);
    void RepositionOverlay(HWND duiHwnd);
    void DestroyAllOverlays();

    // State
    bool m_initialized = false;
    mutable std::mutex m_mutex;

    bool                  m_enabled      = false;
    BackgroundPositionMode m_positionMode = BackgroundPositionMode::kBottomRight;
    BYTE                  m_opacity      = 200;

    std::unordered_map<std::wstring, std::shared_ptr<BackgroundBitmap>> m_folderBitmaps;
    std::shared_ptr<BackgroundBitmap> m_universalBitmap;

    // Per-frame current folder path (keyed by CabinetWClass HWND)
    std::unordered_map<HWND, std::wstring> m_frameFolderPaths;

    // Currently subclassed windows
    std::unordered_set<HWND> m_subclassedWindows;

    // Overlay windows: DirectUIHWND → overlay HWND
    std::unordered_map<HWND, HWND> m_overlayWindows;
    // Track last overlay size per DirectUIHWND to skip redundant updates
    std::unordered_map<HWND, SIZE> m_lastOverlaySize;

    static ATOM s_overlayClass;
    static inline decltype(&CreateWindowExW) s_originalCreateWindowExW = nullptr;
};

bool InitializeFolderBackgroundHooks();
void ShutdownFolderBackgroundHooks();

}  // namespace shelltabs
