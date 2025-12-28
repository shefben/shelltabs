/*
 * ShellTabs - Folder Background Hooks
 * DirectUI-based folder background rendering system using API hooking
 *
 * Based on ExplorerBgTool by Maplespe (winmoes.com)
 * Adapted for ShellTabs by integrating with existing infrastructure
 */

#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <objidl.h>  // Required for IStream before gdiplus.h
#include <gdiplus.h>
#include <string>
#include <unordered_map>
#include <mutex>
#include <memory>

#include "OptionsStore.h"  // For BackgroundPositionMode

namespace shelltabs {

// Background configuration
struct FolderBackgroundConfig {
    bool enabled = false;
    BackgroundPositionMode positionMode = BackgroundPositionMode::kBottomRight;
    BYTE opacity = 255;  // 0-255
    std::wstring imagePath;
};

// GDI-compatible bitmap wrapper for efficient painting
class BackgroundBitmap {
public:
    BackgroundBitmap(const std::wstring& path);
    ~BackgroundBitmap();

    bool IsValid() const { return m_memDC != nullptr && m_bitmap != nullptr; }
    HDC GetMemDC() const { return m_memDC; }
    SIZE GetSize() const { return m_size; }
    Gdiplus::Bitmap* GetSource() const { return m_source.get(); }

private:
    HDC m_memDC = nullptr;
    HBITMAP m_bitmap = nullptr;
    SIZE m_size = { 0, 0 };
    std::unique_ptr<Gdiplus::Bitmap> m_source;
};

// Main hook management class (singleton)
class FolderBackgroundHooks {
public:
    static FolderBackgroundHooks& Instance();

    // Initialization and cleanup
    bool Initialize();
    void Shutdown();
    bool IsInitialized() const { return m_initialized; }

    // Configuration
    void SetConfig(const FolderBackgroundConfig& config);
    FolderBackgroundConfig GetConfig() const;

    // Image loading
    bool LoadImage(const std::wstring& path);
    void ClearImage();

    // Refresh settings (called when options change)
    void ReloadSettings();

private:
    FolderBackgroundHooks() = default;
    ~FolderBackgroundHooks();
    FolderBackgroundHooks(const FolderBackgroundHooks&) = delete;
    FolderBackgroundHooks& operator=(const FolderBackgroundHooks&) = delete;

    // Hook implementations (static for MinHook compatibility)
    static HWND WINAPI HookedCreateWindowExW(
        DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
        DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
        HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam);

    static BOOL WINAPI HookedDestroyWindow(HWND hWnd);
    static HDC WINAPI HookedBeginPaint(HWND hWnd, LPPAINTSTRUCT lpPaint);
    static int WINAPI HookedFillRect(HDC hDC, const RECT* lprc, HBRUSH hbr);
    static HDC WINAPI HookedCreateCompatibleDC(HDC hDC);

    // Drawing implementation
    void DrawBackground(HDC hDC, HWND hWnd, const RECT* lprc);
    void CalculateImagePosition(const SIZE& wndSize, const SIZE& imgSize,
                                POINT& pos, SIZE& dstSize) const;

    // Per-thread window tracking
    struct WindowData {
        HDC hdc = nullptr;
        int imageIndex = 0;  // For future multi-image support
    };

    struct ThreadData {
        std::unordered_map<HWND, WindowData> duiWindows;
        SIZE lastSize = { 0, 0 };
    };

    // Get window from thread data by DC
    HWND GetWindowByDC(const ThreadData& threadData, HDC hDC, WindowData& outData) const;

    // Member variables
    bool m_initialized = false;
    mutable std::mutex m_mutex;
    std::unordered_map<DWORD, ThreadData> m_threadData;
    FolderBackgroundConfig m_config;
    std::unique_ptr<BackgroundBitmap> m_backgroundBitmap;

    // Original function pointers
    static inline decltype(&CreateWindowExW) s_originalCreateWindowExW = nullptr;
    static inline decltype(&DestroyWindow) s_originalDestroyWindow = nullptr;
    static inline decltype(&BeginPaint) s_originalBeginPaint = nullptr;
    static inline decltype(&FillRect) s_originalFillRect = nullptr;
    static inline decltype(&CreateCompatibleDC) s_originalCreateCompatibleDC = nullptr;
};

// Initialization functions called from dllmain
bool InitializeFolderBackgroundHooks();
void ShutdownFolderBackgroundHooks();

}  // namespace shelltabs
