/*
 * ShellTabs - Folder Background Hooks Implementation
 * DirectUI-based folder background rendering system using API hooking
 *
 * Based on ExplorerBgTool by Maplespe (winmoes.com)
 * Adapted for ShellTabs by integrating with existing infrastructure
 */

#include "../include/FolderBackgroundHooks.h"
#include "../include/MinHook.h"
#include "../include/Logging.h"

// GDI+ and AlphaBlend
#pragma comment(lib, "GdiPlus.lib")
#pragma comment(lib, "Msimg32.lib")

namespace shelltabs {

// ============================================================================
// BackgroundBitmap Implementation
// ============================================================================

BackgroundBitmap::BackgroundBitmap(const std::wstring& path) {
    // Load image using GDI+
    m_source = std::make_unique<Gdiplus::Bitmap>(path.c_str());
    if (!m_source || m_source->GetLastStatus() != Gdiplus::Ok) {
        LogMessage(LogLevel::Error, L"Failed to load background image: %ls", path.c_str());
        m_source.reset();
        return;
    }

    m_size.cx = static_cast<LONG>(m_source->GetWidth());
    m_size.cy = static_cast<LONG>(m_source->GetHeight());

    if (m_size.cx == 0 || m_size.cy == 0) {
        LogMessage(LogLevel::Error, L"Background image has zero dimensions: %ls", path.c_str());
        m_source.reset();
        return;
    }

    // Create a memory DC and compatible bitmap for efficient AlphaBlend
    HDC screenDC = GetDC(nullptr);
    m_memDC = CreateCompatibleDC(screenDC);

    // Create a 32-bit DIB section for alpha blending
    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = m_size.cx;
    bmi.bmiHeader.biHeight = -m_size.cy;  // Top-down DIB
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* pBits = nullptr;
    m_bitmap = CreateDIBSection(screenDC, &bmi, DIB_RGB_COLORS, &pBits, nullptr, 0);
    ReleaseDC(nullptr, screenDC);

    if (!m_bitmap || !pBits) {
        LogMessage(LogLevel::Error, L"Failed to create DIB section for background image");
        if (m_memDC) {
            DeleteDC(m_memDC);
            m_memDC = nullptr;
        }
        m_source.reset();
        return;
    }

    SelectObject(m_memDC, m_bitmap);

    // Draw the GDI+ image to our bitmap
    Gdiplus::Graphics graphics(m_memDC);
    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
    graphics.DrawImage(m_source.get(), 0, 0, m_size.cx, m_size.cy);

    LogMessage(LogLevel::Info, L"Loaded background image: %ls (%ldx%ld)", path.c_str(), m_size.cx, m_size.cy);
}

BackgroundBitmap::~BackgroundBitmap() {
    if (m_bitmap) {
        DeleteObject(m_bitmap);
        m_bitmap = nullptr;
    }
    if (m_memDC) {
        DeleteDC(m_memDC);
        m_memDC = nullptr;
    }
    m_source.reset();
}

// ============================================================================
// FolderBackgroundHooks Implementation
// ============================================================================

FolderBackgroundHooks& FolderBackgroundHooks::Instance() {
    static FolderBackgroundHooks instance;
    return instance;
}

FolderBackgroundHooks::~FolderBackgroundHooks() {
    Shutdown();
}

bool FolderBackgroundHooks::Initialize() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (m_initialized) {
        return true;
    }

    // Initialize MinHook
    MH_STATUS status = MH_Initialize();
    if (status != MH_OK && status != MH_ERROR_ALREADY_INITIALIZED) {
        LogMessage(LogLevel::Error, L"Failed to initialize MinHook: %d", static_cast<int>(status));
        return false;
    }

    // Create hooks for Win32 API functions
    bool allHooksCreated = true;

    // Hook CreateWindowExW
    if (MH_CreateHook(&CreateWindowExW, &HookedCreateWindowExW,
                      reinterpret_cast<LPVOID*>(&s_originalCreateWindowExW)) != MH_OK) {
        LogMessage(LogLevel::Error, L"Failed to create hook for CreateWindowExW");
        allHooksCreated = false;
    }

    // Hook DestroyWindow
    if (MH_CreateHook(&DestroyWindow, &HookedDestroyWindow,
                      reinterpret_cast<LPVOID*>(&s_originalDestroyWindow)) != MH_OK) {
        LogMessage(LogLevel::Error, L"Failed to create hook for DestroyWindow");
        allHooksCreated = false;
    }

    // Hook BeginPaint
    if (MH_CreateHook(&BeginPaint, &HookedBeginPaint,
                      reinterpret_cast<LPVOID*>(&s_originalBeginPaint)) != MH_OK) {
        LogMessage(LogLevel::Error, L"Failed to create hook for BeginPaint");
        allHooksCreated = false;
    }

    // Hook FillRect
    if (MH_CreateHook(&FillRect, &HookedFillRect,
                      reinterpret_cast<LPVOID*>(&s_originalFillRect)) != MH_OK) {
        LogMessage(LogLevel::Error, L"Failed to create hook for FillRect");
        allHooksCreated = false;
    }

    // Hook CreateCompatibleDC
    if (MH_CreateHook(&CreateCompatibleDC, &HookedCreateCompatibleDC,
                      reinterpret_cast<LPVOID*>(&s_originalCreateCompatibleDC)) != MH_OK) {
        LogMessage(LogLevel::Error, L"Failed to create hook for CreateCompatibleDC");
        allHooksCreated = false;
    }

    if (!allHooksCreated) {
        MH_Uninitialize();
        return false;
    }

    // Enable all hooks
    if (MH_EnableHook(MH_ALL_HOOKS) != MH_OK) {
        LogMessage(LogLevel::Error, L"Failed to enable hooks");
        MH_Uninitialize();
        return false;
    }

    m_initialized = true;
    LogMessage(LogLevel::Info, L"FolderBackgroundHooks initialized successfully");
    return true;
}

void FolderBackgroundHooks::Shutdown() {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (!m_initialized) {
        return;
    }

    // Disable and remove all hooks
    MH_DisableHook(MH_ALL_HOOKS);
    MH_Uninitialize();

    // Clear cached data
    m_threadData.clear();
    m_backgroundBitmap.reset();
    m_initialized = false;

    LogMessage(LogLevel::Info, L"FolderBackgroundHooks shutdown complete");
}

void FolderBackgroundHooks::SetConfig(const FolderBackgroundConfig& config) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_config = config;
}

FolderBackgroundConfig FolderBackgroundHooks::GetConfig() const {
    std::lock_guard<std::mutex> lock(m_mutex);
    return m_config;
}

bool FolderBackgroundHooks::LoadImage(const std::wstring& path) {
    std::lock_guard<std::mutex> lock(m_mutex);

    if (path.empty()) {
        m_backgroundBitmap.reset();
        return true;
    }

    auto newBitmap = std::make_unique<BackgroundBitmap>(path);
    if (!newBitmap->IsValid()) {
        return false;
    }

    m_backgroundBitmap = std::move(newBitmap);
    m_config.imagePath = path;
    return true;
}

void FolderBackgroundHooks::ClearImage() {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_backgroundBitmap.reset();
    m_config.imagePath.clear();
}

void FolderBackgroundHooks::ReloadSettings() {
    // TODO: Reload settings from OptionsStore
    // This will be called when the user changes background settings
}

// ============================================================================
// Helper function to get class name safely
// ============================================================================

static std::wstring GetWindowClassName(HWND hWnd) {
    if (!hWnd) return L"";
    wchar_t className[256] = {};
    GetClassNameW(hWnd, className, 256);
    return className;
}

// ============================================================================
// Hook Implementations
// ============================================================================

HWND WINAPI FolderBackgroundHooks::HookedCreateWindowExW(
    DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
    DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam)
{
    // Call original function first
    HWND hWnd = s_originalCreateWindowExW(dwExStyle, lpClassName, lpWindowName, dwStyle,
        X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);

    if (!hWnd) {
        return hWnd;
    }

    // Check if this is a DirectUIHWND window under SHELLDLL_DefView
    std::wstring className = GetWindowClassName(hWnd);
    if (className != L"DirectUIHWND") {
        return hWnd;
    }

    std::wstring parentClassName = GetWindowClassName(hWndParent);
    if (parentClassName != L"SHELLDLL_DefView") {
        return hWnd;
    }

    // Check grandparent - should be ShellTabWindowClass or #32770 (dialog)
    HWND grandParent = GetParent(hWndParent);
    std::wstring grandParentClassName = GetWindowClassName(grandParent);
    if (grandParentClassName != L"ShellTabWindowClass" && grandParentClassName != L"#32770") {
        return hWnd;
    }

    // This is an Explorer folder view - track it
    FolderBackgroundHooks& instance = Instance();
    std::lock_guard<std::mutex> lock(instance.m_mutex);

    DWORD threadId = GetCurrentThreadId();
    ThreadData& threadData = instance.m_threadData[threadId];

    WindowData windowData;
    windowData.hdc = nullptr;
    windowData.imageIndex = 0;
    threadData.duiWindows[hWnd] = windowData;

    return hWnd;
}

BOOL WINAPI FolderBackgroundHooks::HookedDestroyWindow(HWND hWnd) {
    // Remove window from tracking before destroying
    FolderBackgroundHooks& instance = Instance();
    {
        std::lock_guard<std::mutex> lock(instance.m_mutex);

        DWORD threadId = GetCurrentThreadId();
        auto threadIt = instance.m_threadData.find(threadId);
        if (threadIt != instance.m_threadData.end()) {
            auto windowIt = threadIt->second.duiWindows.find(hWnd);
            if (windowIt != threadIt->second.duiWindows.end()) {
                threadIt->second.duiWindows.erase(windowIt);

                // Clean up empty thread data
                if (threadIt->second.duiWindows.empty()) {
                    instance.m_threadData.erase(threadIt);
                }
            }
        }
    }

    return s_originalDestroyWindow(hWnd);
}

HDC WINAPI FolderBackgroundHooks::HookedBeginPaint(HWND hWnd, LPPAINTSTRUCT lpPaint) {
    HDC hDC = s_originalBeginPaint(hWnd, lpPaint);

    FolderBackgroundHooks& instance = Instance();
    std::lock_guard<std::mutex> lock(instance.m_mutex);

    DWORD threadId = GetCurrentThreadId();
    auto threadIt = instance.m_threadData.find(threadId);
    if (threadIt != instance.m_threadData.end()) {
        auto windowIt = threadIt->second.duiWindows.find(hWnd);
        if (windowIt != threadIt->second.duiWindows.end()) {
            windowIt->second.hdc = hDC;
        }
    }

    return hDC;
}

int WINAPI FolderBackgroundHooks::HookedFillRect(HDC hDC, const RECT* lprc, HBRUSH hbr) {
    // Call original first
    int ret = s_originalFillRect(hDC, lprc, hbr);

    FolderBackgroundHooks& instance = Instance();

    // Quick check without lock to avoid performance impact
    if (!instance.m_config.enabled || !instance.m_backgroundBitmap) {
        return ret;
    }

    std::lock_guard<std::mutex> lock(instance.m_mutex);

    // Double-check after acquiring lock
    if (!instance.m_config.enabled || !instance.m_backgroundBitmap ||
        !instance.m_backgroundBitmap->IsValid()) {
        return ret;
    }

    DWORD threadId = GetCurrentThreadId();
    auto threadIt = instance.m_threadData.find(threadId);
    if (threadIt == instance.m_threadData.end()) {
        return ret;
    }

    // Find the window for this DC
    WindowData windowData;
    HWND hWnd = instance.GetWindowByDC(threadIt->second, hDC, windowData);
    if (!hWnd) {
        return ret;
    }

    // Draw the background
    instance.DrawBackground(hDC, hWnd, lprc);

    return ret;
}

HDC WINAPI FolderBackgroundHooks::HookedCreateCompatibleDC(HDC hDC) {
    HDC retDC = s_originalCreateCompatibleDC(hDC);

    FolderBackgroundHooks& instance = Instance();
    std::lock_guard<std::mutex> lock(instance.m_mutex);

    DWORD threadId = GetCurrentThreadId();
    auto threadIt = instance.m_threadData.find(threadId);
    if (threadIt != instance.m_threadData.end()) {
        // Find window from source DC and associate the new compatible DC
        HWND hWnd = WindowFromDC(hDC);
        auto windowIt = threadIt->second.duiWindows.find(hWnd);
        if (windowIt != threadIt->second.duiWindows.end()) {
            windowIt->second.hdc = retDC;
        }
    }

    return retDC;
}

// ============================================================================
// Drawing Implementation
// ============================================================================

HWND FolderBackgroundHooks::GetWindowByDC(const ThreadData& threadData, HDC hDC,
                                           WindowData& outData) const {
    for (const auto& pair : threadData.duiWindows) {
        if (pair.second.hdc == hDC) {
            outData = pair.second;
            return pair.first;
        }
    }
    return nullptr;
}

void FolderBackgroundHooks::CalculateImagePosition(const SIZE& wndSize, const SIZE& imgSize,
                                                    POINT& pos, SIZE& dstSize) const {
    dstSize = imgSize;

    switch (m_config.positionMode) {
        case BackgroundPositionMode::kTopLeft:
            pos = { 0, 0 };
            break;

        case BackgroundPositionMode::kTopRight:
            pos.x = wndSize.cx - imgSize.cx;
            pos.y = 0;
            break;

        case BackgroundPositionMode::kBottomLeft:
            pos.x = 0;
            pos.y = wndSize.cy - imgSize.cy;
            break;

        case BackgroundPositionMode::kBottomRight:
            pos.x = wndSize.cx - imgSize.cx;
            pos.y = wndSize.cy - imgSize.cy;
            break;

        case BackgroundPositionMode::kCenter:
            pos.x = (wndSize.cx - imgSize.cx) / 2;
            pos.y = (wndSize.cy - imgSize.cy) / 2;
            break;

        case BackgroundPositionMode::kStretch:
            pos = { 0, 0 };
            dstSize.cx = wndSize.cx;
            dstSize.cy = wndSize.cy;
            break;

        case BackgroundPositionMode::kZoomFill:
            {
                // Calculate aspect ratios
                auto calcAspectRatio = [](int fromWidth, int fromHeight,
                                          int toWidthOrHeight, bool isWidth) -> int {
                    if (isWidth) {
                        return static_cast<int>(round(
                            static_cast<float>(fromHeight) *
                            (static_cast<float>(toWidthOrHeight) / static_cast<float>(fromWidth))));
                    } else {
                        return static_cast<int>(round(
                            static_cast<float>(fromWidth) *
                            (static_cast<float>(toWidthOrHeight) / static_cast<float>(fromHeight))));
                    }
                };

                // Scale by height first
                int newWidth = calcAspectRatio(imgSize.cx, imgSize.cy, wndSize.cy, false);
                int newHeight = wndSize.cy;

                pos.x = (newWidth - wndSize.cx) / 2;
                if (pos.x != 0) pos.x = -pos.x;
                pos.y = 0;

                // If height-based scaling doesn't fill width, scale by width instead
                if (newWidth < wndSize.cx) {
                    newWidth = wndSize.cx;
                    newHeight = calcAspectRatio(imgSize.cx, imgSize.cy, wndSize.cx, true);
                    pos.x = 0;
                    pos.y = (newHeight - wndSize.cy) / 2;
                    if (pos.y != 0) pos.y = -pos.y;
                }

                dstSize.cx = newWidth;
                dstSize.cy = newHeight;
            }
            break;

        default:
            // Default to bottom-right
            pos.x = wndSize.cx - imgSize.cx;
            pos.y = wndSize.cy - imgSize.cy;
            break;
    }
}

void FolderBackgroundHooks::DrawBackground(HDC hDC, HWND hWnd, const RECT* lprc) {
    if (!m_backgroundBitmap || !m_backgroundBitmap->IsValid()) {
        return;
    }

    // Get window size
    RECT windowRect;
    GetWindowRect(hWnd, &windowRect);
    SIZE wndSize = {
        windowRect.right - windowRect.left,
        windowRect.bottom - windowRect.top
    };

    // Get thread data for size change detection
    DWORD threadId = GetCurrentThreadId();
    auto threadIt = m_threadData.find(threadId);
    if (threadIt != m_threadData.end()) {
        // If window size changed and we're not using top-left positioning,
        // invalidate to trigger full repaint (prevents artifacts)
        if ((threadIt->second.lastSize.cx != wndSize.cx ||
             threadIt->second.lastSize.cy != wndSize.cy) &&
            m_config.positionMode != BackgroundPositionMode::kTopLeft) {
            InvalidateRect(hWnd, nullptr, TRUE);
        }
        threadIt->second.lastSize = wndSize;
    }

    // Calculate image position and destination size
    SIZE imgSize = m_backgroundBitmap->GetSize();
    POINT pos;
    SIZE dstSize;
    CalculateImagePosition(wndSize, imgSize, pos, dstSize);

    // Save DC state and set clip rect
    SaveDC(hDC);
    IntersectClipRect(hDC, lprc->left, lprc->top, lprc->right, lprc->bottom);

    // Draw with alpha blending
    BLENDFUNCTION bf = {};
    bf.BlendOp = AC_SRC_OVER;
    bf.BlendFlags = 0;
    bf.SourceConstantAlpha = m_config.opacity;
    bf.AlphaFormat = AC_SRC_ALPHA;

    AlphaBlend(hDC, pos.x, pos.y, dstSize.cx, dstSize.cy,
               m_backgroundBitmap->GetMemDC(), 0, 0, imgSize.cx, imgSize.cy, bf);

    RestoreDC(hDC, -1);
}

// ============================================================================
// Module-level initialization functions
// ============================================================================

bool InitializeFolderBackgroundHooks() {
    return FolderBackgroundHooks::Instance().Initialize();
}

void ShutdownFolderBackgroundHooks() {
    FolderBackgroundHooks::Instance().Shutdown();
}

}  // namespace shelltabs
