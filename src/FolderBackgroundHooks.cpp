/*
 * ShellTabs - Folder Background Hooks Implementation
 *
 * Uses a WS_EX_LAYERED overlay child window on top of DirectUIHWND to display
 * background images.  The overlay is composited by DWM, eliminating flicker,
 * z-order issues, and opacity fade that occur when drawing directly in WM_PAINT.
 */

#include "../include/FolderBackgroundHooks.h"
#include "../include/MinHook.h"
#include "../include/Logging.h"

// GDI+ and AlphaBlend
#pragma comment(lib, "GdiPlus.lib")
#pragma comment(lib, "Msimg32.lib")

namespace shelltabs {

// ============================================================================
// BackgroundBitmap
// ============================================================================

BackgroundBitmap::BackgroundBitmap(const std::wstring& path) {
    auto src = std::make_unique<Gdiplus::Bitmap>(path.c_str());
    if (!src || src->GetLastStatus() != Gdiplus::Ok) {
        LogMessage(LogLevel::Warning, L"FolderBg: failed to load image: %ls", path.c_str());
        return;
    }

    m_size.cx = static_cast<LONG>(src->GetWidth());
    m_size.cy = static_cast<LONG>(src->GetHeight());
    if (m_size.cx == 0 || m_size.cy == 0) return;

    // Lock GDI+ bitmap bits as 32bpp ARGB so GDI+ handles any format conversion.
    Gdiplus::BitmapData bmpData = {};
    Gdiplus::Rect lockRect(0, 0, m_size.cx, m_size.cy);
    if (src->LockBits(&lockRect, Gdiplus::ImageLockModeRead,
                      PixelFormat32bppARGB, &bmpData) != Gdiplus::Ok) {
        LogMessage(LogLevel::Warning, L"FolderBg: LockBits failed: %ls", path.c_str());
        return;
    }

    const bool hasAlpha = (src->GetPixelFormat() & PixelFormatAlpha) != 0;

    HDC screenDC = CreateDCW(L"DISPLAY", nullptr, nullptr, nullptr);

    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = m_size.cx;
    bmi.bmiHeader.biHeight      = -m_size.cy;  // top-down
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* pBits = nullptr;
    m_bitmap = CreateDIBSection(screenDC, &bmi, DIB_RGB_COLORS, &pBits, nullptr, 0);
    if (screenDC) DeleteDC(screenDC);

    if (!m_bitmap || !pBits) {
        src->UnlockBits(&bmpData);
        return;
    }

    // Copy pixels, pre-multiplying alpha for AC_SRC_ALPHA.
    const BYTE* src_row = static_cast<const BYTE*>(bmpData.Scan0);
    BYTE*       dst_row = static_cast<BYTE*>(pBits);
    const int   dst_stride = m_size.cx * 4;

    for (int y = 0; y < m_size.cy; ++y) {
        const BYTE* s = src_row + y * bmpData.Stride;
              BYTE* d = dst_row + y * dst_stride;
        for (int x = 0; x < m_size.cx; ++x, s += 4, d += 4) {
            if (hasAlpha) {
                const UINT a = s[3];
                d[0] = static_cast<BYTE>(static_cast<UINT>(s[0]) * a / 255u);
                d[1] = static_cast<BYTE>(static_cast<UINT>(s[1]) * a / 255u);
                d[2] = static_cast<BYTE>(static_cast<UINT>(s[2]) * a / 255u);
                d[3] = static_cast<BYTE>(a);
            } else {
                d[0] = s[0];
                d[1] = s[1];
                d[2] = s[2];
                d[3] = 255;
            }
        }
    }

    src->UnlockBits(&bmpData);
    src.reset();

    LogMessage(LogLevel::Info, L"FolderBg: loaded %ls (%ldx%ld%ls)",
               path.c_str(), m_size.cx, m_size.cy, hasAlpha ? L" alpha" : L"");
}

BackgroundBitmap::~BackgroundBitmap() {
    if (m_bitmap) { DeleteObject(m_bitmap); m_bitmap = nullptr; }
}

// ============================================================================
// FolderBackgroundHooks
// ============================================================================

ATOM FolderBackgroundHooks::s_overlayClass = 0;

FolderBackgroundHooks& FolderBackgroundHooks::Instance() {
    static FolderBackgroundHooks instance;
    return instance;
}

FolderBackgroundHooks::~FolderBackgroundHooks() {
    Shutdown();
}

bool FolderBackgroundHooks::Initialize() {
    std::lock_guard<std::mutex> lock(m_mutex);
    if (m_initialized) return true;

    // Register overlay window class (once per process).
    if (s_overlayClass == 0) {
        WNDCLASSEXW wc = {};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = DefWindowProcW;
        wc.hInstance      = GetModuleHandleW(nullptr);
        wc.lpszClassName  = L"ShellTabsBgOverlay";
        wc.style          = CS_NOCLOSE;
        s_overlayClass = RegisterClassExW(&wc);
        if (s_overlayClass == 0) {
            LogMessage(LogLevel::Warning, L"FolderBg: RegisterClassExW failed for overlay (err=%lu)", GetLastError());
        }
    }

    // Hook CreateWindowExW to subclass newly-created DirectUIHWND windows.
    MH_STATUS st = MH_CreateHook(&CreateWindowExW, &HookedCreateWindowExW,
                                  (LPVOID*)&s_originalCreateWindowExW);
    if (st == MH_OK) {
        st = MH_EnableHook(&CreateWindowExW);
        if (st != MH_OK) {
            MH_RemoveHook(&CreateWindowExW);
            s_originalCreateWindowExW = nullptr;
        }
    }
    if (st != MH_OK) {
        LogMessage(LogLevel::Warning, L"FolderBg: CreateWindowExW hook failed (st=%d), "
                   L"new windows handled via SetFrameFolderPath", (int)st);
    }

    m_initialized = true;
    LogMessage(LogLevel::Info, L"FolderBackgroundHooks initialized");
    return true;
}

void FolderBackgroundHooks::Shutdown() {
    UnsubclassAll();
    DestroyAllOverlays();

    std::lock_guard<std::mutex> lock(m_mutex);
    if (!m_initialized) return;

    if (s_originalCreateWindowExW) {
        MH_DisableHook(&CreateWindowExW);
        MH_RemoveHook(&CreateWindowExW);
        s_originalCreateWindowExW = nullptr;
    }

    m_folderBitmaps.clear();
    m_universalBitmap.reset();
    m_frameFolderPaths.clear();
    m_enabled = false;
    m_initialized = false;

    LogMessage(LogLevel::Info, L"FolderBackgroundHooks shutdown");
}

// ============================================================================
// Helpers (must be above SetImages / SetFrameFolderPath which use them)
// ============================================================================

static std::wstring GetWindowClass(HWND hWnd) {
    if (!hWnd) return L"";
    wchar_t buf[128] = {};
    GetClassNameW(hWnd, buf, 128);
    return buf;
}

static HWND FindExplorerFrame(HWND startBelow) {
    HWND cur = GetParent(startBelow);
    for (int i = 0; i < 10 && cur; ++i) {
        wchar_t cls[64] = {};
        GetClassNameW(cur, cls, 64);
        if (_wcsicmp(cls, L"CabinetWClass") == 0) return cur;
        cur = GetParent(cur);
    }
    return nullptr;
}

struct DefViewInfo { HWND defView = nullptr; HWND duiHwnd = nullptr; };
static DefViewInfo FindDefViewAndDUI(HWND explorerFrame) {
    DefViewInfo info;
    EnumChildWindows(explorerFrame, [](HWND hwnd, LPARAM lp) -> BOOL {
        wchar_t cls[64] = {};
        GetClassNameW(hwnd, cls, 64);
        if (_wcsicmp(cls, L"SHELLDLL_DefView") == 0) {
            reinterpret_cast<DefViewInfo*>(lp)->defView = hwnd;
            return FALSE;
        }
        return TRUE;
    }, (LPARAM)&info);

    if (info.defView) {
        info.duiHwnd = FindWindowExW(info.defView, nullptr, L"DirectUIHWND", nullptr);
    }
    return info;
}

// ============================================================================
// SetImages
// ============================================================================

void FolderBackgroundHooks::SetImages(bool enabled,
                                       BackgroundPositionMode positionMode,
                                       BYTE opacity,
                                       const std::wstring& universalImagePath,
                                       const std::unordered_map<std::wstring, std::wstring>& folderImagePaths) {
    static thread_local bool s_inSetImages = false;
    if (s_inSetImages) {
        LogMessage(LogLevel::Warning, L"FolderBg: SetImages re-entrant call suppressed");
        return;
    }
    s_inSetImages = true;
    struct SetImagesGuard { ~SetImagesGuard() { s_inSetImages = false; } } setImGuard;

    // Load bitmaps OUTSIDE the mutex.
    std::shared_ptr<BackgroundBitmap> newUniversal;
    std::unordered_map<std::wstring, std::shared_ptr<BackgroundBitmap>> newFolderBitmaps;

    if (enabled) {
        if (!universalImagePath.empty()) {
            auto bmp = std::make_shared<BackgroundBitmap>(universalImagePath);
            if (bmp->IsValid()) newUniversal = std::move(bmp);
        }
        for (const auto& [folderKey, imagePath] : folderImagePaths) {
            if (folderKey.empty() || imagePath.empty()) continue;
            auto bmp = std::make_shared<BackgroundBitmap>(imagePath);
            if (bmp->IsValid()) newFolderBitmaps[folderKey] = std::move(bmp);
        }
    }

    const size_t folderCount = newFolderBitmaps.size();

    // Collect overlay entries before swapping bitmaps so we can update them after.
    std::vector<std::pair<HWND, HWND>> overlayEntries;

    {
        std::shared_ptr<BackgroundBitmap> oldUniversal;
        std::unordered_map<std::wstring, std::shared_ptr<BackgroundBitmap>> oldFolderBitmaps;
        {
            std::lock_guard<std::mutex> lock(m_mutex);
            m_enabled      = enabled;
            m_positionMode = positionMode;
            m_opacity      = opacity;
            oldUniversal          = std::move(m_universalBitmap);
            m_universalBitmap     = std::move(newUniversal);
            oldFolderBitmaps      = std::move(m_folderBitmaps);
            m_folderBitmaps       = std::move(newFolderBitmaps);

            for (const auto& [duiHwnd, overlayHwnd] : m_overlayWindows) {
                overlayEntries.push_back({duiHwnd, overlayHwnd});
            }
        }
    }

    LogMessage(LogLevel::Info, L"FolderBg: SetImages enabled=%d universal=%ls folders=%zu",
               (int)enabled,
               universalImagePath.empty() ? L"(none)" : universalImagePath.c_str(),
               folderCount);

    // Update all existing overlay windows with the new image/settings.
    for (const auto& [duiHwnd, overlayHwnd] : overlayEntries) {
        if (!IsWindow(duiHwnd) || !IsWindow(overlayHwnd)) continue;
        HWND explorerFrame = FindExplorerFrame(duiHwnd);
        if (explorerFrame) {
            // Force size tracking to refresh
            {
                std::lock_guard<std::mutex> lock(m_mutex);
                m_lastOverlaySize.erase(duiHwnd);
            }
            UpdateOverlayContent(duiHwnd, explorerFrame);
        }
    }
}

// ============================================================================
// SetFrameFolderPath / ClearFrameFolderPath
// ============================================================================

void FolderBackgroundHooks::SetFrameFolderPath(HWND explorerFrame,
                                                const std::wstring& normalizedFolderPath) {
    LogMessage(LogLevel::Verbose, L"FolderBg: SetFrameFolderPath frame=%p key=%ls",
               explorerFrame, normalizedFolderPath.empty() ? L"(empty)" : normalizedFolderPath.c_str());
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (explorerFrame) m_frameFolderPaths[explorerFrame] = normalizedFolderPath;
    }

    if (explorerFrame) {
        TrySubclassFrame(explorerFrame);

        // Update overlay content for any existing overlays in this frame.
        auto info = FindDefViewAndDUI(explorerFrame);
        if (info.duiHwnd) {
            {
                std::lock_guard<std::mutex> lock(m_mutex);
                m_lastOverlaySize.erase(info.duiHwnd);
            }
            UpdateOverlayContent(info.duiHwnd, explorerFrame);
        }
    }
}

void FolderBackgroundHooks::ClearFrameFolderPath(HWND explorerFrame) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_frameFolderPaths.erase(explorerFrame);
}

// ============================================================================
// Subclassing helpers
// ============================================================================

void FolderBackgroundHooks::TrySubclassFrame(HWND explorerFrame) {
    if (!explorerFrame || !IsWindow(explorerFrame)) return;
    auto info = FindDefViewAndDUI(explorerFrame);
    LogMessage(LogLevel::Info, L"FolderBg: TrySubclassFrame frame=%p defView=%p dui=%p",
               explorerFrame, info.defView, info.duiHwnd);

    if (info.defView) {
        SubclassHwnd(info.defView, explorerFrame, /*id=*/1);
    }
    if (info.duiHwnd) {
        SubclassHwnd(info.duiHwnd, explorerFrame, /*id=*/2);
        // Create the overlay window for this DirectUIHWND.
        CreateOverlayForDUI(info.duiHwnd, info.defView);
    }
}

void FolderBackgroundHooks::SubclassHwnd(HWND hwnd, HWND explorerFrame, UINT_PTR id) {
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (m_subclassedWindows.count(hwnd)) return;
    }

    if (!SetWindowSubclass(hwnd, SubclassProc, id, reinterpret_cast<DWORD_PTR>(explorerFrame))) {
        LogMessage(LogLevel::Warning, L"FolderBg: SetWindowSubclass failed hwnd=%p id=%llu",
                   hwnd, (unsigned long long)id);
        return;
    }

    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_subclassedWindows.insert(hwnd);
    }

    wchar_t cls[64] = {};
    GetClassNameW(hwnd, cls, 64);
    LogMessage(LogLevel::Info, L"FolderBg: subclassed %ls hwnd=%p frame=%p id=%llu",
               cls, hwnd, explorerFrame, (unsigned long long)id);
}

void FolderBackgroundHooks::UnsubclassAll() {
    std::unordered_set<HWND> toRemove;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        toRemove = m_subclassedWindows;
        m_subclassedWindows.clear();
    }
    for (HWND hwnd : toRemove) {
        if (!IsWindow(hwnd)) continue;
        RemoveWindowSubclass(hwnd, SubclassProc, 1);
        RemoveWindowSubclass(hwnd, SubclassProc, 2);
    }
}

// ============================================================================
// Overlay Window Management
// ============================================================================

void FolderBackgroundHooks::CreateOverlayForDUI(HWND duiHwnd, HWND defViewParent) {
    if (!duiHwnd || !defViewParent || s_overlayClass == 0) return;

    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (m_overlayWindows.count(duiHwnd)) return;  // already exists
    }

    RECT rc;
    if (!GetClientRect(duiHwnd, &rc)) return;

    // Create as a child of SHELLDLL_DefView, positioned over DirectUIHWND.
    // WS_EX_LAYERED: composited by DWM, persistent across DirectUI repaints.
    // WS_EX_TRANSPARENT: passes all mouse/keyboard input through to DirectUI.
    // WS_EX_NOACTIVATE: never steals focus.
    HWND overlay = CreateWindowExW(
        WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE,
        MAKEINTATOM(s_overlayClass),
        nullptr,
        WS_CHILD | WS_VISIBLE,
        0, 0, rc.right, rc.bottom,
        defViewParent,
        nullptr,
        GetModuleHandleW(nullptr),
        nullptr);

    if (!overlay) {
        LogMessage(LogLevel::Warning, L"FolderBg: CreateOverlayForDUI failed (err=%lu)", GetLastError());
        return;
    }

    // Place overlay above DirectUIHWND in z-order.
    SetWindowPos(overlay, duiHwnd, 0, 0, 0, 0,
                 SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_overlayWindows[duiHwnd] = overlay;
    }

    LogMessage(LogLevel::Info, L"FolderBg: created overlay %p for DUI %p", overlay, duiHwnd);

    // Populate with initial content.
    HWND explorerFrame = FindExplorerFrame(duiHwnd);
    if (explorerFrame) {
        UpdateOverlayContent(duiHwnd, explorerFrame);
    }
}

void FolderBackgroundHooks::DestroyOverlayForDUI(HWND duiHwnd) {
    HWND overlay = nullptr;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_overlayWindows.find(duiHwnd);
        if (it == m_overlayWindows.end()) return;
        overlay = it->second;
        m_overlayWindows.erase(it);
        m_lastOverlaySize.erase(duiHwnd);
    }
    if (overlay && IsWindow(overlay)) {
        DestroyWindow(overlay);
    }
    LogMessage(LogLevel::Info, L"FolderBg: destroyed overlay for DUI %p", duiHwnd);
}

void FolderBackgroundHooks::DestroyAllOverlays() {
    std::unordered_map<HWND, HWND> overlays;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        overlays = m_overlayWindows;
        m_overlayWindows.clear();
        m_lastOverlaySize.clear();
    }
    for (const auto& [duiHwnd, overlay] : overlays) {
        if (overlay && IsWindow(overlay)) {
            DestroyWindow(overlay);
        }
    }
}

void FolderBackgroundHooks::RepositionOverlay(HWND duiHwnd) {
    HWND overlay = nullptr;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_overlayWindows.find(duiHwnd);
        if (it == m_overlayWindows.end()) return;
        overlay = it->second;
    }
    if (!overlay || !IsWindow(overlay)) return;

    RECT rc;
    if (!GetClientRect(duiHwnd, &rc)) return;

    SetWindowPos(overlay, duiHwnd,
                 0, 0, rc.right, rc.bottom,
                 SWP_NOACTIVATE);
}

void FolderBackgroundHooks::UpdateOverlayContent(HWND duiHwnd, HWND explorerFrame) {
    HWND overlay = nullptr;
    std::shared_ptr<const BackgroundBitmap> bmp;
    BackgroundPositionMode posMode;
    BYTE opacity;
    bool enabled;

    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto it = m_overlayWindows.find(duiHwnd);
        if (it == m_overlayWindows.end()) return;
        overlay = it->second;
        enabled = m_enabled;
        if (!enabled) {
            // Hide overlay when disabled.
            if (overlay && IsWindow(overlay)) {
                // Set fully transparent to hide.
                SIZE zero = {0, 0};
                POINT ptZero = {0, 0};
                BLENDFUNCTION bf = {};
                bf.BlendOp = AC_SRC_OVER;
                bf.SourceConstantAlpha = 0;
                bf.AlphaFormat = AC_SRC_ALPHA;
                UpdateLayeredWindow(overlay, nullptr, nullptr, &zero, nullptr, &ptZero, 0, &bf, ULW_ALPHA);
            }
            return;
        }
        bmp = ResolveBitmapForFrame(explorerFrame);
        posMode = m_positionMode;
        opacity = m_opacity;
    }

    if (!overlay || !IsWindow(overlay)) return;

    RECT cr;
    if (!GetClientRect(duiHwnd, &cr)) return;
    SIZE wndSize = {cr.right, cr.bottom};
    if (wndSize.cx <= 0 || wndSize.cy <= 0) return;

    // Check if update is needed (size unchanged and we've already rendered).
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        auto sizeIt = m_lastOverlaySize.find(duiHwnd);
        if (sizeIt != m_lastOverlaySize.end() &&
            sizeIt->second.cx == wndSize.cx && sizeIt->second.cy == wndSize.cy) {
            return;  // Already up to date
        }
        m_lastOverlaySize[duiHwnd] = wndSize;
    }

    if (!bmp || !bmp->IsValid()) {
        // No image: make overlay fully transparent.
        SIZE zero = {0, 0};
        POINT ptZero = {0, 0};
        BLENDFUNCTION bf = {};
        bf.BlendOp = AC_SRC_OVER;
        bf.SourceConstantAlpha = 0;
        bf.AlphaFormat = AC_SRC_ALPHA;
        UpdateLayeredWindow(overlay, nullptr, nullptr, &zero, nullptr, &ptZero, 0, &bf, ULW_ALPHA);
        return;
    }

    SIZE imgSize = bmp->GetSize();
    if (imgSize.cx <= 0 || imgSize.cy <= 0) return;

    // Create 32bpp ARGB DIB section for the overlay surface.
    HDC screenDC = GetDC(nullptr);
    if (!screenDC) return;

    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = wndSize.cx;
    bmi.bmiHeader.biHeight      = -wndSize.cy;  // top-down
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* pBits = nullptr;
    HBITMAP dibSection = CreateDIBSection(screenDC, &bmi, DIB_RGB_COLORS, &pBits, nullptr, 0);
    if (!dibSection || !pBits) {
        ReleaseDC(nullptr, screenDC);
        return;
    }

    // Fill with transparent (zeroed ARGB = fully transparent).
    memset(pBits, 0, static_cast<size_t>(wndSize.cx) * wndSize.cy * 4);

    HDC memDC = CreateCompatibleDC(screenDC);
    if (!memDC) {
        DeleteObject(dibSection);
        ReleaseDC(nullptr, screenDC);
        return;
    }
    HGDIOBJ oldMemBmp = SelectObject(memDC, dibSection);

    // Draw the background image into the DIB.
    HDC srcDC = CreateCompatibleDC(screenDC);
    if (srcDC) {
        HGDIOBJ oldSrcBmp = SelectObject(srcDC, bmp->GetBitmap());
        if (oldSrcBmp && oldSrcBmp != HGDI_ERROR) {
            // Use full opacity in the bitmap — SourceConstantAlpha in
            // UpdateLayeredWindow will apply the user's configured opacity.
            BLENDFUNCTION bfDraw = {};
            bfDraw.BlendOp             = AC_SRC_OVER;
            bfDraw.SourceConstantAlpha = 255;
            bfDraw.AlphaFormat         = AC_SRC_ALPHA;

            if (posMode == BackgroundPositionMode::kTile) {
                for (int ty = 0; ty < wndSize.cy; ty += imgSize.cy) {
                    for (int tx = 0; tx < wndSize.cx; tx += imgSize.cx) {
                        AlphaBlend(memDC, tx, ty, imgSize.cx, imgSize.cy,
                                   srcDC, 0, 0, imgSize.cx, imgSize.cy, bfDraw);
                    }
                }
            } else {
                POINT pos;
                SIZE  dstSize;
                CalculateImagePosition(wndSize, imgSize, posMode, pos, dstSize);
                if (dstSize.cx > 0 && dstSize.cy > 0) {
                    AlphaBlend(memDC, pos.x, pos.y, dstSize.cx, dstSize.cy,
                               srcDC, 0, 0, imgSize.cx, imgSize.cy, bfDraw);
                }
            }

            SelectObject(srcDC, oldSrcBmp);
        }
        DeleteDC(srcDC);
    }

    // UpdateLayeredWindow with the user's configured opacity.
    POINT ptSrc = {0, 0};
    BLENDFUNCTION bf = {};
    bf.BlendOp             = AC_SRC_OVER;
    bf.SourceConstantAlpha = opacity;
    bf.AlphaFormat         = AC_SRC_ALPHA;

    UpdateLayeredWindow(overlay, screenDC, nullptr, &wndSize, memDC, &ptSrc, 0, &bf, ULW_ALPHA);

    SelectObject(memDC, oldMemBmp);
    DeleteDC(memDC);
    DeleteObject(dibSection);
    ReleaseDC(nullptr, screenDC);

    LogMessage(LogLevel::Verbose, L"FolderBg: updated overlay for DUI %p (%ldx%ld) opacity=%d",
               duiHwnd, wndSize.cx, wndSize.cy, (int)opacity);
}

// ============================================================================
// SubclassProc
// ============================================================================

LRESULT CALLBACK FolderBackgroundHooks::SubclassProc(
    HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    // Always handle WM_NCDESTROY first — must never be skipped.
    if (msg == WM_NCDESTROY) {
        FolderBackgroundHooks& inst = Instance();
        inst.DestroyOverlayForDUI(hwnd);
        RemoveWindowSubclass(hwnd, SubclassProc, uIdSubclass);
        {
            std::lock_guard<std::mutex> lock(inst.m_mutex);
            inst.m_subclassedWindows.erase(hwnd);
        }
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    HWND explorerFrame = reinterpret_cast<HWND>(dwRefData);
    FolderBackgroundHooks& inst = Instance();

    switch (msg) {
    case WM_SIZE:
        // Reposition and update the overlay when the DirectUI window resizes.
        if (inst.m_enabled) {
            inst.RepositionOverlay(hwnd);
            inst.UpdateOverlayContent(hwnd, explorerFrame);
        }
        break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

// ============================================================================
// HookedCreateWindowExW
// ============================================================================

HWND WINAPI FolderBackgroundHooks::HookedCreateWindowExW(
    DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
    DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam)
{
    HWND hWnd = s_originalCreateWindowExW(dwExStyle, lpClassName, lpWindowName, dwStyle,
                                           X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
    if (!hWnd) return hWnd;

    if (GetWindowClass(hWnd) != L"DirectUIHWND") return hWnd;
    if (GetWindowClass(hWndParent) != L"SHELLDLL_DefView") return hWnd;

    HWND explorerFrame = FindExplorerFrame(hWndParent);
    if (!explorerFrame) return hWnd;

    FolderBackgroundHooks& inst = Instance();
    inst.SubclassHwnd(hWndParent, explorerFrame, /*id=*/1);
    inst.SubclassHwnd(hWnd,       explorerFrame, /*id=*/2);
    inst.CreateOverlayForDUI(hWnd, hWndParent);

    return hWnd;
}

// ============================================================================
// Drawing helpers (used by overlay)
// ============================================================================

std::shared_ptr<const BackgroundBitmap>
FolderBackgroundHooks::ResolveBitmapForFrame(HWND explorerFrame) const {
    auto frameIt = m_frameFolderPaths.find(explorerFrame);
    if (frameIt != m_frameFolderPaths.end() && !frameIt->second.empty()) {
        auto folderIt = m_folderBitmaps.find(frameIt->second);
        if (folderIt != m_folderBitmaps.end() && folderIt->second && folderIt->second->IsValid())
            return folderIt->second;
    }
    if (m_universalBitmap && m_universalBitmap->IsValid())
        return m_universalBitmap;
    return nullptr;
}

void FolderBackgroundHooks::CalculateImagePosition(const SIZE& wndSize, const SIZE& imgSize,
                                                    BackgroundPositionMode mode,
                                                    POINT& pos, SIZE& dstSize) const {
    dstSize = imgSize;
    switch (mode) {
    case BackgroundPositionMode::kTopLeft:
        pos = {0, 0};
        break;
    case BackgroundPositionMode::kTopRight:
        pos = {wndSize.cx - imgSize.cx, 0};
        break;
    case BackgroundPositionMode::kBottomLeft:
        pos = {0, wndSize.cy - imgSize.cy};
        break;
    case BackgroundPositionMode::kBottomRight:
        pos = {wndSize.cx - imgSize.cx, wndSize.cy - imgSize.cy};
        break;
    case BackgroundPositionMode::kCenter:
        pos = {(wndSize.cx - imgSize.cx) / 2, (wndSize.cy - imgSize.cy) / 2};
        break;
    case BackgroundPositionMode::kStretch:
        pos = {0, 0};
        dstSize = wndSize;
        break;
    case BackgroundPositionMode::kZoomFill: {
        int newW = wndSize.cx;
        int newH = static_cast<int>(static_cast<float>(imgSize.cy) *
                                    static_cast<float>(wndSize.cx) / imgSize.cx);
        if (newH < wndSize.cy) {
            newH = wndSize.cy;
            newW = static_cast<int>(static_cast<float>(imgSize.cx) *
                                    static_cast<float>(wndSize.cy) / imgSize.cy);
        }
        pos     = {(wndSize.cx - newW) / 2, (wndSize.cy - newH) / 2};
        dstSize = {newW, newH};
        break;
    }
    default:
        pos = {wndSize.cx - imgSize.cx, wndSize.cy - imgSize.cy};
        break;
    }
}

// ============================================================================
// Module-level init/shutdown
// ============================================================================

bool InitializeFolderBackgroundHooks() {
    return FolderBackgroundHooks::Instance().Initialize();
}

void ShutdownFolderBackgroundHooks() {
    FolderBackgroundHooks::Instance().Shutdown();
}

}  // namespace shelltabs
