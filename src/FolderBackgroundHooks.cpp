/*
 * ShellTabs - Folder Background Hooks Implementation
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

    // Create a screen DC just for CreateDIBSection; we do NOT keep a per-bitmap DC
    // because GDI DCs are thread-affine: a DC created on the SetImages thread cannot
    // be used on the window's owner thread in DrawBackground → crash.
    // Instead we store only the HBITMAP (cross-thread safe) and create a temporary
    // DC on the drawing thread at paint time.
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

    // Copy pixels: GDI+ ARGB in memory [B,G,R,A] → DIBSection BGRA [B,G,R,A].
    // Pre-multiply alpha for AC_SRC_ALPHA; set alpha=255 for opaque formats (JPEG, etc.).
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
                d[3] = 255;  // fully opaque
            }
        }
    }

    src->UnlockBits(&bmpData);
    // Don't keep m_source alive.  The Gdiplus::Bitmap is no longer needed
    // after pixel data has been copied into the DIBSection.  Releasing it NOW
    // (inside the ctor, while the call stack is still shallow) is critical:
    // Gdiplus::~Bitmap triggers WIC/COM cleanup which can pump STA messages
    // (KiUserCallbackDispatcher) on this thread.  If the bitmap is destroyed
    // later (e.g. at the end of SetImages during a deeply-nested re-entrant
    // callback chain from SetSite), the message pump can re-enter
    // explorerframe → ThemeHooks → use-after-free → crash.
    src.reset();

    LogMessage(LogLevel::Info, L"FolderBg: loaded %ls (%ldx%ld%ls)",
               path.c_str(), m_size.cx, m_size.cy, hasAlpha ? L" alpha" : L"");
}

BackgroundBitmap::~BackgroundBitmap() {
    if (m_bitmap) { DeleteObject(m_bitmap); m_bitmap = nullptr; }
    // No DC to delete — we don't store a per-bitmap DC (see constructor comment).
}

// ============================================================================
// FolderBackgroundHooks
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
    if (m_initialized) return true;

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
    // Non-fatal if CreateWindowExW hook fails (e.g. ThemeHooks already owns it).
    // Existing windows will be subclassed via SetFrameFolderPath.
    if (st != MH_OK) {
        LogMessage(LogLevel::Warning, L"FolderBg: CreateWindowExW hook failed (st=%d), "
                   L"new windows handled via SetFrameFolderPath", (int)st);
    }

    m_initialized = true;
    LogMessage(LogLevel::Info, L"FolderBackgroundHooks initialized");
    return true;
}

void FolderBackgroundHooks::Shutdown() {
    // Unsubclass all windows (must be called on UI thread if windows are still alive).
    UnsubclassAll();

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
// SetImages
// ============================================================================

void FolderBackgroundHooks::SetImages(bool enabled,
                                       BackgroundPositionMode positionMode,
                                       BYTE opacity,
                                       const std::wstring& universalImagePath,
                                       const std::unordered_map<std::wstring, std::wstring>& folderImagePaths) {
    // Re-entrancy guard: GDI+ (used in BackgroundBitmap ctor) is NOT re-entrant.
    // During JPEG loading, GDI+ calls into WIC (COM), which can trigger apartment
    // message delivery (KiUserCallbackDispatcher → COM window → BHO::Invoke →
    // NavigateComplete2 → UpdateBreadcrumbSubclass → SetImages again).
    // A re-entrant call would call Gdiplus::Bitmap while GDI+ is already active on
    // this thread, corrupting GDI+'s thread-local state → crash.
    static thread_local bool s_inSetImages = false;
    if (s_inSetImages) {
        LogMessage(LogLevel::Warning, L"FolderBg: SetImages re-entrant call suppressed");
        return;
    }
    s_inSetImages = true;
    struct SetImagesGuard { ~SetImagesGuard() { s_inSetImages = false; } } setImGuard;

    // Load bitmaps OUTSIDE the mutex: BackgroundBitmap ctor calls GDI/GDI+ operations.
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

    // Swap bitmaps under lock.  Old bitmaps are saved into locals and explicitly
    // destroyed BEFORE any InvalidateRect calls, so the GDI DeleteObject runs
    // in a controlled scope with a shallow call stack.
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
        }
        // Old bitmaps destroyed here at end of block — safe, mutex is not held,
        // and since we no longer keep Gdiplus::Bitmap alive (see BackgroundBitmap ctor),
        // destruction is just DeleteObject(HBITMAP) which doesn't pump COM messages.
    }

    LogMessage(LogLevel::Info, L"FolderBg: SetImages enabled=%d universal=%ls folders=%zu",
               (int)enabled,
               universalImagePath.empty() ? L"(none)" : universalImagePath.c_str(),
               folderCount);

    // Invalidate all subclassed windows so they repaint with the new settings.
    std::unordered_set<HWND> subclassed;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        subclassed = m_subclassedWindows;
    }
    for (HWND hwnd : subclassed) {
        if (IsWindow(hwnd))
            InvalidateRect(hwnd, nullptr, TRUE);
    }
}

// ============================================================================
// SetFrameFolderPath / ClearFrameFolderPath
// ============================================================================

// Helpers -----------------------------------------------------------------

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

// Find SHELLDLL_DefView (anywhere under explorerFrame) and return it + its DirectUIHWND child.
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

// -------------------------------------------------------------------------

void FolderBackgroundHooks::SetFrameFolderPath(HWND explorerFrame,
                                                const std::wstring& normalizedFolderPath) {
    LogMessage(LogLevel::Verbose, L"FolderBg: SetFrameFolderPath frame=%p key=%ls",
               explorerFrame, normalizedFolderPath.empty() ? L"(empty)" : normalizedFolderPath.c_str());
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (explorerFrame) m_frameFolderPaths[explorerFrame] = normalizedFolderPath;
    }

    if (explorerFrame) TrySubclassFrame(explorerFrame);
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
    // Must be called from the UI thread (same thread that owns the windows).
    auto info = FindDefViewAndDUI(explorerFrame);
    LogMessage(LogLevel::Info, L"FolderBg: TrySubclassFrame frame=%p defView=%p dui=%p",
               explorerFrame, info.defView, info.duiHwnd);

    if (info.defView) {
        SubclassHwnd(info.defView, explorerFrame, /*id=*/1);
    }
    if (info.duiHwnd) {
        SubclassHwnd(info.duiHwnd, explorerFrame, /*id=*/2);
    }
}

void FolderBackgroundHooks::SubclassHwnd(HWND hwnd, HWND explorerFrame, UINT_PTR id) {
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (m_subclassedWindows.count(hwnd)) return;  // already subclassed
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

    // Trigger deferred repaint so background is visible.
    // Do NOT use RDW_UPDATENOW here: forcing synchronous paint immediately after
    // SetWindowSubclass can cause DefSubclassProc to dispatch WM_ERASEBKGND with
    // wParam=0 during window initialization, leading to a null-HDC crash in DrawBackground.
    InvalidateRect(hwnd, nullptr, TRUE);
}

void FolderBackgroundHooks::UnsubclassAll() {
    std::unordered_set<HWND> toRemove;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        toRemove = m_subclassedWindows;
        m_subclassedWindows.clear();
    }
    // Iterate and remove subclass. We try both id=1 and id=2 since each window
    // only has one subclass installed.
    for (HWND hwnd : toRemove) {
        if (!IsWindow(hwnd)) continue;  // Window already destroyed; skip to avoid corrupting a recycled HWND
        RemoveWindowSubclass(hwnd, SubclassProc, 1);
        RemoveWindowSubclass(hwnd, SubclassProc, 2);
    }
}

// ============================================================================
// SubclassProc
// ============================================================================

LRESULT CALLBACK FolderBackgroundHooks::SubclassProc(
    HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
    UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    // IMPORTANT: No global re-entrancy guard here.
    //
    // The previous version used a `static thread_local bool` guard that blocked
    // ALL messages when re-entrant.  This prevented WM_ERASEBKGND from executing
    // during the WM_PAINT cycle (DefSubclassProc → BeginPaint → WM_ERASEBKGND),
    // which caused:
    //   1. Background image never visible (default erase overwrites it).
    //   2. WM_NCDESTROY cleanup skipped → stale HWNDs in m_subclassedWindows →
    //      UnsubclassAll could corrupt a recycled HWND → crash.
    //
    // DrawBackground has its own thread-local re-entrancy guard (`s_drawing`)
    // which is sufficient to prevent infinite recursion within the drawing path.

    // Always handle WM_NCDESTROY first — must never be skipped.
    if (msg == WM_NCDESTROY) {
        RemoveWindowSubclass(hwnd, SubclassProc, uIdSubclass);
        FolderBackgroundHooks& inst = Instance();
        {
            std::lock_guard<std::mutex> lock(inst.m_mutex);
            inst.m_subclassedWindows.erase(hwnd);
        }
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    HWND explorerFrame = reinterpret_cast<HWND>(dwRefData);
    FolderBackgroundHooks& inst = Instance();

    switch (msg) {
    case WM_ERASEBKGND:
        // Do NOT draw the background image here.
        //
        // DirectUIHWND uses DirectUI's internal rendering pipeline, which fills
        // its own background during WM_PAINT — overwriting anything we draw in
        // WM_ERASEBKGND.  The result: image invisible except for brief flashes
        // during scrolling (the gap between our draw and DirectUI's overwrite).
        //
        // Drawing is done after WM_PAINT instead (below).
        break;

    case WM_PAINT: {
        // Let DirectUI's WM_PAINT handler run to completion first.  This calls
        // BeginPaint → renders the entire visual tree (background, icons, text,
        // selection) → EndPaint.  The update region is now validated.
        LRESULT ret = DefSubclassProc(hwnd, msg, wParam, lParam);

        // Overlay our background image AFTER DirectUI has finished painting.
        // GetDC() returns a non-paint DC, but since EndPaint already validated
        // the update region, this won't trigger another WM_PAINT cycle.
        if (inst.m_enabled) {
            HDC hDC = GetDC(hwnd);
            if (hDC) {
                inst.DrawBackground(hDC, hwnd, explorerFrame);
                ReleaseDC(hwnd, hDC);
            }
        }
        return ret;
    }

    case WM_SIZE:
        // Redraw so image position updates on resize.
        if (inst.m_enabled) {
            InvalidateRect(hwnd, nullptr, TRUE);
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

    // We care about DirectUIHWND whose parent is SHELLDLL_DefView inside a CabinetWClass Explorer.
    if (GetWindowClass(hWnd) != L"DirectUIHWND") return hWnd;
    if (GetWindowClass(hWndParent) != L"SHELLDLL_DefView") return hWnd;

    HWND explorerFrame = FindExplorerFrame(hWndParent);
    if (!explorerFrame) return hWnd;

    FolderBackgroundHooks& inst = Instance();
    // Subclass both SHELLDLL_DefView and the new DirectUIHWND.
    inst.SubclassHwnd(hWndParent, explorerFrame, /*id=*/1);
    inst.SubclassHwnd(hWnd,       explorerFrame, /*id=*/2);

    return hWnd;
}

// ============================================================================
// Drawing
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

void FolderBackgroundHooks::DrawBackground(HDC hDC, HWND hWnd, HWND explorerFrame) {
    if (!hDC) return;
    // Guard against re-entrant calls (e.g. GDI operations in SetImages triggering messages).
    static thread_local bool s_drawing = false;
    if (s_drawing) return;
    s_drawing = true;
    struct DrawGuard { ~DrawGuard() { s_drawing = false; } } guard;

    // Snapshot all state we need, then draw without holding the lock.
    std::shared_ptr<const BackgroundBitmap> bmp;
    BackgroundPositionMode posMode;
    BYTE opacity;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (!m_enabled) return;
        bmp = ResolveBitmapForFrame(explorerFrame);
        if (!bmp) return;
        posMode = m_positionMode;
        opacity = m_opacity;
    }

    RECT cr;
    if (!GetClientRect(hWnd, &cr)) return;
    SIZE wndSize = {cr.right, cr.bottom};
    if (wndSize.cx <= 0 || wndSize.cy <= 0) return;

    SIZE imgSize = bmp->GetSize();
    if (imgSize.cx <= 0 || imgSize.cy <= 0) return;

    BLENDFUNCTION bf = {};
    bf.BlendOp             = AC_SRC_OVER;
    bf.SourceConstantAlpha = opacity;
    bf.AlphaFormat         = AC_SRC_ALPHA;  // source has pre-multiplied alpha

    // Create a temporary DC on the current (drawing) thread.
    // HBITMAP is cross-thread safe; HDC is not — this is why we don't cache the DC.
    HDC srcDC = CreateCompatibleDC(hDC);
    if (!srcDC) return;
    HGDIOBJ oldBmp = SelectObject(srcDC, bmp->GetBitmap());
    if (!oldBmp || oldBmp == HGDI_ERROR) {
        DeleteDC(srcDC);
        return;
    }

    if (posMode == BackgroundPositionMode::kTile) {
        // Tile: repeat the image across the entire client area
        for (int ty = 0; ty < wndSize.cy; ty += imgSize.cy) {
            for (int tx = 0; tx < wndSize.cx; tx += imgSize.cx) {
                AlphaBlend(hDC, tx, ty, imgSize.cx, imgSize.cy,
                           srcDC, 0, 0, imgSize.cx, imgSize.cy, bf);
            }
        }
    } else {
        POINT pos;
        SIZE  dstSize;
        CalculateImagePosition(wndSize, imgSize, posMode, pos, dstSize);
        if (dstSize.cx > 0 && dstSize.cy > 0) {
            AlphaBlend(hDC, pos.x, pos.y, dstSize.cx, dstSize.cy,
                       srcDC, 0, 0, imgSize.cx, imgSize.cy, bf);
        }
    }

    SelectObject(srcDC, oldBmp);
    DeleteDC(srcDC);

    LogMessage(LogLevel::Verbose, L"FolderBg: drew background hwnd=%p frame=%p mode=%d",
               hWnd, explorerFrame, static_cast<int>(posMode));
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
