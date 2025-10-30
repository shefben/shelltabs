#include "PreviewCache.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <mutex>
#include <optional>
#include <vector>

#include "Utilities.h"

namespace shelltabs {

namespace {
SIZE ComputeScaledSize(int srcWidth, int srcHeight, const SIZE& desiredSize) {
    SIZE result{srcWidth, srcHeight};
    if (srcWidth <= 0 || srcHeight <= 0) {
        return result;
    }
    if (desiredSize.cx <= 0 || desiredSize.cy <= 0) {
        return result;
    }

    const double srcAspect = static_cast<double>(srcWidth) / static_cast<double>(srcHeight);
    const double destAspect = static_cast<double>(desiredSize.cx) / static_cast<double>(desiredSize.cy);
    if (!std::isfinite(srcAspect) || !std::isfinite(destAspect) || srcAspect <= 0.0) {
        return desiredSize;
    }

    if (srcAspect > destAspect) {
        result.cx = desiredSize.cx;
        result.cy = std::max(1, static_cast<int>(std::round(desiredSize.cx / srcAspect)));
    } else {
        result.cy = desiredSize.cy;
        result.cx = std::max(1, static_cast<int>(std::round(desiredSize.cy * srcAspect)));
    }
    return result;
}

void EnsureOpaqueAlpha(void* bits, int width, int height) {
    if (!bits || width <= 0 || height <= 0) {
        return;
    }

    const int stride = width * 4;
    auto* row = static_cast<std::uint8_t*>(bits);
    for (int y = 0; y < height; ++y) {
        std::uint8_t* pixel = row + y * stride;
        for (int x = 0; x < width; ++x) {
            pixel[3] = 0xFF;
            pixel += 4;
        }
    }
}

bool RenderWindowToDc(HWND window, HDC dc, int width, int height) {
    if (!window || !dc || width <= 0 || height <= 0) {
        return false;
    }

    auto attemptPrintWindow = [&](HWND target, const POINT& offset) {
        if (!target) {
            return false;
        }

        const int state = SaveDC(dc);
        BOOL printed = FALSE;
        if (state != 0) {
            SetViewportOrgEx(dc, -offset.x, -offset.y, nullptr);
            printed = PrintWindow(target, dc, PW_RENDERFULLCONTENT);
            RestoreDC(dc, state);
        } else {
            POINT oldOrigin{};
            SetViewportOrgEx(dc, -offset.x, -offset.y, &oldOrigin);
            printed = PrintWindow(target, dc, PW_RENDERFULLCONTENT);
            SetViewportOrgEx(dc, oldOrigin.x, oldOrigin.y, nullptr);
        }

        return printed != FALSE;
    };

    POINT origin{0, 0};
    if (attemptPrintWindow(window, origin)) {
        return true;
    }

    const HWND root = GetAncestor(window, GA_ROOT);
    if (root && root != window) {
        RECT childRect{};
        RECT rootRect{};
        if (GetWindowRect(window, &childRect) && GetWindowRect(root, &rootRect)) {
            POINT offset{childRect.left - rootRect.left, childRect.top - rootRect.top};
            if (attemptPrintWindow(root, offset)) {
                return true;
            }
        }
    }

    DWORD_PTR result = 0;
    if (SendMessageTimeoutW(window, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc),
                            PRF_CLIENT | PRF_CHILDREN | PRF_OWNED, SMTO_NORMAL, 200,
                            reinterpret_cast<PDWORD_PTR>(&result))) {
        return true;
    }

    HDC windowDc = GetDC(window);
    if (!windowDc) {
        return false;
    }
    const BOOL copied = BitBlt(dc, 0, 0, width, height, windowDc, 0, 0, SRCCOPY);
    ReleaseDC(window, windowDc);
    return copied != FALSE;
}

}  // namespace

PreviewCache& PreviewCache::Instance() {
    static PreviewCache cache;
    return cache;
}

PreviewCache::~PreviewCache() {
    Clear();
}

std::optional<PreviewImage> PreviewCache::GetPreview(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize) {
    if (!pidl) {
        return std::nullopt;
    }

    UNREFERENCED_PARAMETER(desiredSize);

    const std::wstring key = BuildCacheKey(pidl);
    if (key.empty()) {
        return std::nullopt;
    }

    std::scoped_lock lock(m_mutex);
    auto it = m_entries.find(key);
    if (it == m_entries.end() || !it->second.bitmap) {
        return std::nullopt;
    }

    it->second.lastAccess = GetTickCount64();
    return PreviewImage{it->second.bitmap, it->second.size};
}

void PreviewCache::StorePreviewFromWindow(PCIDLIST_ABSOLUTE pidl, HWND window, const SIZE& desiredSize) {
    if (!pidl || !window || !IsWindow(window)) {
        return;
    }

    const std::wstring key = BuildCacheKey(pidl);
    if (key.empty()) {
        return;
    }

    RECT client{};
    if (!GetClientRect(window, &client)) {
        return;
    }

    const int srcWidth = client.right - client.left;
    const int srcHeight = client.bottom - client.top;
    if (srcWidth <= 0 || srcHeight <= 0) {
        return;
    }

    const SIZE targetSize = ComputeScaledSize(srcWidth, srcHeight, desiredSize);

    HDC screenDc = GetDC(nullptr);
    if (!screenDc) {
        return;
    }

    HDC srcDc = nullptr;
    HBITMAP srcBitmap = nullptr;
    void* srcBits = nullptr;
    HDC destDc = nullptr;
    HBITMAP destBitmap = nullptr;
    void* destBits = nullptr;
    HBITMAP finalBitmap = nullptr;
    SIZE finalSize{0, 0};
    HGDIOBJ oldSrc = nullptr;
    HGDIOBJ oldDest = nullptr;
    bool success = false;

    do {
        BITMAPINFO srcInfo{};
        srcInfo.bmiHeader.biSize = sizeof(srcInfo.bmiHeader);
        srcInfo.bmiHeader.biWidth = srcWidth;
        srcInfo.bmiHeader.biHeight = -srcHeight;
        srcInfo.bmiHeader.biPlanes = 1;
        srcInfo.bmiHeader.biBitCount = 32;
        srcInfo.bmiHeader.biCompression = BI_RGB;

        srcBitmap = CreateDIBSection(screenDc, &srcInfo, DIB_RGB_COLORS, &srcBits, nullptr, 0);
        if (!srcBitmap || !srcBits) {
            break;
        }

        srcDc = CreateCompatibleDC(screenDc);
        if (!srcDc) {
            break;
        }

        oldSrc = SelectObject(srcDc, srcBitmap);
        if (!RenderWindowToDc(window, srcDc, srcWidth, srcHeight)) {
            break;
        }

        EnsureOpaqueAlpha(srcBits, srcWidth, srcHeight);

        if (targetSize.cx != srcWidth || targetSize.cy != srcHeight) {
            BITMAPINFO destInfo = srcInfo;
            destInfo.bmiHeader.biWidth = targetSize.cx;
            destInfo.bmiHeader.biHeight = -targetSize.cy;

            destBitmap = CreateDIBSection(screenDc, &destInfo, DIB_RGB_COLORS, &destBits, nullptr, 0);
            if (!destBitmap || !destBits) {
                break;
            }

            destDc = CreateCompatibleDC(screenDc);
            if (!destDc) {
                break;
            }

            oldDest = SelectObject(destDc, destBitmap);
            SetStretchBltMode(destDc, HALFTONE);
            SetBrushOrgEx(destDc, 0, 0, nullptr);
            if (!StretchBlt(destDc, 0, 0, targetSize.cx, targetSize.cy, srcDc, 0, 0, srcWidth, srcHeight, SRCCOPY)) {
                break;
            }

            EnsureOpaqueAlpha(destBits, targetSize.cx, targetSize.cy);
            success = true;
            finalBitmap = destBitmap;
            destBitmap = nullptr;
            finalSize = targetSize;
        } else {
            success = true;
            finalBitmap = srcBitmap;
            srcBitmap = nullptr;
            finalSize = {srcWidth, srcHeight};
        }
    } while (false);

    if (oldDest && destDc) {
        SelectObject(destDc, oldDest);
    }
    if (destDc) {
        DeleteDC(destDc);
    }
    if (destBitmap) {
        DeleteObject(destBitmap);
    }
    if (oldSrc && srcDc) {
        SelectObject(srcDc, oldSrc);
    }
    if (srcDc) {
        DeleteDC(srcDc);
    }
    if (srcBitmap) {
        DeleteObject(srcBitmap);
    }
    ReleaseDC(nullptr, screenDc);

    if (!success || !finalBitmap) {
        if (finalBitmap) {
            DeleteObject(finalBitmap);
        }
        return;
    }

    std::scoped_lock lock(m_mutex);
    Entry& entry = m_entries[key];
    if (entry.bitmap) {
        DeleteObject(entry.bitmap);
    }
    entry.bitmap = finalBitmap;
    entry.size = finalSize;
    entry.lastAccess = GetTickCount64();
    TrimCacheLocked();
}

void PreviewCache::Clear() {
    std::scoped_lock lock(m_mutex);
    for (auto& [_, entry] : m_entries) {
        if (entry.bitmap) {
            DeleteObject(entry.bitmap);
            entry.bitmap = nullptr;
        }
    }
    m_entries.clear();
}

void PreviewCache::TrimCacheLocked() {
    if (m_entries.size() <= kMaxEntries) {
        return;
    }

    std::vector<std::pair<std::wstring, Entry*>> ordered;
    ordered.reserve(m_entries.size());
    for (auto& [key, entry] : m_entries) {
        ordered.emplace_back(key, &entry);
    }

    std::sort(ordered.begin(), ordered.end(), [](const auto& left, const auto& right) {
        return left.second->lastAccess < right.second->lastAccess;
    });

    while (m_entries.size() > kMaxEntries && !ordered.empty()) {
        const auto& victim = ordered.front();
        auto it = m_entries.find(victim.first);
        if (it != m_entries.end()) {
            if (it->second.bitmap) {
                DeleteObject(it->second.bitmap);
            }
            m_entries.erase(it);
        }
        ordered.erase(ordered.begin());
    }
}

std::wstring PreviewCache::BuildCacheKey(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return {};
    }
    std::wstring key = GetParsingName(pidl);
    if (!key.empty()) {
        return key;
    }
    wchar_t buffer[32];
    _snwprintf_s(buffer, _TRUNCATE, L"pidl:%p", pidl);
    return buffer;
}

}  // namespace shelltabs

