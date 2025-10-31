#include "PreviewCache.h"

#include <algorithm>
#include <atomic>
#include <cmath>
#include <cstdint>
#include <mutex>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "Utilities.h"

namespace shelltabs {

struct PreviewCache::AsyncRequest {
    enum class Kind {
        kShellPreview,
        kWindowCapture,
    };

    uint64_t id = 0;
    std::wstring key;
    UniquePidl pidl;
    SIZE size{};
    HWND notify = nullptr;
    UINT message = 0;
    std::atomic<bool> cancelled{false};
    Kind kind = Kind::kShellPreview;
    HWND window = nullptr;
    std::wstring ownerToken;
};

namespace {
constexpr size_t kMaxPendingCaptureRequests = 8;

HBITMAP LoadShellItemPreview(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize, SIZE* outSize) {
    if (!pidl) {
        return nullptr;
    }

    Microsoft::WRL::ComPtr<IShellItem> item;
    if (FAILED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&item))) || !item) {
        return nullptr;
    }
    Microsoft::WRL::ComPtr<IShellItemImageFactory> factory;
    if (FAILED(item.As(&factory)) || !factory) {
        return nullptr;
    }

    SIZE requestSize = desiredSize;
    if (requestSize.cx <= 0 || requestSize.cy <= 0) {
        requestSize = kPreviewImageSize;
    }

    HBITMAP bitmap = nullptr;
    HRESULT hr = factory->GetImage(requestSize,
                                   SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK | SIIGBF_THUMBNAILONLY,
                                   &bitmap);
    if (FAILED(hr)) {
        hr = factory->GetImage(requestSize, SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK, &bitmap);
    }
    if (FAILED(hr)) {
        hr = factory->GetImage(requestSize, SIIGBF_ICONONLY, &bitmap);
    }
    if (FAILED(hr) || !bitmap) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return nullptr;
    }

    BITMAP bmp{};
    if (GetObject(bitmap, sizeof(bmp), &bmp) <= 0) {
        DeleteObject(bitmap);
        return nullptr;
    }
    if (outSize) {
        outSize->cx = bmp.bmWidth;
        outSize->cy = bmp.bmHeight;
    }
    return bitmap;
}

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

HBITMAP CaptureWindowPreview(HWND window, const SIZE& desiredSize, SIZE* outSize) {
    if (!window || !IsWindow(window)) {
        return nullptr;
    }

    RECT client{};
    if (!GetClientRect(window, &client)) {
        return nullptr;
    }

    const int srcWidth = client.right - client.left;
    const int srcHeight = client.bottom - client.top;
    if (srcWidth <= 0 || srcHeight <= 0) {
        return nullptr;
    }

    const SIZE targetSize = ComputeScaledSize(srcWidth, srcHeight, desiredSize);

    HDC screenDc = GetDC(nullptr);
    if (!screenDc) {
        return nullptr;
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
        return nullptr;
    }

    if (outSize) {
        *outSize = finalSize;
    }
    return finalBitmap;
}

}  // namespace

PreviewCache& PreviewCache::Instance() {
    static PreviewCache cache;
    return cache;
}

PreviewCache::~PreviewCache() {
    {
        std::unique_lock lock(m_requestMutex);
        m_shutdown = true;
    }
    m_requestCv.notify_all();
    if (m_workerThread.joinable()) {
        m_workerThread.join();
    }
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

void PreviewCache::StorePreviewFromWindow(PCIDLIST_ABSOLUTE pidl, HWND window, const SIZE& desiredSize,
                                          std::wstring_view ownerToken) {
    if (!pidl || !window || !IsWindow(window)) {
        return;
    }

    const std::wstring key = BuildCacheKey(pidl);
    if (key.empty()) {
        return;
    }

    UniquePidl clone = ClonePidl(pidl);
    if (!clone) {
        return;
    }

    EnsureWorkerThread();

    auto request = std::make_shared<AsyncRequest>();
    request->kind = AsyncRequest::Kind::kWindowCapture;
    request->key = key;
    request->pidl = std::move(clone);
    request->size = desiredSize;
    request->window = window;
    if (!ownerToken.empty()) {
        request->ownerToken.assign(ownerToken);
    }

    {
        std::scoped_lock lock(m_requestMutex);

        for (auto it = m_requestQueue.begin(); it != m_requestQueue.end();) {
            const auto& pending = *it;
            if (!pending || pending->kind != AsyncRequest::Kind::kWindowCapture) {
                ++it;
                continue;
            }
            const bool sameKey = pending->key == request->key;
            const bool sameOwner = !request->ownerToken.empty() && pending->ownerToken == request->ownerToken;
            if (sameKey || sameOwner) {
                pending->cancelled.store(true, std::memory_order_release);
                m_requestMap.erase(pending->id);
                it = m_requestQueue.erase(it);
            } else {
                ++it;
            }
        }

        for (auto& [_, pending] : m_requestMap) {
            if (pending && pending->kind == AsyncRequest::Kind::kWindowCapture) {
                if (pending->key == request->key ||
                    (!request->ownerToken.empty() && pending->ownerToken == request->ownerToken)) {
                    pending->cancelled.store(true, std::memory_order_release);
                }
            }
        }

        size_t pendingCaptures = 0;
        for (const auto& pending : m_requestQueue) {
            if (pending && pending->kind == AsyncRequest::Kind::kWindowCapture) {
                ++pendingCaptures;
            }
        }
        while (pendingCaptures >= kMaxPendingCaptureRequests) {
            auto eraseIt = std::find_if(m_requestQueue.begin(), m_requestQueue.end(), [](const auto& candidate) {
                return candidate && candidate->kind == AsyncRequest::Kind::kWindowCapture;
            });
            if (eraseIt == m_requestQueue.end()) {
                break;
            }
            (*eraseIt)->cancelled.store(true, std::memory_order_release);
            m_requestMap.erase((*eraseIt)->id);
            eraseIt = m_requestQueue.erase(eraseIt);
            if (pendingCaptures > 0) {
                --pendingCaptures;
            }
        }

        request->id = m_nextRequestId++;
        if (m_nextRequestId == 0 || m_nextRequestId > 0xFFFFFFFFULL) {
            m_nextRequestId = 1;
        }

        m_requestQueue.push_back(request);
        m_requestMap[request->id] = request;
    }

    m_requestCv.notify_one();
}

uint64_t PreviewCache::RequestPreviewAsync(PCIDLIST_ABSOLUTE pidl, const SIZE& desiredSize, HWND notifyHwnd, UINT message) {
    if (!pidl) {
        return 0;
    }
    const std::wstring key = BuildCacheKey(pidl);
    if (key.empty()) {
        return 0;
    }
    UniquePidl clone = ClonePidl(pidl);
    if (!clone) {
        return 0;
    }

    EnsureWorkerThread();

    auto request = std::make_shared<AsyncRequest>();
    {
        std::scoped_lock lock(m_requestMutex);
        request->id = m_nextRequestId++;
        if (m_nextRequestId == 0 || m_nextRequestId > 0xFFFFFFFFULL) {
            m_nextRequestId = 1;
        }
        request->kind = AsyncRequest::Kind::kShellPreview;
        request->key = key;
        request->pidl = std::move(clone);
        request->size = desiredSize;
        request->notify = notifyHwnd;
        request->message = message;
        m_requestQueue.push_back(request);
        m_requestMap[request->id] = request;
    }
    m_requestCv.notify_one();
    return request->id;
}

void PreviewCache::CancelRequest(uint64_t requestId) {
    if (requestId == 0) {
        return;
    }
    std::scoped_lock lock(m_requestMutex);
    auto it = m_requestMap.find(requestId);
    if (it == m_requestMap.end()) {
        return;
    }
    it->second->cancelled.store(true, std::memory_order_release);
    it->second->notify = nullptr;
    it->second->message = 0;
    for (auto queueIt = m_requestQueue.begin(); queueIt != m_requestQueue.end(); ++queueIt) {
        if ((*queueIt)->id == requestId) {
            m_requestQueue.erase(queueIt);
            m_requestMap.erase(it);
            return;
        }
    }
}

void PreviewCache::CancelPendingCapturesForKey(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return;
    }
    const std::wstring key = BuildCacheKey(pidl);
    if (key.empty()) {
        return;
    }

    std::scoped_lock lock(m_requestMutex);
    for (auto it = m_requestQueue.begin(); it != m_requestQueue.end();) {
        const auto& pending = *it;
        if (pending && pending->kind == AsyncRequest::Kind::kWindowCapture && pending->key == key) {
            pending->cancelled.store(true, std::memory_order_release);
            m_requestMap.erase(pending->id);
            it = m_requestQueue.erase(it);
        } else {
            ++it;
        }
    }
    for (auto& [_, pending] : m_requestMap) {
        if (pending && pending->kind == AsyncRequest::Kind::kWindowCapture && pending->key == key) {
            pending->cancelled.store(true, std::memory_order_release);
        }
    }
}

void PreviewCache::CancelPendingCapturesForOwner(std::wstring_view ownerToken) {
    if (ownerToken.empty()) {
        return;
    }

    std::scoped_lock lock(m_requestMutex);
    for (auto it = m_requestQueue.begin(); it != m_requestQueue.end();) {
        const auto& pending = *it;
        if (pending && pending->kind == AsyncRequest::Kind::kWindowCapture && pending->ownerToken == ownerToken) {
            pending->cancelled.store(true, std::memory_order_release);
            m_requestMap.erase(pending->id);
            it = m_requestQueue.erase(it);
        } else {
            ++it;
        }
    }
    for (auto& [_, pending] : m_requestMap) {
        if (pending && pending->kind == AsyncRequest::Kind::kWindowCapture && pending->ownerToken == ownerToken) {
            pending->cancelled.store(true, std::memory_order_release);
        }
    }
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

void PreviewCache::EnsureWorkerThread() {
    std::scoped_lock lock(m_requestMutex);
    if (m_workerThread.joinable()) {
        return;
    }
    m_shutdown = false;
    m_workerThread = std::thread([this]() { ProcessRequests(); });
}

void PreviewCache::ProcessRequests() {
    const HRESULT coInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    while (true) {
        std::shared_ptr<AsyncRequest> request;
        {
            std::unique_lock lock(m_requestMutex);
            m_requestCv.wait(lock, [this]() { return m_shutdown || !m_requestQueue.empty(); });
            if (m_shutdown && m_requestQueue.empty()) {
                break;
            }
            request = m_requestQueue.front();
            m_requestQueue.pop_front();
        }
        if (!request) {
            continue;
        }

        const bool cancelledBeforeWork = request->cancelled.load(std::memory_order_acquire);

        SIZE generatedSize{};
        HBITMAP bitmap = nullptr;
        if (!cancelledBeforeWork) {
            if (request->kind == AsyncRequest::Kind::kShellPreview) {
                bitmap = LoadShellItemPreview(request->pidl.get(), request->size, &generatedSize);
            } else {
                bitmap = CaptureWindowPreview(request->window, request->size, &generatedSize);
            }
        }

        const bool cancelledAfterWork = request->cancelled.load(std::memory_order_acquire);
        bool stored = false;
        if (!cancelledAfterWork && bitmap) {
            StoreBitmapForKey(request->key, bitmap, generatedSize);
            stored = true;
        }
        if (!stored && bitmap) {
            DeleteObject(bitmap);
            bitmap = nullptr;
        }

        const HWND target = request->notify;
        const UINT message = request->message;
        const uint64_t id = request->id;

        {
            std::scoped_lock lock(m_requestMutex);
            m_requestMap.erase(id);
        }

        if (request->kind == AsyncRequest::Kind::kShellPreview && !cancelledAfterWork && target && message != 0) {
            PostMessageW(target, message, static_cast<WPARAM>(id), 0);
        }
    }
    if (SUCCEEDED(coInit)) {
        CoUninitialize();
    }
}

void PreviewCache::StoreBitmapForKey(const std::wstring& key, HBITMAP bitmap, const SIZE& size) {
    if (key.empty() || !bitmap) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return;
    }
    std::scoped_lock lock(m_mutex);
    Entry& entry = m_entries[key];
    if (entry.bitmap) {
        DeleteObject(entry.bitmap);
    }
    entry.bitmap = bitmap;
    entry.size = size;
    entry.lastAccess = GetTickCount64();
    TrimCacheLocked();
}

}  // namespace shelltabs

