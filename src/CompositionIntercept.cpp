#include "CompositionIntercept.h"

#include <MinHook.h>

#include <dcomp.h>
#include <d2d1_1.h>
#include <d2d1_1helper.h>
#include <wrl/client.h>

#include <algorithm>
#include <cmath>
#include <mutex>
#include <new>
#include <optional>
#include <unordered_map>
#include <unordered_set>

#include "Logging.h"

namespace shelltabs {
namespace {

using Microsoft::WRL::ComPtr;

struct HwndHasher {
    size_t operator()(HWND hwnd) const noexcept {
        return reinterpret_cast<size_t>(hwnd);
    }
};

struct PointerHasher {
    size_t operator()(const void* pointer) const noexcept {
        return reinterpret_cast<size_t>(pointer);
    }
};

using DCompositionCreateDeviceFn = HRESULT(WINAPI*)(IUnknown*, REFIID, void**);
using DeviceCreateTargetForHwndFn =
    HRESULT(STDMETHODCALLTYPE*)(IDCompositionDevice*, HWND, BOOL, IDCompositionTarget**);
using TargetSetRootFn = HRESULT(STDMETHODCALLTYPE*)(IDCompositionTarget*, IDCompositionVisual*);
using TargetReleaseFn = ULONG(STDMETHODCALLTYPE*)(IDCompositionTarget*);
using GdiAlphaBlendFn = BOOL(WINAPI*)(HDC, int, int, int, int, HDC, int, int, int, int, BLENDFUNCTION);

void* g_createDeviceAddress = nullptr;
void* g_createDevice2Address = nullptr;
void* g_createDevice3Address = nullptr;

DCompositionCreateDeviceFn g_originalCreateDevice = nullptr;
DCompositionCreateDeviceFn g_originalCreateDevice2 = nullptr;
DCompositionCreateDeviceFn g_originalCreateDevice3 = nullptr;

DeviceCreateTargetForHwndFn g_originalDeviceCreateTargetForHwnd = nullptr;
TargetSetRootFn g_originalTargetSetRoot = nullptr;
TargetReleaseFn g_originalTargetRelease = nullptr;
GdiAlphaBlendFn g_originalGdiAlphaBlend = nullptr;

void* g_deviceCreateTargetForHwndAddress = nullptr;
void* g_targetSetRootAddress = nullptr;
void* g_targetReleaseAddress = nullptr;
void* g_gdiAlphaBlendAddress = nullptr;

bool g_hooksInstalled = false;

ComPtr<ID2D1Factory1> g_d2dFactory;
std::once_flag g_d2dFactoryInit;

struct CompositionRegistration {
    ExplorerGlowCoordinator* coordinator = nullptr;
    size_t refCount = 0;
};

struct ExplorerTargetContext : public std::enable_shared_from_this<ExplorerTargetContext> {
    HWND hwnd = nullptr;
    ExplorerGlowCoordinator* coordinator = nullptr;
    ComPtr<IDCompositionDevice> deviceV1;
    ComPtr<IDCompositionDevice2> deviceV2;
    ComPtr<IDCompositionDevice3> deviceV3;
    ComPtr<IDCompositionTarget> target;
    ComPtr<IDCompositionVisual2> container;
    ComPtr<IDCompositionVisual2> backgroundVisual;
    ComPtr<IDCompositionVisual2> foregroundVisual;
    ComPtr<IDCompositionSurface> backgroundSurface;
    ComPtr<IDCompositionSurface> foregroundSurface;
    GlowColorSet lastColors{};
    SIZE lastSize{0, 0};
    bool surfacesDirty = true;
    std::once_flag firstWrapLog;
    std::mutex mutex;
    HBITMAP backgroundBitmap = nullptr;
    HBITMAP foregroundBitmap = nullptr;

    ~ExplorerTargetContext() {
        if (backgroundBitmap) {
            DeleteObject(backgroundBitmap);
            backgroundBitmap = nullptr;
        }
        if (foregroundBitmap) {
            DeleteObject(foregroundBitmap);
            foregroundBitmap = nullptr;
        }
    }
};

std::mutex g_registrationMutex;
std::unordered_map<HWND, CompositionRegistration, HwndHasher> g_registrations;

std::mutex g_contextMutex;
std::unordered_map<IDCompositionTarget*, std::shared_ptr<ExplorerTargetContext>, PointerHasher> g_targetContexts;
std::unordered_map<HWND, std::weak_ptr<ExplorerTargetContext>, HwndHasher> g_windowContexts;

std::mutex g_bitmapMutex;
std::unordered_map<HBITMAP, std::weak_ptr<ExplorerTargetContext>, PointerHasher> g_trackedBitmaps;

bool EnsureD2DFactory() {
    bool created = false;
    std::call_once(g_d2dFactoryInit, [&]() {
        D2D1_FACTORY_OPTIONS options{};
#if defined(_DEBUG)
        options.debugLevel = D2D1_DEBUG_LEVEL_INFORMATION;
#endif
        ComPtr<ID2D1Factory1> factory;
        HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), &options,
                                       reinterpret_cast<void**>(factory.GetAddressOf()));
        if (SUCCEEDED(hr)) {
            g_d2dFactory = std::move(factory);
            created = true;
        } else {
            LogHrFailure(L"CompositionIntercept: D2D1CreateFactory", hr);
        }
    });
    return g_d2dFactory != nullptr;
}

HWND NormalizeWindow(HWND hwnd) noexcept {
    if (!hwnd) {
        return nullptr;
    }
    HWND root = GetAncestor(hwnd, GA_ROOT);
    return root ? root : hwnd;
}

bool IsExplorerWindow(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }
    wchar_t className[64] = {};
    if (!GetClassNameW(hwnd, className, ARRAYSIZE(className))) {
        return false;
    }
    return (_wcsicmp(className, L"CabinetWClass") == 0) || (_wcsicmp(className, L"ExploreWClass") == 0);
}

void TrackBitmap(HBITMAP bitmap, const std::shared_ptr<ExplorerTargetContext>& context) {
    if (!bitmap || !context) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_bitmapMutex);
    g_trackedBitmaps[bitmap] = context;
}

void UntrackBitmap(HBITMAP bitmap) {
    if (!bitmap) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_bitmapMutex);
    g_trackedBitmaps.erase(bitmap);
}

std::shared_ptr<ExplorerTargetContext> LookupBitmapContext(HBITMAP bitmap) {
    if (!bitmap) {
        return nullptr;
    }
    std::lock_guard<std::mutex> guard(g_bitmapMutex);
    auto it = g_trackedBitmaps.find(bitmap);
    if (it == g_trackedBitmaps.end()) {
        return nullptr;
    }
    return it->second.lock();
}

uint32_t PremultiplyChannel(uint8_t value, uint8_t alpha) noexcept {
    return static_cast<uint32_t>(value) * static_cast<uint32_t>(alpha) / 255u;
}

void FillGradientPixels(uint32_t* pixels, int width, int height, const GlowColorSet& colors) {
    if (!pixels || width <= 0 || height <= 0) {
        return;
    }

    const uint8_t alpha = 0xFF;
    const uint8_t startR = GetRValue(colors.start);
    const uint8_t startG = GetGValue(colors.start);
    const uint8_t startB = GetBValue(colors.start);
    const uint8_t endR = GetRValue(colors.end);
    const uint8_t endG = GetGValue(colors.end);
    const uint8_t endB = GetBValue(colors.end);

    for (int y = 0; y < height; ++y) {
        double ratio = static_cast<double>(y) / static_cast<double>(std::max(1, height - 1));
        uint8_t r = startR;
        uint8_t g = startG;
        uint8_t b = startB;
        if (colors.gradient) {
            r = static_cast<uint8_t>(std::clamp(static_cast<int>(std::lround(startR + (endR - startR) * ratio)), 0, 255));
            g = static_cast<uint8_t>(std::clamp(static_cast<int>(std::lround(startG + (endG - startG) * ratio)), 0, 255));
            b = static_cast<uint8_t>(std::clamp(static_cast<int>(std::lround(startB + (endB - startB) * ratio)), 0, 255));
        }
        const uint32_t pixel = (alpha << 24) | (PremultiplyChannel(r, alpha) << 16) |
                               (PremultiplyChannel(g, alpha) << 8) | PremultiplyChannel(b, alpha);
        for (int x = 0; x < width; ++x) {
            pixels[y * width + x] = pixel;
        }
    }
}

HBITMAP CreateGradientBitmap(int width, int height, const GlowColorSet& colors) {
    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap || !bits) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return nullptr;
    }
    FillGradientPixels(static_cast<uint32_t*>(bits), width, height, colors);
    return bitmap;
}

void UpdateGradientBitmap(HBITMAP bitmap, int width, int height, const GlowColorSet& colors) {
    if (!bitmap) {
        return;
    }
    DIBSECTION section{};
    if (GetObjectW(bitmap, sizeof(section), &section) != sizeof(section)) {
        return;
    }
    if (!section.dsBm.bmBits || section.dsBm.bmWidth != width || section.dsBm.bmHeight != -height) {
        return;
    }
    FillGradientPixels(static_cast<uint32_t*>(section.dsBm.bmBits), width, height, colors);
}

HRESULT DrawGradientToSurface(IDCompositionSurface* surface, int width, int height, const GlowColorSet& colors) {
    if (!surface || width <= 0 || height <= 0) {
        return E_INVALIDARG;
    }
    if (!EnsureD2DFactory()) {
        return E_FAIL;
    }

    RECT updateRect{0, 0, width, height};
    POINT offset{};
    ComPtr<ID2D1DeviceContext> dc;
    HRESULT hr = surface->BeginDraw(&updateRect, __uuidof(ID2D1DeviceContext),
                                    reinterpret_cast<void**>(dc.GetAddressOf()), &offset);
    if (FAILED(hr)) {
        return hr;
    }

    dc->Clear(D2D1::ColorF(0.f, 0.f));

    D2D1_GRADIENT_STOP stops[2];
    stops[0].position = 0.0f;
    stops[0].color = D2D1::ColorF(static_cast<float>(GetRValue(colors.start)) / 255.f,
                                  static_cast<float>(GetGValue(colors.start)) / 255.f,
                                  static_cast<float>(GetBValue(colors.start)) / 255.f, 1.f);
    stops[1].position = 1.0f;
    stops[1].color = D2D1::ColorF(static_cast<float>(GetRValue(colors.end)) / 255.f,
                                  static_cast<float>(GetGValue(colors.end)) / 255.f,
                                  static_cast<float>(GetBValue(colors.end)) / 255.f, 1.f);

    ComPtr<ID2D1GradientStopCollection> collection;
    hr = dc->CreateGradientStopCollection(stops, 2, D2D1_GAMMA_2_2, D2D1_EXTEND_MODE_CLAMP, &collection);
    if (FAILED(hr)) {
        surface->EndDraw();
        return hr;
    }

    ComPtr<ID2D1LinearGradientBrush> brush;
    hr = dc->CreateLinearGradientBrush(D2D1::LinearGradientBrushProperties(D2D1::Point2F(0.f, 0.f),
                                                                          D2D1::Point2F(0.f, static_cast<float>(height))),
                                       collection.Get(), &brush);
    if (FAILED(hr)) {
        surface->EndDraw();
        return hr;
    }

    dc->FillRectangle(D2D1::RectF(0.f, 0.f, static_cast<float>(width), static_cast<float>(height)), brush.Get());
    hr = dc->EndDraw();
    surface->EndDraw();
    return hr;
}

void UpdateTrackedBitmapsLocked(const std::shared_ptr<ExplorerTargetContext>& context, int width, int height,
                                const GlowColorSet& colors) {
    if (!context || !colors.valid) {
        return;
    }

    if (context->backgroundBitmap) {
        UpdateGradientBitmap(context->backgroundBitmap, width, height, colors);
    } else {
        context->backgroundBitmap = CreateGradientBitmap(width, height, colors);
        if (context->backgroundBitmap) {
            TrackBitmap(context->backgroundBitmap, context);
        }
    }

    if (context->foregroundBitmap) {
        UpdateGradientBitmap(context->foregroundBitmap, width, height, colors);
    } else {
        context->foregroundBitmap = CreateGradientBitmap(width, height, colors);
        if (context->foregroundBitmap) {
            TrackBitmap(context->foregroundBitmap, context);
        }
    }
}

void UpdateGradientSurfacesLocked(const std::shared_ptr<ExplorerTargetContext>& context) {
    if (!context) {
        return;
    }
    RECT rect{};
    if (!context->hwnd || !IsWindow(context->hwnd) || !GetClientRect(context->hwnd, &rect)) {
        return;
    }
    const int width = std::max(1, static_cast<int>(rect.right - rect.left));
    const int height = std::max(1, static_cast<int>(rect.bottom - rect.top));

    GlowColorSet colors{};
    if (context->coordinator) {
        colors = context->coordinator->ResolveColors(ExplorerSurfaceKind::ListView);
    }
    if (!colors.valid) {
        return;
    }

    bool recreate = !context->backgroundSurface || context->lastSize.cx != width || context->lastSize.cy != height;

    if (recreate) {
        context->backgroundSurface.Reset();
        context->foregroundSurface.Reset();
        ComPtr<IDCompositionDevice3> device = context->deviceV3;
        if (!device && context->deviceV2) {
            context->deviceV2.As(&device);
        }
        if (!device && context->deviceV1) {
            context->deviceV1.As(&device);
        }
        if (!device) {
            return;
        }
        HRESULT hr = device->CreateSurface(width, height, DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_ALPHA_MODE_PREMULTIPLIED,
                                            &context->backgroundSurface);
        if (FAILED(hr)) {
            LogHrFailure(L"CompositionIntercept: CreateSurface(background)", hr);
            return;
        }
        hr = device->CreateSurface(width, height, DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_ALPHA_MODE_PREMULTIPLIED,
                                   &context->foregroundSurface);
        if (FAILED(hr)) {
            LogHrFailure(L"CompositionIntercept: CreateSurface(foreground)", hr);
            return;
        }
    }

    HRESULT hr = DrawGradientToSurface(context->backgroundSurface.Get(), width, height, colors);
    if (FAILED(hr)) {
        LogHrFailure(L"CompositionIntercept: DrawGradientToSurface(background)", hr);
        return;
    }
    hr = DrawGradientToSurface(context->foregroundSurface.Get(), width, height, colors);
    if (FAILED(hr)) {
        LogHrFailure(L"CompositionIntercept: DrawGradientToSurface(foreground)", hr);
        return;
    }

    if (context->backgroundVisual) {
        context->backgroundVisual->SetContent(context->backgroundSurface.Get());
    }
    if (context->foregroundVisual) {
        context->foregroundVisual->SetContent(context->foregroundSurface.Get());
    }

    context->lastColors = colors;
    context->lastSize.cx = width;
    context->lastSize.cy = height;
    context->surfacesDirty = false;

    UpdateTrackedBitmapsLocked(context, width, height, colors);
}

HRESULT InstallHook(void* target, void* detour, void** original, const wchar_t* name) {
    if (!target || !detour) {
        return E_POINTER;
    }

    MH_STATUS status = MH_CreateHook(target, detour, original);
    if (status != MH_OK) {
        LogMessage(LogLevel::Error, L"CompositionIntercept: MH_CreateHook failed for %ls (status=%d)", name,
                   static_cast<int>(status));
        return E_FAIL;
    }
    status = MH_EnableHook(target);
    if (status != MH_OK) {
        LogMessage(LogLevel::Error, L"CompositionIntercept: MH_EnableHook failed for %ls (status=%d)", name,
                   static_cast<int>(status));
        MH_RemoveHook(target);
        return E_FAIL;
    }
    return S_OK;
}

void RemoveHook(void* target, const wchar_t* name) {
    if (!target) {
        return;
    }
    MH_STATUS status = MH_DisableHook(target);
    if (status != MH_OK && status != MH_ERROR_NOT_INITIALIZED) {
        LogMessage(LogLevel::Warning, L"CompositionIntercept: MH_DisableHook failed for %ls (status=%d)", name,
                   static_cast<int>(status));
    }
    status = MH_RemoveHook(target);
    if (status != MH_OK && status != MH_ERROR_NOT_INITIALIZED) {
        LogMessage(LogLevel::Warning, L"CompositionIntercept: MH_RemoveHook failed for %ls (status=%d)", name,
                   static_cast<int>(status));
    }
}

void InstallDeviceHooks(IDCompositionDevice* device) {
    if (!device || g_deviceCreateTargetForHwndAddress) {
        return;
    }
    void** vtable = *reinterpret_cast<void***>(device);
    if (!vtable) {
        return;
    }
    constexpr size_t kCreateTargetForHwndIndex = 6;
    void* target = vtable[kCreateTargetForHwndIndex];
    if (!target) {
        return;
    }
    if (FAILED(InstallHook(target, reinterpret_cast<void*>(&DeviceCreateTargetForHwndDetour),
                           reinterpret_cast<void**>(&g_originalDeviceCreateTargetForHwnd),
                           L"IDCompositionDevice::CreateTargetForHwnd"))) {
        return;
    }
    g_deviceCreateTargetForHwndAddress = target;
}

void InstallTargetHooks(IDCompositionTarget* target) {
    if (!target) {
        return;
    }
    void** vtable = *reinterpret_cast<void***>(target);
    if (!vtable) {
        return;
    }

    if (!g_targetSetRootAddress) {
        constexpr size_t kSetRootIndex = 3;
        void* setRoot = vtable[kSetRootIndex];
        if (setRoot &&
            SUCCEEDED(InstallHook(setRoot, reinterpret_cast<void*>(&TargetSetRootDetour),
                                  reinterpret_cast<void**>(&g_originalTargetSetRoot),
                                  L"IDCompositionTarget::SetRoot"))) {
            g_targetSetRootAddress = setRoot;
        }
    }
    if (!g_targetReleaseAddress) {
        constexpr size_t kReleaseIndex = 2;
        void* release = vtable[kReleaseIndex];
        if (release &&
            SUCCEEDED(InstallHook(release, reinterpret_cast<void*>(&TargetReleaseDetour),
                                  reinterpret_cast<void**>(&g_originalTargetRelease),
                                  L"IDCompositionTarget::Release"))) {
            g_targetReleaseAddress = release;
        }
    }
}

void RemoveContext(IDCompositionTarget* target) {
    std::shared_ptr<ExplorerTargetContext> context;
    {
        std::lock_guard<std::mutex> guard(g_contextMutex);
        auto it = g_targetContexts.find(target);
        if (it != g_targetContexts.end()) {
            context = it->second;
            g_targetContexts.erase(it);
        }
        if (context && context->hwnd) {
            auto wndIt = g_windowContexts.find(context->hwnd);
            if (wndIt != g_windowContexts.end()) {
                auto existing = wndIt->second.lock();
                if (!existing || existing.get() == context.get()) {
                    g_windowContexts.erase(wndIt);
                }
            }
        }
    }
    if (context) {
        if (context->backgroundBitmap) {
            UntrackBitmap(context->backgroundBitmap);
        }
        if (context->foregroundBitmap) {
            UntrackBitmap(context->foregroundBitmap);
        }
    }
}

HRESULT STDMETHODCALLTYPE DeviceCreateTargetForHwndDetour(IDCompositionDevice* self, HWND hwnd, BOOL topmost,
                                                          IDCompositionTarget** target) {
    HRESULT hr = g_originalDeviceCreateTargetForHwnd ?
                     g_originalDeviceCreateTargetForHwnd(self, hwnd, topmost, target)
                                                    : E_FAIL;
    if (FAILED(hr) || !target || !*target) {
        return hr;
    }

    HWND normalized = NormalizeWindow(hwnd);
    if (!IsExplorerWindow(normalized)) {
        return hr;
    }

    ExplorerGlowCoordinator* coordinator = nullptr;
    {
        std::lock_guard<std::mutex> guard(g_registrationMutex);
        auto it = g_registrations.find(normalized);
        if (it != g_registrations.end()) {
            coordinator = it->second.coordinator;
        }
    }
    if (!coordinator) {
        return hr;
    }

    InstallTargetHooks(*target);

    auto context = std::make_shared<ExplorerTargetContext>();
    context->hwnd = normalized;
    context->coordinator = coordinator;
    context->target = *target;

    if (self) {
        context->deviceV1 = self;
        self->QueryInterface(__uuidof(IDCompositionDevice2), reinterpret_cast<void**>(context->deviceV2.GetAddressOf()));
        self->QueryInterface(__uuidof(IDCompositionDevice3), reinterpret_cast<void**>(context->deviceV3.GetAddressOf()));
    }

    {
        std::lock_guard<std::mutex> guard(g_contextMutex);
        g_targetContexts[*target] = context;
        g_windowContexts[normalized] = context;
    }

    return hr;
}

HRESULT STDMETHODCALLTYPE TargetSetRootDetour(IDCompositionTarget* target, IDCompositionVisual* visual) {
    if (!g_originalTargetSetRoot) {
        return E_FAIL;
    }

    std::shared_ptr<ExplorerTargetContext> context;
    {
        std::lock_guard<std::mutex> guard(g_contextMutex);
        auto it = g_targetContexts.find(target);
        if (it != g_targetContexts.end()) {
            context = it->second;
        }
    }

    if (!context || !visual) {
        return g_originalTargetSetRoot(target, visual);
    }

    std::lock_guard<std::mutex> guard(context->mutex);

    ComPtr<IDCompositionVisual2> visual2;
    HRESULT hr = visual->QueryInterface(__uuidof(IDCompositionVisual2), reinterpret_cast<void**>(visual2.GetAddressOf()));
    if (FAILED(hr) || !visual2) {
        return g_originalTargetSetRoot(target, visual);
    }

    if (!context->deviceV3 && context->deviceV2) {
        context->deviceV2.As(&context->deviceV3);
    }
    if (!context->deviceV3 && context->deviceV1) {
        context->deviceV1.As(&context->deviceV3);
    }

    ComPtr<IDCompositionDevice3> device = context->deviceV3;
    if (!device) {
        return g_originalTargetSetRoot(target, visual);
    }

    if (!context->container) {
        hr = device->CreateVisual(&context->container);
        if (FAILED(hr)) {
            LogHrFailure(L"CompositionIntercept: CreateVisual(container)", hr);
            return g_originalTargetSetRoot(target, visual);
        }
    }
    if (!context->backgroundVisual) {
        hr = device->CreateVisual(&context->backgroundVisual);
        if (FAILED(hr)) {
            LogHrFailure(L"CompositionIntercept: CreateVisual(background)", hr);
            return g_originalTargetSetRoot(target, visual);
        }
    }
    if (!context->foregroundVisual) {
        hr = device->CreateVisual(&context->foregroundVisual);
        if (FAILED(hr)) {
            LogHrFailure(L"CompositionIntercept: CreateVisual(foreground)", hr);
            return g_originalTargetSetRoot(target, visual);
        }
    }

    context->container->RemoveAllVisuals();
    context->container->AddVisual(context->backgroundVisual.Get(), FALSE, nullptr);
    context->container->AddVisual(visual2.Get(), FALSE, context->backgroundVisual.Get());
    context->container->AddVisual(context->foregroundVisual.Get(), TRUE, visual2.Get());

    UpdateGradientSurfacesLocked(context);

    std::call_once(context->firstWrapLog, [&]() {
        LogMessage(LogLevel::Info, L"CompositionIntercept: Wrapped Explorer composition surface (hwnd=%p)",
                   context->hwnd);
    });

    hr = g_originalTargetSetRoot(target, context->container.Get());
    if (SUCCEEDED(hr)) {
        context->surfacesDirty = false;
    }
    return hr;
}

ULONG STDMETHODCALLTYPE TargetReleaseDetour(IDCompositionTarget* target) {
    ULONG result = g_originalTargetRelease ? g_originalTargetRelease(target) : 0;
    if (result == 0) {
        RemoveContext(target);
    }
    return result;
}

BOOL WINAPI GdiAlphaBlendDetour(HDC dest, int destX, int destY, int destW, int destH, HDC src, int srcX, int srcY,
                                int srcW, int srcH, BLENDFUNCTION blend) {
    if (src) {
        HBITMAP bitmap = reinterpret_cast<HBITMAP>(GetCurrentObject(src, OBJ_BITMAP));
        if (bitmap) {
            auto context = LookupBitmapContext(bitmap);
            if (context) {
                std::lock_guard<std::mutex> guard(context->mutex);
                if (context->surfacesDirty) {
                    UpdateGradientSurfacesLocked(context);
                }
                if (context->lastColors.valid) {
                    UpdateGradientBitmap(bitmap, context->lastSize.cx, context->lastSize.cy, context->lastColors);
                }
            }
        }
    }
    return g_originalGdiAlphaBlend ?
               g_originalGdiAlphaBlend(dest, destX, destY, destW, destH, src, srcX, srcY, srcW, srcH, blend)
                                   : FALSE;
}

HRESULT WrapDevice(IUnknown** object, REFIID iid) {
    if (!object || !*object) {
        return S_OK;
    }
    ComPtr<IUnknown> unknown;
    unknown.Attach(*object);

    ComPtr<IDCompositionDevice> device;
    if (FAILED(unknown.As(&device)) || !device) {
        ComPtr<IDCompositionDevice2> device2;
        if (SUCCEEDED(unknown.As(&device2)) && device2) {
            device2.As(&device);
        }
    }
    if (!device) {
        *object = unknown.Detach();
        return S_OK;
    }

    InstallDeviceHooks(device.Get());
    *object = unknown.Detach();
    return S_OK;
}

HRESULT WINAPI DCompositionCreateDeviceDetour(IUnknown* renderingDevice, REFIID iid, void** object) {
    HRESULT hr = g_originalCreateDevice ? g_originalCreateDevice(renderingDevice, iid, object) : E_FAIL;
    if (FAILED(hr)) {
        return hr;
    }
    return WrapDevice(reinterpret_cast<IUnknown**>(object), iid);
}

HRESULT WINAPI DCompositionCreateDevice2Detour(IUnknown* renderingDevice, REFIID iid, void** object) {
    HRESULT hr = g_originalCreateDevice2 ? g_originalCreateDevice2(renderingDevice, iid, object) : E_FAIL;
    if (FAILED(hr)) {
        return hr;
    }
    return WrapDevice(reinterpret_cast<IUnknown**>(object), iid);
}

HRESULT WINAPI DCompositionCreateDevice3Detour(IUnknown* renderingDevice, REFIID iid, void** object) {
    HRESULT hr = g_originalCreateDevice3 ? g_originalCreateDevice3(renderingDevice, iid, object) : E_FAIL;
    if (FAILED(hr)) {
        return hr;
    }
    return WrapDevice(reinterpret_cast<IUnknown**>(object), iid);
}

void ClearRegistration(HWND hwnd) {
    std::shared_ptr<ExplorerTargetContext> context;
    {
        std::lock_guard<std::mutex> guard(g_contextMutex);
        auto wndIt = g_windowContexts.find(hwnd);
        if (wndIt != g_windowContexts.end()) {
            context = wndIt->second.lock();
            g_windowContexts.erase(wndIt);
        }
    }
    if (!context) {
        return;
    }
    if (context->target) {
        RemoveContext(context->target.Get());
    }
}

}  // namespace

bool InitializeCompositionIntercept() {
    MH_STATUS status = MH_Initialize();
    if (status != MH_OK && status != MH_ERROR_ALREADY_INITIALIZED) {
        LogMessage(LogLevel::Error, L"CompositionIntercept: MH_Initialize failed (status=%d)", static_cast<int>(status));
        return false;
    }

    HMODULE dcomp = GetModuleHandleW(L"dcomp.dll");
    if (!dcomp) {
        dcomp = LoadLibraryW(L"dcomp.dll");
    }
    if (!dcomp) {
        LogLastError(L"CompositionIntercept: LoadLibrary(dcomp.dll)", GetLastError());
        return false;
    }

    HMODULE msimg = GetModuleHandleW(L"msimg32.dll");
    if (!msimg) {
        msimg = LoadLibraryW(L"msimg32.dll");
    }
    if (!msimg) {
        LogLastError(L"CompositionIntercept: LoadLibrary(msimg32.dll)", GetLastError());
        return false;
    }

    auto createDevice = reinterpret_cast<DCompositionCreateDeviceFn>(GetProcAddress(dcomp, "DCompositionCreateDevice"));
    auto createDevice2 = reinterpret_cast<DCompositionCreateDeviceFn>(GetProcAddress(dcomp, "DCompositionCreateDevice2"));
    auto createDevice3 = reinterpret_cast<DCompositionCreateDeviceFn>(GetProcAddress(dcomp, "DCompositionCreateDevice3"));
    auto alphaBlend = reinterpret_cast<GdiAlphaBlendFn>(GetProcAddress(msimg, "GdiAlphaBlend"));

    if (createDevice &&
        FAILED(InstallHook(reinterpret_cast<void*>(createDevice),
                           reinterpret_cast<void*>(&DCompositionCreateDeviceDetour),
                           reinterpret_cast<void**>(&g_originalCreateDevice), L"DCompositionCreateDevice"))) {
        return false;
    }
    if (createDevice) {
        g_createDeviceAddress = reinterpret_cast<void*>(createDevice);
    }
    if (createDevice2 &&
        FAILED(InstallHook(reinterpret_cast<void*>(createDevice2),
                           reinterpret_cast<void*>(&DCompositionCreateDevice2Detour),
                           reinterpret_cast<void**>(&g_originalCreateDevice2), L"DCompositionCreateDevice2"))) {
        return false;
    }
    if (createDevice2) {
        g_createDevice2Address = reinterpret_cast<void*>(createDevice2);
    }
    if (createDevice3 &&
        FAILED(InstallHook(reinterpret_cast<void*>(createDevice3),
                           reinterpret_cast<void*>(&DCompositionCreateDevice3Detour),
                           reinterpret_cast<void**>(&g_originalCreateDevice3), L"DCompositionCreateDevice3"))) {
        return false;
    }
    if (createDevice3) {
        g_createDevice3Address = reinterpret_cast<void*>(createDevice3);
    }

    if (alphaBlend &&
        FAILED(InstallHook(reinterpret_cast<void*>(alphaBlend), reinterpret_cast<void*>(&GdiAlphaBlendDetour),
                           reinterpret_cast<void**>(&g_originalGdiAlphaBlend), L"GdiAlphaBlend"))) {
        return false;
    }
    g_gdiAlphaBlendAddress = reinterpret_cast<void*>(alphaBlend);

    g_hooksInstalled = true;
    LogMessage(LogLevel::Info, L"CompositionIntercept: hooks initialized");
    return true;
}

void ShutdownCompositionIntercept() {
    if (g_hooksInstalled) {
        RemoveHook(g_createDeviceAddress, L"DCompositionCreateDevice");
        RemoveHook(g_createDevice2Address, L"DCompositionCreateDevice2");
        RemoveHook(g_createDevice3Address, L"DCompositionCreateDevice3");
        RemoveHook(g_deviceCreateTargetForHwndAddress, L"IDCompositionDevice::CreateTargetForHwnd");
        RemoveHook(g_targetSetRootAddress, L"IDCompositionTarget::SetRoot");
        RemoveHook(g_targetReleaseAddress, L"IDCompositionTarget::Release");
        RemoveHook(g_gdiAlphaBlendAddress, L"GdiAlphaBlend");
        g_hooksInstalled = false;
    }

    {
        std::lock_guard<std::mutex> guard(g_bitmapMutex);
        g_trackedBitmaps.clear();
    }
    {
        std::lock_guard<std::mutex> guard(g_contextMutex);
        g_targetContexts.clear();
        g_windowContexts.clear();
    }

    MH_STATUS status = MH_Uninitialize();
    if (status != MH_OK && status != MH_ERROR_NOT_INITIALIZED) {
        LogMessage(LogLevel::Warning, L"CompositionIntercept: MH_Uninitialize failed (status=%d)",
                   static_cast<int>(status));
    }
}

void RegisterCompositionSurface(HWND hwnd, ExplorerGlowCoordinator* coordinator) noexcept {
    if (!hwnd || !coordinator) {
        return;
    }
    HWND normalized = NormalizeWindow(hwnd);
    if (!IsExplorerWindow(normalized)) {
        return;
    }
    std::lock_guard<std::mutex> guard(g_registrationMutex);
    auto& entry = g_registrations[normalized];
    entry.coordinator = coordinator;
    ++entry.refCount;
}

void UnregisterCompositionSurface(HWND hwnd) noexcept {
    if (!hwnd) {
        return;
    }
    HWND normalized = NormalizeWindow(hwnd);
    std::lock_guard<std::mutex> guard(g_registrationMutex);
    auto it = g_registrations.find(normalized);
    if (it == g_registrations.end()) {
        return;
    }
    if (it->second.refCount > 0) {
        --it->second.refCount;
    }
    if (it->second.refCount == 0) {
        g_registrations.erase(it);
        ClearRegistration(normalized);
    }
}

void NotifyCompositionColorChange(HWND hwnd) noexcept {
    HWND normalized = NormalizeWindow(hwnd);
    std::shared_ptr<ExplorerTargetContext> context;
    {
        std::lock_guard<std::mutex> guard(g_contextMutex);
        auto it = g_windowContexts.find(normalized);
        if (it != g_windowContexts.end()) {
            context = it->second.lock();
        }
    }
    if (!context) {
        return;
    }
    std::lock_guard<std::mutex> guard(context->mutex);
    context->surfacesDirty = true;
    UpdateGradientSurfacesLocked(context);
    if (context->deviceV1) {
        context->deviceV1->Commit();
    } else if (context->deviceV2) {
        context->deviceV2->Commit();
    } else if (context->deviceV3) {
        context->deviceV3->Commit();
    }
}

}  // namespace shelltabs

