#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <unknwn.h>
#include <windef.h>
#include <wingdi.h>
#include <wtypes.h>
#include <shobjidl.h>

// Mirror QTTabBar's IVisualProperties contract which is used by Explorer's folder view
// to manage background watermarks. Some Windows SDKs omit this definition, so we replicate it here.

#ifndef __IVisualProperties_INTERFACE_DEFINED__

enum VPWATERMARKFLAGS : UINT {
    VPWF_DEFAULT = 0,
    VPWF_ALPHABLEND = 1,
};

enum VPCOLORFLAGS : UINT;

struct __declspec(uuid("e693cf68-d967-4112-8763-99172aee5e5a")) IVisualProperties : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE SetWatermark(HBITMAP hbmp, VPWATERMARKFLAGS flags) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetColor(VPCOLORFLAGS colorFlag, COLORREF color) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetColor(VPCOLORFLAGS colorFlag, COLORREF* color) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetItemHeight(int itemHeightPixels) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetItemHeight(int* itemHeightPixels) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetFont(const LOGFONTW* logFont, BOOL redraw) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetFont(LOGFONTW** logFont) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetTheme(PCWSTR subAppName, PCWSTR subIdList) = 0;
};

#endif  // __IVisualProperties_INTERFACE_DEFINED__

