#pragma once

#include <Unknwn.h>

namespace shelltabs {

HRESULT CreateTabBandClassFactory(REFIID riid, void** object);
HRESULT CreateTagColumnProviderClassFactory(REFIID riid, void** object);
HRESULT CreateBrowserHelperClassFactory(REFIID riid, void** object);

}  // namespace shelltabs

