#pragma once

#include <Unknwn.h>

namespace shelltabs {

HRESULT CreateTabBandClassFactory(REFIID riid, void** object);
HRESULT CreateBrowserHelperClassFactory(REFIID riid, void** object);
HRESULT CreateOpenFolderCommandClassFactory(REFIID riid, void** object);

}  // namespace shelltabs

