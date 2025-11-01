#pragma once

#include <Unknwn.h>

namespace shelltabs {

HRESULT CreateBrowserHelperClassFactory(REFIID riid, void** object);
HRESULT CreateOpenFolderCommandClassFactory(REFIID riid, void** object);
HRESULT CreateFtpFolderClassFactory(REFIID riid, void** object);

}  // namespace shelltabs

