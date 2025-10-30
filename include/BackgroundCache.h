#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>

#include <memory>
#include <string>

#include <gdiplus.h>

#include "OptionsStore.h"

namespace shelltabs {

std::wstring EnsureBackgroundCacheDirectory();

bool CopyImageToBackgroundCache(const std::wstring& sourcePath,
                                 const std::wstring& displayName,
                                 CachedImageMetadata* metadata,
                                 std::wstring* createdPath = nullptr);

std::unique_ptr<Gdiplus::Bitmap> LoadBackgroundBitmap(const std::wstring& path);

}  // namespace shelltabs

