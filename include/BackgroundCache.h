#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>

#include <objidl.h>
#include <propidl.h>

#include <memory>
#include <string>
#include <vector>

#include <gdiplus.h>

#include "OptionsStore.h"

namespace shelltabs {

struct CacheMaintenanceFailure {
    std::wstring path;
    DWORD error = 0;
    std::wstring message;
};

struct CacheMaintenanceResult {
    std::vector<std::wstring> removedPaths;
    std::vector<CacheMaintenanceFailure> failures;
};

std::wstring EnsureBackgroundCacheDirectory();

bool CopyImageToBackgroundCache(const std::wstring& sourcePath,
                                 const std::wstring& displayName,
                                 CachedImageMetadata* metadata,
                                 std::wstring* createdPath = nullptr,
                                 std::wstring* errorMessage = nullptr);

std::unique_ptr<Gdiplus::Bitmap> LoadBackgroundBitmap(const std::wstring& path);

void TouchCachedImage(const std::wstring& path) noexcept;

std::vector<std::wstring> CollectCachedImageReferences(const ShellTabsOptions& options);

void UpdateCachedImageUsage(const ShellTabsOptions& options);

CacheMaintenanceResult RemoveOrphanedCacheEntries(const ShellTabsOptions& options,
                                                  const std::vector<std::wstring>& protectedPaths = {});

}  // namespace shelltabs

