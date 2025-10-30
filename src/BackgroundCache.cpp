#include "BackgroundCache.h"

#include "Logging.h"
#include "Utilities.h"

#include <KnownFolders.h>
#include <ShlObj.h>
#include <Shlwapi.h>
#include <objbase.h>

#include <algorithm>

namespace shelltabs {
namespace {

std::wstring EnsureShellTabsDataDirectory() {
    PWSTR knownFolder = nullptr;
    if (FAILED(SHGetKnownFolderPath(FOLDERID_RoamingAppData, KF_FLAG_CREATE, nullptr, &knownFolder)) ||
        !knownFolder) {
        return {};
    }

    std::wstring directory(knownFolder);
    CoTaskMemFree(knownFolder);
    if (directory.empty()) {
        return {};
    }

    if (directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += L"ShellTabs";
    CreateDirectoryW(directory.c_str(), nullptr);
    return directory;
}

std::wstring NormalizeAndEnsureTrailingSlash(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }

    if (normalized.back() != L'\\') {
        normalized.push_back(L'\\');
    }
    return normalized;
}

}  // namespace

std::wstring EnsureBackgroundCacheDirectory() {
    std::wstring directory = EnsureShellTabsDataDirectory();
    if (directory.empty()) {
        return {};
    }

    if (directory.back() != L'\\') {
        directory.push_back(L'\\');
    }
    directory += L"Backgrounds";
    CreateDirectoryW(directory.c_str(), nullptr);
    return NormalizeFileSystemPath(directory);
}

bool CopyImageToBackgroundCache(const std::wstring& sourcePath,
                                const std::wstring& displayName,
                                CachedImageMetadata* metadata,
                                std::wstring* createdPath) {
    if (!metadata) {
        return false;
    }

    if (createdPath) {
        createdPath->clear();
    }

    std::wstring normalizedSource = NormalizeFileSystemPath(sourcePath);
    if (normalizedSource.empty()) {
        return false;
    }

    std::wstring cacheDirectory = EnsureBackgroundCacheDirectory();
    if (cacheDirectory.empty()) {
        return false;
    }

    std::wstring targetPath = normalizedSource;
    std::wstring normalizedCacheDirectory = NormalizeAndEnsureTrailingSlash(cacheDirectory);
    if (normalizedCacheDirectory.empty()) {
        return false;
    }

    if (_wcsnicmp(targetPath.c_str(), normalizedCacheDirectory.c_str(), normalizedCacheDirectory.size()) != 0) {
        GUID guid{};
        if (FAILED(CoCreateGuid(&guid))) {
            return false;
        }

        wchar_t guidBuffer[64];
        if (StringFromGUID2(guid, guidBuffer, ARRAYSIZE(guidBuffer)) <= 0) {
            return false;
        }

        std::wstring fileName(guidBuffer);
        fileName.erase(std::remove(fileName.begin(), fileName.end(), L'{'), fileName.end());
        fileName.erase(std::remove(fileName.begin(), fileName.end(), L'}'), fileName.end());
        fileName.erase(std::remove(fileName.begin(), fileName.end(), L'-'), fileName.end());

        const wchar_t* extension = PathFindExtensionW(normalizedSource.c_str());
        std::wstring ext = (extension && *extension) ? extension : L".img";

        if (normalizedCacheDirectory.back() != L'\\') {
            normalizedCacheDirectory.push_back(L'\\');
        }
        targetPath = normalizedCacheDirectory + fileName + ext;

        if (!CopyFileW(normalizedSource.c_str(), targetPath.c_str(), FALSE)) {
            LogLastError(L"CopyFileW(background cache)", GetLastError());
            return false;
        }

        if (createdPath) {
            *createdPath = targetPath;
        }
    }

    std::wstring normalizedTarget = NormalizeFileSystemPath(targetPath);
    if (normalizedTarget.empty()) {
        return false;
    }

    metadata->cachedImagePath = normalizedTarget;
    if (createdPath && !createdPath->empty()) {
        *createdPath = normalizedTarget;
    }
    if (!displayName.empty()) {
        metadata->displayName = displayName;
    } else {
        const wchar_t* nameOnly = PathFindFileNameW(normalizedSource.c_str());
        metadata->displayName = nameOnly ? nameOnly : normalizedSource;
    }

    return true;
}

std::unique_ptr<Gdiplus::Bitmap> LoadBackgroundBitmap(const std::wstring& path) {
    if (path.empty()) {
        return nullptr;
    }

    auto bitmap = std::make_unique<Gdiplus::Bitmap>(path.c_str());
    if (!bitmap || bitmap->GetLastStatus() != Gdiplus::Ok) {
        return nullptr;
    }

    return bitmap;
}

}  // namespace shelltabs

