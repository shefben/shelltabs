#pragma once

#include <windows.h>

namespace shelltabs {

constexpr UINT WM_SHELLTABS_OPEN_FOLDER = WM_APP + 63;

struct OpenFolderMessagePayload {
    const wchar_t* path = nullptr;
    size_t length = 0;
};

}  // namespace shelltabs

