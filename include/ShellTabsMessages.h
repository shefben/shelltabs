#pragma once

#include <windows.h>

namespace shelltabs {

constexpr UINT WM_SHELLTABS_OPEN_FOLDER = WM_APP + 63;
constexpr UINT WM_SHELLTABS_REOPEN_LAST_CLOSED = WM_APP + 64;
constexpr UINT WM_SHELLTABS_CLOSE_OTHERS = WM_APP + 65;
constexpr UINT WM_SHELLTABS_CLOSE_TABS_RIGHT = WM_APP + 66;
constexpr UINT WM_SHELLTABS_CLOSE_TABS_LEFT = WM_APP + 67;

struct OpenFolderMessagePayload {
    const wchar_t* path = nullptr;
    size_t length = 0;
};

UINT GetOptionsChangedMessage();
UINT GetProgressUpdateMessage();

}  // namespace shelltabs

