#pragma once

#include <windows.h>

namespace shelltabs {

constexpr UINT WM_SHELLTABS_OPEN_FOLDER = WM_APP + 63;

struct OpenFolderMessagePayload {
    const wchar_t* path = nullptr;
    size_t length = 0;
};

constexpr ULONG_PTR SHELLTABS_COPYDATA_OPEN_FOLDER = 'STNT';
constexpr ULONG_PTR SHELLTABS_COPYDATA_SAVE_ISLAND = 'STSI';

UINT GetOptionsChangedMessage();

}  // namespace shelltabs

