#pragma once

#include <cstdint>
#include <windows.h>

namespace shelltabs {

constexpr UINT WM_SHELLTABS_OPEN_FOLDER = WM_APP + 63;
constexpr UINT WM_SHELLTABS_REOPEN_LAST_CLOSED = WM_APP + 64;
constexpr UINT WM_SHELLTABS_CLOSE_OTHERS = WM_APP + 65;
constexpr UINT WM_SHELLTABS_CLOSE_TABS_RIGHT = WM_APP + 66;
constexpr UINT WM_SHELLTABS_CLOSE_TABS_LEFT = WM_APP + 67;
constexpr UINT WM_SHELLTABS_LISTVIEW_BACKGROUND_READY = WM_APP + 68;
constexpr UINT WM_SHELLTABS_INITIALIZATION_COMPLETE = WM_APP + 81;
constexpr UINT WM_SHELLTABS_SHOW_HISTORY_MENU = WM_APP + 82;

enum class HistoryMenuKind : uint32_t {
    kBack = 0,
    kForward = 1,
};

struct HistoryMenuRequest {
    HistoryMenuKind kind = HistoryMenuKind::kBack;
    RECT buttonRect{};  // Screen coordinates of the toolbar button bounds
};

struct OpenFolderMessagePayload {
    const wchar_t* path = nullptr;
    size_t length = 0;
};

UINT GetOptionsChangedMessage();
UINT GetProgressUpdateMessage();
UINT GetSavedGroupsChangedMessage();

}  // namespace shelltabs

