#pragma once

#include <windows.h>

namespace shelltabs {

struct OptionsDialogResult {
    bool saved = false;
    bool optionsChanged = false;
    bool groupsChanged = false;
};

OptionsDialogResult ShowOptionsDialog(HWND parent, int initialTab = 0);

}  // namespace shelltabs

