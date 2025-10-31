#pragma once

#include <windows.h>

#include <string>
#include <utility>
#include <vector>

#include "GroupStore.h"

namespace shelltabs {

struct OptionsDialogResult {
    bool saved = false;
    bool optionsChanged = false;
    bool groupsChanged = false;
    std::vector<SavedGroup> savedGroups;
    std::vector<std::pair<std::wstring, std::wstring>> renamedGroups;
    std::vector<std::wstring> removedGroupIds;
};

OptionsDialogResult ShowOptionsDialog(HWND parent, int initialTab = 0,
                                     const wchar_t* focusSavedGroupId = nullptr,
                                     bool editFocusedGroup = false);

}  // namespace shelltabs

