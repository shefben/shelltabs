#pragma once

#include <windows.h>

#include <string>
#include <utility>
#include <vector>

#include "GroupStore.h"

namespace shelltabs {

enum class OptionsDialogPage : int {
    kGeneral = 0,
    kCustomizations = 1,
    kContextMenus = 2,
    kGlow = 3,
    kGroups = 4,
};

struct OptionsDialogResult {
    bool saved = false;
    bool optionsChanged = false;
    bool groupsChanged = false;
    std::vector<SavedGroup> savedGroups;
    std::vector<std::pair<std::wstring, std::wstring>> renamedGroups;
    std::vector<std::wstring> removedGroupIds;
};

OptionsDialogResult ShowOptionsDialog(HWND parent,
                                     OptionsDialogPage initialPage = OptionsDialogPage::kGeneral,
                                     const wchar_t* focusSavedGroupId = nullptr,
                                     bool editFocusedGroup = false);

}  // namespace shelltabs

