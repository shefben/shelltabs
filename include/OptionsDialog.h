#pragma once

#include <windows.h>
#include <string>
#include <utility>
#include <vector>

#include "GroupStore.h"

namespace shelltabs {

// Modern options dialog with improved UX and organization
enum class OptionsDialogPage : int {
    kGeneral = 0,
    kAppearance = 1,
    kGlowEffects = 2,
    kBackgrounds = 3,
    kContextMenus = 4,
    kGroups = 5,
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
