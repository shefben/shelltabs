#include "ShellTabsMessages.h"

#include <windows.h>

namespace shelltabs {

UINT GetOptionsChangedMessage() {
    static const UINT message = RegisterWindowMessageW(L"ShellTabs.OptionsChanged");
    return message;
}

}  // namespace shelltabs
