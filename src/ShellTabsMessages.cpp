#include "ShellTabsMessages.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>

namespace shelltabs {

UINT GetOptionsChangedMessage() {
    static const UINT message = RegisterWindowMessageW(L"ShellTabs.OptionsChanged");
    return message;
}

UINT GetProgressUpdateMessage() {
    static const UINT message = RegisterWindowMessageW(L"ShellTabs.ProgressUpdated");
    return message;
}

UINT GetSavedGroupsChangedMessage() {
    static const UINT message = RegisterWindowMessageW(L"ShellTabs.SavedGroupsChanged");
    return message;
}

}  // namespace shelltabs
