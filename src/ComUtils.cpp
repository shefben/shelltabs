#include "ComUtils.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <oleauto.h>

namespace shelltabs {

std::wstring GuidToString(REFGUID guid) {
    wchar_t buffer[64] = {};
    const int length = StringFromGUID2(guid, buffer, ARRAYSIZE(buffer));
    if (length <= 0) {
        return {};
    }
    return std::wstring(buffer, buffer + length - 1);
}

}  // namespace shelltabs

