#include "ComUtils.h"

#include <windows.h>

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

