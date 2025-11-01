#pragma once

#ifdef _WINSOCKAPI_
#undef _WINSOCKAPI_
#endif
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>

namespace shelltabs {

class Module {
public:
    static bool Initialize();
    static void Shutdown();
};

void SetModuleHandleInstance(HMODULE module) noexcept;
HMODULE GetModuleHandleInstance() noexcept;
void ModuleAddRef() noexcept;
void ModuleRelease() noexcept;
bool ModuleCanUnload() noexcept;

}  // namespace shelltabs

