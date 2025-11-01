#pragma once

#ifdef _WINSOCKAPI_
#undef _WINSOCKAPI_
#endif
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>

namespace shelltabs {

struct ShellTabsOptions;

void SetModuleHandleInstance(HMODULE module) noexcept;
HMODULE GetModuleHandleInstance() noexcept;
void ModuleAddRef() noexcept;
void ModuleRelease() noexcept;
bool ModuleCanUnload() noexcept;

// Initializes global infrastructure shared across the DLL (theme hooks, etc.).
bool ModuleInitialize();
void ModuleShutdown();
void ModuleOnOptionsChanged(const ShellTabsOptions& options);
void ModuleOnThemeChanged();
void ModuleOnSettingChanged();

}  // namespace shelltabs

