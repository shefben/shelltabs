#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#include <winsock2.h>
#include <windows.h>

namespace shelltabs {

void SetModuleHandleInstance(HMODULE module) noexcept;
HMODULE GetModuleHandleInstance() noexcept;
void ModuleAddRef() noexcept;
void ModuleRelease() noexcept;
bool ModuleCanUnload() noexcept;

}  // namespace shelltabs

