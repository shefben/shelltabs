#pragma once

#include <windows.h>

namespace shelltabs {

void SetModuleHandleInstance(HMODULE module) noexcept;
HMODULE GetModuleHandleInstance() noexcept;
void ModuleAddRef() noexcept;
void ModuleRelease() noexcept;
bool ModuleCanUnload() noexcept;

}  // namespace shelltabs

