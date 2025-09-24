#include "Module.h"

#include <atomic>

namespace shelltabs {

namespace {
std::atomic<long> g_moduleRefs{0};
HMODULE g_moduleHandle = nullptr;
}

void SetModuleHandleInstance(HMODULE module) noexcept {
    g_moduleHandle = module;
}

HMODULE GetModuleHandleInstance() noexcept {
    return g_moduleHandle;
}

void ModuleAddRef() noexcept {
    ++g_moduleRefs;
}

void ModuleRelease() noexcept {
    --g_moduleRefs;
}

bool ModuleCanUnload() noexcept {
    return g_moduleRefs.load() == 0;
}

}  // namespace shelltabs

