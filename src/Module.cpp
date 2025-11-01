#include "Module.h"

#include <atomic>

#include "ThemeHooks.h"

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

bool Module::Initialize() {
    return ThemeHooks::Instance().Initialize();
}

void Module::Shutdown() {
    ThemeHooks::Instance().Shutdown();
}

}  // namespace shelltabs

