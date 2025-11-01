#include "Module.h"

#include <atomic>
#include <mutex>

#include "OptionsStore.h"
#include "ThemeHooks.h"

namespace shelltabs {

namespace {
std::atomic<long> g_moduleRefs{0};
HMODULE g_moduleHandle = nullptr;
std::mutex g_moduleInitMutex;
bool g_themeHooksInitialized = false;
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

bool ModuleInitialize() {
    std::lock_guard<std::mutex> lock(g_moduleInitMutex);
    if (g_themeHooksInitialized) {
        return true;
    }

    auto& store = OptionsStore::Instance();
    store.Load();
    g_themeHooksInitialized = InitializeThemeHooks(store.Get());
    return g_themeHooksInitialized;
}

void ModuleShutdown() {
    std::lock_guard<std::mutex> lock(g_moduleInitMutex);
    if (!g_themeHooksInitialized) {
        ShutdownThemeHooks();
        return;
    }

    ShutdownThemeHooks();
    g_themeHooksInitialized = false;
}

void ModuleOnOptionsChanged(const ShellTabsOptions& options) {
    std::lock_guard<std::mutex> lock(g_moduleInitMutex);
    if (!g_themeHooksInitialized) {
        g_themeHooksInitialized = InitializeThemeHooks(options);
        return;
    }
    UpdateThemeHooks(options);
}

void ModuleOnThemeChanged() {
    std::lock_guard<std::mutex> lock(g_moduleInitMutex);
    if (!g_themeHooksInitialized) {
        return;
    }
    NotifyThemeHooksThemeChanged();
}

void ModuleOnSettingChanged() {
    std::lock_guard<std::mutex> lock(g_moduleInitMutex);
    if (!g_themeHooksInitialized) {
        return;
    }
    NotifyThemeHooksSettingChanged();
}

}  // namespace shelltabs

