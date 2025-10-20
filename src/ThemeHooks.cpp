#include "ThemeHooks.h"

#include <windows.h>
#include <CommCtrl.h>
#include <Vssym32.h>

#include <algorithm>
#include <array>
#include <cstring>
#include <mutex>
#include <vector>

#include "Logging.h"

namespace shelltabs {
namespace {

constexpr COLORREF kForcedListTextColor = RGB(255, 0, 0);
constexpr const char kThemeModuleName[] = "uxtheme.dll";

using DrawThemeTextExFn = HRESULT(STDAPICALLTYPE*)(HTHEME, HDC, int, int, LPCWSTR, int, DWORD, LPRECT, const DTTOPTS*);
using DrawThemeTextFn = HRESULT(STDAPICALLTYPE*)(HTHEME, HDC, int, int, LPCWSTR, int, DWORD, DWORD, LPRECT);

DrawThemeTextExFn g_originalDrawThemeTextEx = nullptr;
DrawThemeTextFn g_originalDrawThemeText = nullptr;

struct IatPatch {
    void** slot = nullptr;
    void* original = nullptr;
};

struct HookedModule {
    HMODULE module = nullptr;
    IatPatch drawThemeText;
    IatPatch drawThemeTextEx;
};

std::vector<HookedModule> g_hookedModules;
std::once_flag g_installOnce;
bool g_hooksInstalled = false;
std::mutex g_hookMutex;

bool IsValidModule(HMODULE module) {
    if (!module) {
        return false;
    }
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(module);
    if (!dos || dos->e_magic != IMAGE_DOS_SIGNATURE) {
        return false;
    }
    const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS*>(reinterpret_cast<const BYTE*>(module) + dos->e_lfanew);
    if (!nt || nt->Signature != IMAGE_NT_SIGNATURE) {
        return false;
    }
    return true;
}

IatPatch PatchImport(HMODULE module, const char* importedModule, const char* functionName, void* replacement) {
    if (!IsValidModule(module) || !importedModule || !functionName || !replacement) {
        return {};
    }

    auto* base = reinterpret_cast<BYTE*>(module);
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS*>(base + dos->e_lfanew);
    const IMAGE_DATA_DIRECTORY& importDir = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (importDir.VirtualAddress == 0) {
        return {};
    }

    auto* importDesc = reinterpret_cast<PIMAGE_IMPORT_DESCRIPTOR>(base + importDir.VirtualAddress);
    if (!importDesc) {
        return {};
    }

    for (; importDesc->Name != 0; ++importDesc) {
        const char* moduleName = reinterpret_cast<const char*>(base + importDesc->Name);
        if (!moduleName || _stricmp(moduleName, importedModule) != 0) {
            continue;
        }

        auto* thunk = reinterpret_cast<PIMAGE_THUNK_DATA>(base + importDesc->FirstThunk);
        auto* origThunk = importDesc->OriginalFirstThunk
                               ? reinterpret_cast<PIMAGE_THUNK_DATA>(base + importDesc->OriginalFirstThunk)
                               : reinterpret_cast<PIMAGE_THUNK_DATA>(base + importDesc->FirstThunk);

        for (; thunk && origThunk && thunk->u1.Function != 0; ++thunk, ++origThunk) {
            if (origThunk->u1.Ordinal & IMAGE_ORDINAL_FLAG) {
                continue;
            }

            auto* importByName = reinterpret_cast<PIMAGE_IMPORT_BY_NAME>(base + origThunk->u1.AddressOfData);
            if (!importByName || _stricmp(reinterpret_cast<const char*>(importByName->Name), functionName) != 0) {
                continue;
            }

            auto** slot = reinterpret_cast<void**>(&thunk->u1.Function);
            DWORD oldProtect = 0;
            if (!VirtualProtect(slot, sizeof(void*), PAGE_EXECUTE_READWRITE, &oldProtect)) {
                return {};
            }

            void* original = *slot;
            *slot = replacement;
            VirtualProtect(slot, sizeof(void*), oldProtect, &oldProtect);
            FlushInstructionCache(GetCurrentProcess(), slot, sizeof(void*));
            return {slot, original};
        }
    }

    return {};
}

void RestorePatch(const IatPatch& patch) {
    if (!patch.slot || !patch.original) {
        return;
    }

    DWORD oldProtect = 0;
    if (!VirtualProtect(patch.slot, sizeof(void*), PAGE_EXECUTE_READWRITE, &oldProtect)) {
        return;
    }

    *patch.slot = patch.original;
    VirtualProtect(patch.slot, sizeof(void*), oldProtect, &oldProtect);
    FlushInstructionCache(GetCurrentProcess(), patch.slot, sizeof(void*));
}

bool IsListViewPart(int partId) {
    return partId == LVP_LISTITEM;
}

HRESULT STDAPICALLTYPE HookedDrawThemeText(HTHEME theme, HDC hdc, int partId, int stateId, LPCWSTR text, int length,
                                           DWORD textFlags, DWORD textFlags2, LPRECT rect) {
    if (!g_originalDrawThemeText) {
        return E_UNEXPECTED;
    }

    if (!IsListViewPart(partId)) {
        return g_originalDrawThemeText(theme, hdc, partId, stateId, text, length, textFlags, textFlags2, rect);
    }

    const COLORREF previousColor = hdc ? SetTextColor(hdc, kForcedListTextColor) : CLR_INVALID;
    const int previousBk = hdc ? SetBkMode(hdc, TRANSPARENT) : 0;
    const HRESULT hr = g_originalDrawThemeText(theme, hdc, partId, stateId, text, length, textFlags, textFlags2, rect);
    if (hdc && previousColor != CLR_INVALID) {
        SetTextColor(hdc, previousColor);
        if (previousBk != 0) {
            SetBkMode(hdc, previousBk);
        }
    }
    return hr;
}

HRESULT STDAPICALLTYPE HookedDrawThemeTextEx(HTHEME theme, HDC hdc, int partId, int stateId, LPCWSTR text, int length,
                                             DWORD flags, LPRECT rect, const DTTOPTS* options) {
    if (!g_originalDrawThemeTextEx) {
        return E_UNEXPECTED;
    }

    if (!IsListViewPart(partId)) {
        return g_originalDrawThemeTextEx(theme, hdc, partId, stateId, text, length, flags, rect, options);
    }

    DTTOPTS localOptions{};
    const DTTOPTS* usedOptions = options;
    if (options && options->dwSize > 0) {
        const size_t copy = std::min<size_t>(options->dwSize, sizeof(localOptions));
        std::memcpy(&localOptions, options, copy);
        if (localOptions.dwSize < sizeof(localOptions)) {
            localOptions.dwSize = sizeof(localOptions);
        }
    } else {
        localOptions.dwSize = sizeof(localOptions);
    }

    localOptions.dwFlags |= DTT_TEXTCOLOR;
    localOptions.crText = kForcedListTextColor;
    usedOptions = &localOptions;

    return g_originalDrawThemeTextEx(theme, hdc, partId, stateId, text, length, flags, rect, usedOptions);
}

bool InstallHooksOnce() {
    HMODULE themeModule = GetModuleHandleW(L"uxtheme.dll");
    if (!themeModule) {
        themeModule = LoadLibraryW(L"uxtheme.dll");
    }
    if (!themeModule) {
        LogMessage(LogLevel::Error, L"Failed to load uxtheme.dll for hooking");
        return false;
    }

    g_originalDrawThemeTextEx = reinterpret_cast<DrawThemeTextExFn>(GetProcAddress(themeModule, "DrawThemeTextEx"));
    g_originalDrawThemeText = reinterpret_cast<DrawThemeTextFn>(GetProcAddress(themeModule, "DrawThemeText"));
    if (!g_originalDrawThemeTextEx || !g_originalDrawThemeText) {
        LogMessage(LogLevel::Error, L"Failed to resolve DrawThemeText functions");
        return false;
    }

    const std::array<const wchar_t*, 4> modulesToPatch = {
        L"comctl32.dll", L"ExplorerFrame.dll", L"shell32.dll", L"explorer.exe"};

    std::lock_guard<std::mutex> guard(g_hookMutex);
    g_hookedModules.clear();

    for (const wchar_t* moduleName : modulesToPatch) {
        HMODULE module = GetModuleHandleW(moduleName);
        if (!module) {
            continue;
        }

        HookedModule entry{};
        entry.module = module;
        entry.drawThemeText = PatchImport(module, kThemeModuleName, "DrawThemeText",
                                          reinterpret_cast<void*>(&HookedDrawThemeText));
        entry.drawThemeTextEx = PatchImport(module, kThemeModuleName, "DrawThemeTextEx",
                                            reinterpret_cast<void*>(&HookedDrawThemeTextEx));
        if (entry.drawThemeText.slot || entry.drawThemeTextEx.slot) {
            g_hookedModules.push_back(entry);
            LogMessage(LogLevel::Info, L"Installed UxTheme hooks in %ls", moduleName);
        }
    }

    if (g_hookedModules.empty()) {
        LogMessage(LogLevel::Warning, L"No modules imported DrawThemeText functions to hook");
        return false;
    }

    return true;
}

}  // namespace

bool InitializeThemeHooks() {
    std::call_once(g_installOnce, []() { g_hooksInstalled = InstallHooksOnce(); });
    return g_hooksInstalled;
}

void ShutdownThemeHooks() {
    std::lock_guard<std::mutex> guard(g_hookMutex);
    if (g_hookedModules.empty()) {
        return;
    }

    for (auto it = g_hookedModules.rbegin(); it != g_hookedModules.rend(); ++it) {
        RestorePatch(it->drawThemeText);
        RestorePatch(it->drawThemeTextEx);
    }
    g_hookedModules.clear();
}

}  // namespace shelltabs

