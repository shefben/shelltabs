#include <windows.h>

#include <CommCtrl.h>
#include <ShlGuid.h>
#include <ShlObj.h>
#include <cwchar>
#include <array>
#include <comcat.h>
#include <initializer_list>
#include <string>
#include <vector>

#include "ClassFactory.h"
#include "ComUtils.h"
#include "Guids.h"
#include "Logging.h"
#include "Module.h"

using namespace shelltabs;

namespace {

std::wstring CurrentProcessImageName() {
    wchar_t buffer[MAX_PATH] = {};
    DWORD written = GetModuleFileNameW(nullptr, buffer, ARRAYSIZE(buffer));
    if (written == 0 || written >= ARRAYSIZE(buffer)) {
        return L"(unknown process)";
    }
    return std::wstring(buffer, written);
}

#define RETURN_IF_FAILED_LOG(step, expr)                                                     \
    do {                                                                                     \
        const HRESULT returnIfFailedHr__ = (expr);                                           \
        if (FAILED(returnIfFailedHr__)) {                                                    \
            LogHrFailure(step, returnIfFailedHr__);                                          \
            return returnIfFailedHr__;                                                       \
        }                                                                                    \
        LogMessage(LogLevel::Info, L"%ls succeeded", step);                                 \
    } while (false)

constexpr wchar_t kBandFriendlyName[] = L"Shell Tabs";
constexpr wchar_t kBhoFriendlyName[] = L"Shell Tabs Browser Helper";
constexpr wchar_t kBandProgId[] = L"ShellTabs.Band";
constexpr wchar_t kBandProgIdVersion[] = L"ShellTabs.Band.1";
constexpr wchar_t kBhoProgId[] = L"ShellTabs.BrowserHelper";
constexpr wchar_t kBhoProgIdVersion[] = L"ShellTabs.BrowserHelper.1";
constexpr wchar_t kOpenFolderCommandFriendlyName[] = L"Shell Tabs Open Folder Command";
constexpr wchar_t kOpenFolderCommandVerb[] = L"ShellTabs.OpenInNewTab";
constexpr wchar_t kOpenFolderCommandKeyName[] = L"ShellTabs.OpenInNewTab";
constexpr wchar_t kOpenFolderCommandLabel[] = L"Open in new tab";
constexpr wchar_t kFtpFolderFriendlyName[] = L"Shell Tabs FTP Folder";
constexpr wchar_t kFtpNamespaceFriendlyName[] = L"Shell Tabs FTP Sites";
constexpr wchar_t kFtpNamespaceParsingName[] = L"ftp://";
constexpr DWORD kFtpShellFolderAttributes = SFGAO_FOLDER | SFGAO_HASSUBFOLDER | SFGAO_FILESYSANCESTOR |
                                             SFGAO_STORAGE | SFGAO_STORAGEANCESTOR | SFGAO_STREAM | SFGAO_CANLINK;
constexpr DWORD kFtpShellFolderFlags = 0x00000028;

struct ScopedRegKey {
    ScopedRegKey() = default;
    explicit ScopedRegKey(HKEY value) : handle(value) {}
    ~ScopedRegKey() {
        if (handle) {
            RegCloseKey(handle);
        }
    }

    ScopedRegKey(const ScopedRegKey&) = delete;
    ScopedRegKey& operator=(const ScopedRegKey&) = delete;

    ScopedRegKey(ScopedRegKey&& other) noexcept : handle(other.handle) {
        other.handle = nullptr;
    }

    ScopedRegKey& operator=(ScopedRegKey&& other) noexcept {
        if (this != &other) {
            if (handle) {
                RegCloseKey(handle);
            }
            handle = other.handle;
            other.handle = nullptr;
        }
        return *this;
    }

    HKEY get() const { return handle; }
    HKEY* put() { return &handle; }

private:
    HKEY handle = nullptr;
};

HRESULT WriteRegistryStringValue(HKEY key, const wchar_t* valueName, const wchar_t* value) {
    const DWORD length = static_cast<DWORD>((wcslen(value) + 1) * sizeof(wchar_t));
    const LONG status = RegSetValueExW(key, valueName, 0, REG_SZ, reinterpret_cast<const BYTE*>(value), length);
    return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
}

HRESULT WriteRegistryDwordValue(HKEY key, const wchar_t* valueName, DWORD value) {
    const LONG status = RegSetValueExW(key, valueName, 0, REG_DWORD, reinterpret_cast<const BYTE*>(&value), sizeof(DWORD));
    return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
}

struct RegistryTarget {
    HKEY root;
    REGSAM viewFlags;
};

bool IsCurrentProcessWow64() {
    using IsWow64ProcessFn = BOOL(WINAPI*)(HANDLE, PBOOL);
    HMODULE kernel = GetModuleHandleW(L"kernel32.dll");
    if (!kernel) {
        return false;
    }
    const auto isWow64Process = reinterpret_cast<IsWow64ProcessFn>(GetProcAddress(kernel, "IsWow64Process"));
    if (!isWow64Process) {
        return false;
    }

    BOOL isWow64 = FALSE;
    if (!isWow64Process(GetCurrentProcess(), &isWow64)) {
        return false;
    }
    return isWow64 != FALSE;
}

const std::vector<RegistryTarget>& MachineTargets() {
    static const std::vector<RegistryTarget> targets = [] {
        std::vector<RegistryTarget> result;
        result.push_back({HKEY_LOCAL_MACHINE, 0});
#if defined(KEY_WOW64_64KEY)
        if (IsCurrentProcessWow64()) {
            result.push_back({HKEY_LOCAL_MACHINE, KEY_WOW64_64KEY});
        } else if constexpr (sizeof(void*) == 8) {
#if defined(KEY_WOW64_32KEY)
            result.push_back({HKEY_LOCAL_MACHINE, KEY_WOW64_32KEY});
#endif
        }
#elif defined(KEY_WOW64_32KEY)
        if constexpr (sizeof(void*) == 8) {
            result.push_back({HKEY_LOCAL_MACHINE, KEY_WOW64_32KEY});
        }
#endif
        return result;
    }();
    return targets;
}

const std::vector<RegistryTarget>& UserTargets() {
    static const std::vector<RegistryTarget> targets = {{HKEY_CURRENT_USER, 0}};
    return targets;
}

struct RegistryAttemptResult {
    bool succeeded = false;
    bool sawAccessDenied = false;
    HRESULT firstError = S_OK;
};

template <typename Func>
RegistryAttemptResult ForEachTarget(const std::vector<RegistryTarget>& targets, Func&& func) {
    RegistryAttemptResult result;
    for (const RegistryTarget& target : targets) {
        const HRESULT hr = func(target);
        if (SUCCEEDED(hr)) {
            result.succeeded = true;
            continue;
        }

        if (result.firstError == S_OK) {
            result.firstError = hr;
        }
        if (HRESULT_CODE(hr) == ERROR_ACCESS_DENIED) {
            result.sawAccessDenied = true;
        }
    }

    if (!result.succeeded && result.firstError == S_OK) {
        result.firstError = E_FAIL;
    }
    return result;
}

HRESULT CreateRegistryKey(const RegistryTarget& target, const std::wstring& path, REGSAM access, ScopedRegKey* key) {
    HKEY rawKey = nullptr;
    const LONG status = RegCreateKeyExW(target.root, path.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE,
                                        access | target.viewFlags, nullptr, &rawKey, nullptr);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    *key = ScopedRegKey(rawKey);
    return S_OK;
}

template <typename Func>
HRESULT WriteWithMachinePreference(Func&& func, bool allowUserFallback = true) {
    const RegistryAttemptResult machine = ForEachTarget(MachineTargets(), func);
    if (machine.succeeded) {
        return S_OK;
    }

    if (allowUserFallback && machine.sawAccessDenied) {
        const RegistryAttemptResult user = ForEachTarget(UserTargets(), func);
        if (user.succeeded) {
            return S_OK;
        }
        return user.firstError;
    }

    return machine.firstError;
}

HRESULT DeleteRegistryKeyForTargets(const std::vector<RegistryTarget>& targets, const std::wstring& path,
                                    bool ignoreAccessDenied) {
    auto remover = [&](const RegistryTarget& target) -> HRESULT {
        ScopedRegKey key;
        const LONG openStatus = RegOpenKeyExW(target.root, path.c_str(), 0,
                                              (KEY_READ | KEY_WRITE) | target.viewFlags, key.put());
        if (openStatus == ERROR_FILE_NOT_FOUND) {
            return S_OK;
        }
        if (openStatus == ERROR_ACCESS_DENIED && ignoreAccessDenied) {
            return S_OK;
        }
        if (openStatus != ERROR_SUCCESS) {
            return HRESULT_FROM_WIN32(openStatus);
        }

        const LONG deleteStatus = RegDeleteTreeW(key.get(), nullptr);
        if (deleteStatus == ERROR_SUCCESS || deleteStatus == ERROR_FILE_NOT_FOUND) {
            return S_OK;
        }
        if (deleteStatus == ERROR_ACCESS_DENIED && ignoreAccessDenied) {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(deleteStatus);
    };

    const RegistryAttemptResult result = ForEachTarget(targets, remover);
    if (result.succeeded) {
        return S_OK;
    }
    if (ignoreAccessDenied && result.sawAccessDenied) {
        return S_OK;
    }
    return result.firstError;
}

HRESULT DeleteRegistryValueForTargets(const std::vector<RegistryTarget>& targets, const std::wstring& path,
                                      const std::wstring& valueName, bool ignoreAccessDenied) {
    auto remover = [&](const RegistryTarget& target) -> HRESULT {
        ScopedRegKey key;
        const LONG openStatus = RegOpenKeyExW(target.root, path.c_str(), 0, KEY_SET_VALUE | target.viewFlags, key.put());
        if (openStatus == ERROR_FILE_NOT_FOUND) {
            return S_OK;
        }
        if (openStatus == ERROR_ACCESS_DENIED && ignoreAccessDenied) {
            return S_OK;
        }
        if (openStatus != ERROR_SUCCESS) {
            return HRESULT_FROM_WIN32(openStatus);
        }

        const LONG deleteStatus = RegDeleteValueW(key.get(), valueName.empty() ? nullptr : valueName.c_str());
        if (deleteStatus == ERROR_SUCCESS || deleteStatus == ERROR_FILE_NOT_FOUND) {
            return S_OK;
        }
        if (deleteStatus == ERROR_ACCESS_DENIED && ignoreAccessDenied) {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(deleteStatus);
    };

    const RegistryAttemptResult result = ForEachTarget(targets, remover);
    if (result.succeeded) {
        return S_OK;
    }
    if (ignoreAccessDenied && result.sawAccessDenied) {
        return S_OK;
    }
    return result.firstError;
}

HRESULT DeleteRegistryKeyEverywhere(const std::wstring& path, bool ignoreAccessDenied) {
    const HRESULT machine = DeleteRegistryKeyForTargets(MachineTargets(), path, ignoreAccessDenied);
    const HRESULT user = DeleteRegistryKeyForTargets(UserTargets(), path, ignoreAccessDenied);
    if (FAILED(machine)) {
        return machine;
    }
    if (FAILED(user)) {
        return user;
    }
    return S_OK;
}

HRESULT DeleteRegistryValueEverywhere(const std::wstring& path, const std::wstring& valueName, bool ignoreAccessDenied) {
    const HRESULT machine = DeleteRegistryValueForTargets(MachineTargets(), path, valueName, ignoreAccessDenied);
    const HRESULT user = DeleteRegistryValueForTargets(UserTargets(), path, valueName, ignoreAccessDenied);
    if (FAILED(machine)) {
        return machine;
    }
    if (FAILED(user)) {
        return user;
    }
    return S_OK;
}

std::wstring ExtractFileName(const std::wstring& path) {
    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring::npos) {
        return path;
    }
    return path.substr(separator + 1);
}

HRESULT GetModulePath(std::wstring* modulePath) {
    if (!modulePath) {
        return E_POINTER;
    }

    wchar_t buffer[MAX_PATH] = {};
    if (!GetModuleFileNameW(GetModuleHandleInstance(), buffer, ARRAYSIZE(buffer))) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    *modulePath = buffer;
    return S_OK;
}

HRESULT RegisterProgIds(const std::wstring& clsidString, const wchar_t* currentProgId,
                        const wchar_t* versionIndependentProgId, const wchar_t* friendlyName) {
    if (!currentProgId || !*currentProgId || !versionIndependentProgId || !*versionIndependentProgId) {
        return S_OK;
    }

    const std::wstring currentKey = L"Software\\Classes\\" + std::wstring(currentProgId);
    const std::wstring versionIndependentKey = L"Software\\Classes\\" + std::wstring(versionIndependentProgId);

    DeleteRegistryKeyForTargets(UserTargets(), currentKey, /*ignoreAccessDenied=*/true);
    DeleteRegistryKeyForTargets(UserTargets(), versionIndependentKey, /*ignoreAccessDenied=*/true);

    HRESULT hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, versionIndependentKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            inner = WriteRegistryStringValue(key.get(), L"CLSID", clsidString.c_str());
            if (FAILED(inner)) {
                return inner;
            }
            return WriteRegistryStringValue(key.get(), L"CurVer", currentProgId);
        });
    if (FAILED(hr)) {
        return hr;
    }

    return WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, currentKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            return WriteRegistryStringValue(key.get(), L"CLSID", clsidString.c_str());
        });
}

HRESULT UnregisterProgIds(const wchar_t* currentProgId, const wchar_t* versionIndependentProgId) {
    if (currentProgId && *currentProgId) {
        HRESULT hr = DeleteRegistryKeyEverywhere(L"Software\\Classes\\" + std::wstring(currentProgId),
                                                 /*ignoreAccessDenied=*/true);
        if (FAILED(hr)) {
            return hr;
        }
    }
    if (versionIndependentProgId && *versionIndependentProgId) {
        HRESULT hr = DeleteRegistryKeyEverywhere(L"Software\\Classes\\" + std::wstring(versionIndependentProgId),
                                                 /*ignoreAccessDenied=*/true);
        if (FAILED(hr)) {
            return hr;
        }
    }
    return S_OK;
}

HRESULT RegisterAppId(const std::wstring& appId, const wchar_t* friendlyName, const std::wstring& moduleFileName) {
    const std::wstring guidKey = L"Software\\Classes\\AppID\\" + appId;
    DeleteRegistryKeyForTargets(UserTargets(), guidKey, /*ignoreAccessDenied=*/true);

    HRESULT hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, guidKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            return S_OK;
        },
        /*allowUserFallback=*/false);
    if (FAILED(hr)) {
        return hr;
    }

    if (!moduleFileName.empty()) {
        const std::wstring moduleKey = L"Software\\Classes\\AppID\\" + moduleFileName;
        DeleteRegistryKeyForTargets(UserTargets(), moduleKey, /*ignoreAccessDenied=*/true);

        hr = WriteWithMachinePreference(
            [&](const RegistryTarget& target) -> HRESULT {
                ScopedRegKey key;
                HRESULT inner = CreateRegistryKey(target, moduleKey, KEY_READ | KEY_WRITE, &key);
                if (FAILED(inner)) {
                    return inner;
                }
                return WriteRegistryStringValue(key.get(), L"AppID", appId.c_str());
            },
            /*allowUserFallback=*/false);
        if (FAILED(hr)) {
            return hr;
        }
    }

    return S_OK;
}

HRESULT UnregisterAppId(const std::wstring& appId, const std::wstring& moduleFileName) {
    HRESULT hr = DeleteRegistryKeyEverywhere(L"Software\\Classes\\AppID\\" + appId, /*ignoreAccessDenied=*/true);
    if (FAILED(hr)) {
        return hr;
    }

    if (!moduleFileName.empty()) {
        hr = DeleteRegistryKeyEverywhere(L"Software\\Classes\\AppID\\" + moduleFileName, /*ignoreAccessDenied=*/true);
        if (FAILED(hr)) {
            return hr;
        }
    }

    return S_OK;
}

HRESULT RegisterInprocServer(const std::wstring& modulePath, const std::wstring& clsidString,
                             const wchar_t* friendlyName, const wchar_t* appId,
                             const wchar_t* currentProgId, const wchar_t* versionIndependentProgId,
                             std::initializer_list<GUID> categories, bool markProgrammable = false) {
    const std::wstring baseKey = L"Software\\Classes\\CLSID\\" + clsidString;

    DeleteRegistryKeyForTargets(UserTargets(), baseKey, /*ignoreAccessDenied=*/true);

    HRESULT hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, baseKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            if (appId && *appId) {
                inner = WriteRegistryStringValue(key.get(), L"AppID", appId);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            if (currentProgId && *currentProgId) {
                inner = WriteRegistryStringValue(key.get(), L"ProgID", currentProgId);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            if (versionIndependentProgId && *versionIndependentProgId) {
                inner = WriteRegistryStringValue(key.get(), L"VersionIndependentProgID", versionIndependentProgId);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            if (markProgrammable) {
                const std::wstring programmableKey = baseKey + L"\\Programmable";
                ScopedRegKey programmable;
                inner = CreateRegistryKey(target, programmableKey, KEY_READ | KEY_WRITE, &programmable);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            return S_OK;
        });
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring inprocKey = baseKey + L"\\InprocServer32";
    hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, inprocKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryStringValue(key.get(), nullptr, modulePath.c_str());
            if (FAILED(inner)) {
                return inner;
            }
            constexpr wchar_t kThreadingModel[] = L"Apartment";
            return WriteRegistryStringValue(key.get(), L"ThreadingModel", kThreadingModel);
        });
    if (FAILED(hr)) {
        return hr;
    }

    for (const GUID& category : categories) {
        const std::wstring categoryKey = baseKey + L"\\Implemented Categories\\" + GuidToString(category);
        hr = WriteWithMachinePreference(
            [&](const RegistryTarget& target) -> HRESULT {
                ScopedRegKey key;
                return CreateRegistryKey(target, categoryKey, KEY_READ | KEY_WRITE, &key);
            });
        if (FAILED(hr)) {
            return hr;
        }
    }

    hr = RegisterProgIds(clsidString, currentProgId, versionIndependentProgId, friendlyName);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT RegisterDeskBandKey(const std::wstring& clsidString, const wchar_t* friendlyName) {
    const std::wstring keyPath = L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DeskBand\\" + clsidString;
    DeleteRegistryKeyForTargets(UserTargets(), keyPath, /*ignoreAccessDenied=*/true);

    return WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT hr = CreateRegistryKey(target, keyPath, KEY_READ | KEY_WRITE, &key);
            if (FAILED(hr)) {
                return hr;
            }
            hr = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
            if (FAILED(hr)) {
                return hr;
            }
            hr = WriteRegistryStringValue(key.get(), L"MenuText", friendlyName);
            if (FAILED(hr)) {
                return hr;
            }
            return WriteRegistryStringValue(key.get(), L"HelpText", friendlyName);
        });
}

HRESULT RegisterExplorerBar(const std::wstring& clsidString, const wchar_t* friendlyName) {
    const std::wstring keyPath = L"Software\\Microsoft\\Internet Explorer\\Explorer Bars\\" + clsidString;
    DeleteRegistryKeyForTargets(UserTargets(), keyPath, /*ignoreAccessDenied=*/true);

    return WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT hr = CreateRegistryKey(target, keyPath, KEY_READ | KEY_WRITE, &key);
            if (FAILED(hr)) {
                return hr;
            }
            hr = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
            if (FAILED(hr)) {
                return hr;
            }
            hr = WriteRegistryStringValue(key.get(), L"MenuText", friendlyName);
            if (FAILED(hr)) {
                return hr;
            }
            return WriteRegistryStringValue(key.get(), L"HelpText", friendlyName);
        },
        /*allowUserFallback=*/false);
}

HRESULT RegisterToolbarValue(const std::wstring& clsidString, const wchar_t* friendlyName) {
    constexpr const wchar_t* kToolbarKey = L"Software\\Microsoft\\Internet Explorer\\Toolbar";
    constexpr const wchar_t* kShellBrowserKey =
        L"Software\\Microsoft\\Internet Explorer\\Toolbar\\ShellBrowser";

    DeleteRegistryValueForTargets(UserTargets(), kToolbarKey, clsidString, /*ignoreAccessDenied=*/true);
    DeleteRegistryValueForTargets(UserTargets(), kShellBrowserKey, clsidString,
                                  /*ignoreAccessDenied=*/true);

    auto writeValue = [&](const std::wstring& keyPath) -> HRESULT {
        return WriteWithMachinePreference(
            [&](const RegistryTarget& target) -> HRESULT {
                ScopedRegKey key;
                HRESULT hr = CreateRegistryKey(target, keyPath, KEY_READ | KEY_WRITE, &key);
                if (FAILED(hr)) {
                    return hr;
                }
                return WriteRegistryStringValue(key.get(), clsidString.c_str(), friendlyName);
            },
            /*allowUserFallback=*/false);
    };

    HRESULT hr = writeValue(kToolbarKey);
    if (FAILED(hr)) {
        return hr;
    }
    return writeValue(kShellBrowserKey);
}

HRESULT RegisterBrowserHelper(const std::wstring& clsidString, const wchar_t* friendlyName) {
    const std::wstring keyPath =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects\\" + clsidString;

    return WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT hr = CreateRegistryKey(target, keyPath, KEY_READ | KEY_WRITE, &key);
            if (FAILED(hr)) {
                return hr;
            }
            if (friendlyName && *friendlyName) {
                return WriteRegistryStringValue(key.get(), nullptr, friendlyName);
            }
            return S_OK;
        },
        /*allowUserFallback=*/false);
}

HRESULT RegisterExplorerApproved(const std::wstring& clsidString, const wchar_t* friendlyName) {
    constexpr const wchar_t* kApprovedKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    DeleteRegistryValueForTargets(UserTargets(), kApprovedKey, clsidString, /*ignoreAccessDenied=*/true);

    return WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT hr = CreateRegistryKey(target, kApprovedKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(hr)) {
                return hr;
            }
            return WriteRegistryStringValue(key.get(), clsidString.c_str(), friendlyName);
        },
        /*allowUserFallback=*/false);
}

HRESULT RegisterOpenFolderCommand(const std::wstring& clsidString) {
    constexpr const wchar_t* kScopes[] = {
        L"Software\\Classes\\Directory\\shell\\",
        L"Software\\Classes\\Folder\\shell\\",
        L"Software\\Classes\\Drive\\shell\\",
    };

    for (const wchar_t* scope : kScopes) {
        const std::wstring keyPath = std::wstring(scope) + kOpenFolderCommandKeyName;
        DeleteRegistryKeyForTargets(UserTargets(), keyPath, /*ignoreAccessDenied=*/true);

        const HRESULT hr = WriteWithMachinePreference(
            [&](const RegistryTarget& target) -> HRESULT {
                ScopedRegKey key;
                HRESULT inner = CreateRegistryKey(target, keyPath, KEY_READ | KEY_WRITE, &key);
                if (FAILED(inner)) {
                    return inner;
                }

                if (kOpenFolderCommandLabel[0] != L'\0') {
                    inner = WriteRegistryStringValue(key.get(), nullptr, kOpenFolderCommandLabel);
                    if (FAILED(inner)) {
                        return inner;
                    }
                    inner = WriteRegistryStringValue(key.get(), L"MUIVerb", kOpenFolderCommandLabel);
                    if (FAILED(inner)) {
                        return inner;
                    }
                }

                if (kOpenFolderCommandVerb[0] != L'\0') {
                    inner = WriteRegistryStringValue(key.get(), L"Verb", kOpenFolderCommandVerb);
                    if (FAILED(inner)) {
                        return inner;
                    }
                }

                inner = WriteRegistryStringValue(key.get(), L"ExplorerCommandHandler", clsidString.c_str());
                if (FAILED(inner)) {
                    return inner;
                }

                return WriteRegistryStringValue(key.get(), L"CommandStateSync", L"");
            });
        if (FAILED(hr)) {
            return hr;
        }
    }

    return S_OK;
}

HRESULT RegisterFtpShellFolderClass(const std::wstring& modulePath, const std::wstring& clsidString,
                                    const std::wstring& appIdString) {
    HRESULT hr = RegisterInprocServer(modulePath, clsidString, kFtpFolderFriendlyName, appIdString.c_str(), nullptr, nullptr,
                                      {});
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring shellFolderKey = L"Software\\Classes\\CLSID\\" + clsidString + L"\\ShellFolder";
    hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, shellFolderKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryDwordValue(key.get(), L"Attributes", kFtpShellFolderAttributes);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryDwordValue(key.get(), L"FolderValueFlags", kFtpShellFolderFlags);
            if (FAILED(inner)) {
                return inner;
            }
            return WriteRegistryDwordValue(key.get(), L"WantsFORPARSING", 1);
        },
        /*allowUserFallback=*/false);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT RegisterNamespaceNode(const std::wstring& clsidString, const wchar_t* friendlyName,
                              const std::wstring& parsingName) {
    const std::wstring baseKey = L"Software\\Classes\\CLSID\\" + clsidString;
    DeleteRegistryKeyForTargets(UserTargets(), baseKey, /*ignoreAccessDenied=*/true);

    HRESULT hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, baseKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
                if (FAILED(inner)) {
                    return inner;
                }
            }
            if (!parsingName.empty()) {
                inner = WriteRegistryStringValue(key.get(), L"ParsingName", parsingName.c_str());
                if (FAILED(inner)) {
                    return inner;
                }
            }
            return S_OK;
        });
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring shellFolderKey = baseKey + L"\\ShellFolder";
    hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, shellFolderKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryDwordValue(key.get(), L"Attributes", kFtpShellFolderAttributes);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryDwordValue(key.get(), L"PinToNameSpaceTree", 1);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryDwordValue(key.get(), L"SortOrderIndex", 90);
            if (FAILED(inner)) {
                return inner;
            }
            inner = WriteRegistryDwordValue(key.get(), L"WantsFORPARSING", 1);
            if (FAILED(inner)) {
                return inner;
            }
            return WriteRegistryDwordValue(key.get(), L"FolderValueFlags", kFtpShellFolderFlags);
        });
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring desktopKey =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Desktop\\NameSpace\\" + clsidString;
    hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, desktopKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
            }
            return inner;
        });
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring myComputerKey =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MyComputer\\NameSpace\\" + clsidString;
    hr = WriteWithMachinePreference(
        [&](const RegistryTarget& target) -> HRESULT {
            ScopedRegKey key;
            HRESULT inner = CreateRegistryKey(target, myComputerKey, KEY_READ | KEY_WRITE, &key);
            if (FAILED(inner)) {
                return inner;
            }
            if (friendlyName && *friendlyName) {
                inner = WriteRegistryStringValue(key.get(), nullptr, friendlyName);
            }
            return inner;
        });
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT UnregisterNamespaceNode(const std::wstring& clsidString) {
    const std::wstring baseKey = L"Software\\Classes\\CLSID\\" + clsidString;
    HRESULT hr = DeleteRegistryKeyEverywhere(baseKey, /*ignoreAccessDenied=*/true);
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring desktopKey =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Desktop\\NameSpace\\" + clsidString;
    hr = DeleteRegistryKeyEverywhere(desktopKey, /*ignoreAccessDenied=*/true);
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring myComputerKey =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MyComputer\\NameSpace\\" + clsidString;
    hr = DeleteRegistryKeyEverywhere(myComputerKey, /*ignoreAccessDenied=*/true);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT UnregisterOpenFolderCommand() {
    constexpr const wchar_t* kScopes[] = {
        L"Software\\Classes\\Directory\\shell\\",
        L"Software\\Classes\\Folder\\shell\\",
        L"Software\\Classes\\Drive\\shell\\",
    };

    for (const wchar_t* scope : kScopes) {
        const std::wstring keyPath = std::wstring(scope) + kOpenFolderCommandKeyName;
        const HRESULT hr = DeleteRegistryKeyEverywhere(keyPath, /*ignoreAccessDenied=*/true);
        if (FAILED(hr)) {
            return hr;
        }
    }

    return S_OK;
}

HRESULT ClearExplorerBandCache() {
    constexpr std::array<const wchar_t*, 2> kCacheKeys = {
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Discardable\\PostSetup\\Component Categories\\{00021493-0000-0000-C000-000000000046}\\Enum",
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Discardable\\PostSetup\\Component Categories\\{00021494-0000-0000-C000-000000000046}\\Enum",
    };

    for (const auto* path : kCacheKeys) {
        const LONG status = RegDeleteTreeW(HKEY_CURRENT_USER, path);
        if (status != ERROR_SUCCESS && status != ERROR_FILE_NOT_FOUND) {
            return HRESULT_FROM_WIN32(status);
        }
    }

    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    return S_OK;
}

HRESULT UnregisterApprovedExtension(const std::wstring& clsidString) {
    constexpr const wchar_t* kApprovedKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    return DeleteRegistryValueEverywhere(kApprovedKey, clsidString, /*ignoreAccessDenied=*/true);
}

HRESULT UnregisterToolbarValue(const std::wstring& clsidString) {
    constexpr const wchar_t* kToolbarKey = L"Software\\Microsoft\\Internet Explorer\\Toolbar";
    constexpr const wchar_t* kShellBrowserKey =
        L"Software\\Microsoft\\Internet Explorer\\Toolbar\\ShellBrowser";

    HRESULT hr = DeleteRegistryValueEverywhere(kToolbarKey, clsidString, /*ignoreAccessDenied=*/true);
    if (FAILED(hr)) {
        return hr;
    }
    return DeleteRegistryValueEverywhere(kShellBrowserKey, clsidString, /*ignoreAccessDenied=*/true);
}

HRESULT UnregisterBrowserHelper(const std::wstring& clsidString) {
    const std::wstring keyPath =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects\\" + clsidString;
    return DeleteRegistryKeyEverywhere(keyPath, /*ignoreAccessDenied=*/true);
}

HRESULT UnregisterExplorerBar(const std::wstring& clsidString) {
    const std::wstring keyPath = L"Software\\Microsoft\\Internet Explorer\\Explorer Bars\\" + clsidString;
    return DeleteRegistryKeyEverywhere(keyPath, /*ignoreAccessDenied=*/true);
}

HRESULT UnregisterDeskBandKey(const std::wstring& clsidString) {
    const std::wstring keyPath = L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\DeskBand\\" + clsidString;
    return DeleteRegistryKeyEverywhere(keyPath, /*ignoreAccessDenied=*/true);
}

}  // namespace

namespace {

bool ShouldBlockProcessAttach() {
    wchar_t path[MAX_PATH] = {};
    if (!GetModuleFileNameW(nullptr, path, ARRAYSIZE(path))) {
        return false;
    }

    const wchar_t* fileName = path;
    for (const wchar_t* current = path; *current; ++current) {
        if (*current == L'\\' || *current == L'/') {
            fileName = current + 1;
        }
    }

    if (_wcsicmp(fileName, L"iexplore.exe") == 0) {
        LogMessage(LogLevel::Warning, L"Blocking ShellTabs initialization in %ls", CurrentProcessImageName().c_str());
        return true;
    }
    return false;
}

}  // namespace

BOOL APIENTRY DllMain(HMODULE module, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        InitializeLoggingEarly(module);
        LogMessage(LogLevel::Info, L"DllMain PROCESS_ATTACH for %ls", CurrentProcessImageName().c_str());
        if (ShouldBlockProcessAttach()) {
            ShutdownLogging();
            return FALSE;
        }

        shelltabs::SetModuleHandleInstance(module);
        DisableThreadLibraryCalls(module);

        INITCOMMONCONTROLSEX icc{sizeof(INITCOMMONCONTROLSEX)};
        icc.dwICC = ICC_BAR_CLASSES | ICC_TAB_CLASSES;
        if (!InitCommonControlsEx(&icc)) {
            LogLastError(L"InitCommonControlsEx", GetLastError());
        } else {
            LogMessage(LogLevel::Info, L"InitCommonControlsEx succeeded");
        }
    } else if (reason == DLL_PROCESS_DETACH) {
        LogMessage(LogLevel::Info, L"DllMain PROCESS_DETACH for %ls", CurrentProcessImageName().c_str());
        ModuleShutdown();
        ShutdownLogging();
    }
    return TRUE;
}

STDAPI DllCanUnloadNow(void) {
    return ModuleCanUnload() ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    *object = nullptr;

    if (rclsid == CLSID_ShellTabsBand) {
        return CreateTabBandClassFactory(riid, object);
    }
    if (rclsid == CLSID_ShellTabsBrowserHelper) {
        return CreateBrowserHelperClassFactory(riid, object);
    }
    if (rclsid == CLSID_ShellTabsOpenFolderCommand) {
        return CreateOpenFolderCommandClassFactory(riid, object);
    }
    if (rclsid == CLSID_ShellTabsFtpFolder) {
        return CreateFtpFolderClassFactory(riid, object);
    }

    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllRegisterServer(void) {
    LogScope scope(L"DllRegisterServer");
    LogMessage(LogLevel::Info, L"DllRegisterServer invoked");

    std::wstring modulePath;
    RETURN_IF_FAILED_LOG(L"GetModulePath", GetModulePath(&modulePath));
    LogMessage(LogLevel::Info, L"DllRegisterServer module path: %ls", modulePath.c_str());

    const std::wstring moduleFileName = ExtractFileName(modulePath);
    const std::wstring appIdString = GuidToString(APPID_ShellTabs);

    RETURN_IF_FAILED_LOG(L"RegisterAppId", RegisterAppId(appIdString, kBandFriendlyName, moduleFileName));

    const std::wstring bandClsid = GuidToString(CLSID_ShellTabsBand);
    RETURN_IF_FAILED_LOG(L"RegisterInprocServer (band)",
                         RegisterInprocServer(modulePath, bandClsid, kBandFriendlyName, appIdString.c_str(),
                                              kBandProgIdVersion, kBandProgId,
                                              {CATID_DeskBand, CATID_InfoBand, CATID_CommBand}));
    RETURN_IF_FAILED_LOG(L"RegisterDeskBandKey", RegisterDeskBandKey(bandClsid, kBandFriendlyName));
    RETURN_IF_FAILED_LOG(L"RegisterExplorerBar", RegisterExplorerBar(bandClsid, kBandFriendlyName));
    RETURN_IF_FAILED_LOG(L"RegisterExplorerApproved (band)", RegisterExplorerApproved(bandClsid, kBandFriendlyName));
    RETURN_IF_FAILED_LOG(L"RegisterToolbarValue", RegisterToolbarValue(bandClsid, kBandFriendlyName));
    RETURN_IF_FAILED_LOG(L"ClearExplorerBandCache", ClearExplorerBandCache());

    const std::wstring bhoClsid = GuidToString(CLSID_ShellTabsBrowserHelper);
    RETURN_IF_FAILED_LOG(L"RegisterInprocServer (BHO)",
                         RegisterInprocServer(modulePath, bhoClsid, kBhoFriendlyName, appIdString.c_str(),
                                              kBhoProgIdVersion, kBhoProgId, {}, /*markProgrammable=*/true));
    RETURN_IF_FAILED_LOG(L"RegisterExplorerApproved (BHO)", RegisterExplorerApproved(bhoClsid, kBhoFriendlyName));
    RETURN_IF_FAILED_LOG(L"RegisterBrowserHelper", RegisterBrowserHelper(bhoClsid, kBhoFriendlyName));

    const std::wstring commandClsid = GuidToString(CLSID_ShellTabsOpenFolderCommand);
    RETURN_IF_FAILED_LOG(L"RegisterInprocServer (open folder command)",
                         RegisterInprocServer(modulePath, commandClsid, kOpenFolderCommandFriendlyName,
                                              appIdString.c_str(), nullptr, nullptr, {}));
    RETURN_IF_FAILED_LOG(L"RegisterOpenFolderCommand", RegisterOpenFolderCommand(commandClsid));

    const std::wstring ftpClsid = GuidToString(CLSID_ShellTabsFtpFolder);
    RETURN_IF_FAILED_LOG(L"RegisterInprocServer (FTP folder)",
                         RegisterFtpShellFolderClass(modulePath, ftpClsid, appIdString));
    RETURN_IF_FAILED_LOG(L"RegisterExplorerApproved (FTP folder)",
                         RegisterExplorerApproved(ftpClsid, kFtpFolderFriendlyName));

    const std::wstring ftpNamespaceClsid = GuidToString(CLSID_ShellTabsFtpRoot);
    RETURN_IF_FAILED_LOG(L"RegisterNamespaceNode (FTP)",
                         RegisterNamespaceNode(ftpNamespaceClsid, kFtpNamespaceFriendlyName, kFtpNamespaceParsingName));

    LogMessage(LogLevel::Info, L"DllRegisterServer completed successfully");
    return S_OK;
}

STDAPI DllUnregisterServer(void) {
    LogScope scope(L"DllUnregisterServer");
    LogMessage(LogLevel::Info, L"DllUnregisterServer invoked");

    std::wstring modulePath;
    RETURN_IF_FAILED_LOG(L"GetModulePath", GetModulePath(&modulePath));
    LogMessage(LogLevel::Info, L"DllUnregisterServer module path: %ls", modulePath.c_str());

    const std::wstring moduleFileName = ExtractFileName(modulePath);
    const std::wstring appIdString = GuidToString(APPID_ShellTabs);

    const std::wstring bandClsid = GuidToString(CLSID_ShellTabsBand);
    RETURN_IF_FAILED_LOG(L"DeleteRegistryKey (band CLSID)",
                         DeleteRegistryKeyEverywhere(L"Software\\Classes\\CLSID\\" + bandClsid, /*ignoreAccessDenied=*/true));
    RETURN_IF_FAILED_LOG(L"UnregisterProgIds (band)", UnregisterProgIds(kBandProgIdVersion, kBandProgId));
    RETURN_IF_FAILED_LOG(L"UnregisterDeskBandKey", UnregisterDeskBandKey(bandClsid));
    RETURN_IF_FAILED_LOG(L"UnregisterExplorerBar", UnregisterExplorerBar(bandClsid));
    RETURN_IF_FAILED_LOG(L"UnregisterApprovedExtension (band)", UnregisterApprovedExtension(bandClsid));
    RETURN_IF_FAILED_LOG(L"UnregisterToolbarValue", UnregisterToolbarValue(bandClsid));
    RETURN_IF_FAILED_LOG(L"ClearExplorerBandCache", ClearExplorerBandCache());

    const std::wstring bhoClsid = GuidToString(CLSID_ShellTabsBrowserHelper);
    RETURN_IF_FAILED_LOG(L"DeleteRegistryKey (BHO CLSID)",
                         DeleteRegistryKeyEverywhere(L"Software\\Classes\\CLSID\\" + bhoClsid, /*ignoreAccessDenied=*/true));
    RETURN_IF_FAILED_LOG(L"UnregisterProgIds (BHO)", UnregisterProgIds(kBhoProgIdVersion, kBhoProgId));
    RETURN_IF_FAILED_LOG(L"UnregisterApprovedExtension (BHO)", UnregisterApprovedExtension(bhoClsid));
    RETURN_IF_FAILED_LOG(L"UnregisterBrowserHelper", UnregisterBrowserHelper(bhoClsid));

    const std::wstring commandClsid = GuidToString(CLSID_ShellTabsOpenFolderCommand);
    RETURN_IF_FAILED_LOG(L"DeleteRegistryKey (open folder command CLSID)",
                         DeleteRegistryKeyEverywhere(L"Software\\Classes\\CLSID\\" + commandClsid,
                                                      /*ignoreAccessDenied=*/true));
    RETURN_IF_FAILED_LOG(L"UnregisterOpenFolderCommand", UnregisterOpenFolderCommand());
    const std::wstring ftpClsid = GuidToString(CLSID_ShellTabsFtpFolder);
    RETURN_IF_FAILED_LOG(L"DeleteRegistryKey (FTP folder CLSID)",
                         DeleteRegistryKeyEverywhere(L"Software\\Classes\\CLSID\\" + ftpClsid, /*ignoreAccessDenied=*/true));
    RETURN_IF_FAILED_LOG(L"UnregisterApprovedExtension (FTP folder)", UnregisterApprovedExtension(ftpClsid));
    const std::wstring ftpNamespaceClsid = GuidToString(CLSID_ShellTabsFtpRoot);
    RETURN_IF_FAILED_LOG(L"UnregisterNamespaceNode (FTP)", UnregisterNamespaceNode(ftpNamespaceClsid));
    RETURN_IF_FAILED_LOG(L"UnregisterAppId", UnregisterAppId(appIdString, moduleFileName));

    LogMessage(LogLevel::Info, L"DllUnregisterServer completed successfully");
    return S_OK;
}

#undef RETURN_IF_FAILED_LOG

STDAPI DllInstall(BOOL install, PCWSTR cmdLine) {
    UNREFERENCED_PARAMETER(cmdLine);
    if (install) {
        HRESULT hr = DllRegisterServer();
        if (FAILED(hr)) {
            DllUnregisterServer();
        }
        return hr;
    }
    return DllUnregisterServer();
}

