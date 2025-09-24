#include <windows.h>

#include <CommCtrl.h>
#include <string>

#include "ClassFactory.h"
#include "Guids.h"
#include "Module.h"

using namespace shelltabs;

namespace {

std::wstring GuidToString(REFGUID guid) {
    wchar_t buffer[64] = {};
    const int length = StringFromGUID2(guid, buffer, ARRAYSIZE(buffer));
    if (length <= 0) {
        return {};
    }
    return std::wstring(buffer, buffer + length - 1);
}

HRESULT RegisterInprocServer(const std::wstring& modulePath, const std::wstring& clsidString,
                             const wchar_t* friendlyName, const wchar_t* category) {
    const std::wstring baseKey = L"Software\\Classes\\CLSID\\" + clsidString;

    HKEY key = nullptr;
    LONG status = RegCreateKeyExW(HKEY_CURRENT_USER, baseKey.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    status = RegSetValueExW(key, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(friendlyName),
                            static_cast<DWORD>((wcslen(friendlyName) + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    const std::wstring inprocKey = baseKey + L"\\InprocServer32";
    status = RegCreateKeyExW(HKEY_CURRENT_USER, inprocKey.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    status = RegSetValueExW(key, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(modulePath.c_str()),
                            static_cast<DWORD>((modulePath.size() + 1) * sizeof(wchar_t)));
    if (status == ERROR_SUCCESS) {
        const wchar_t threadingModel[] = L"Apartment";
        status = RegSetValueExW(key, L"ThreadingModel", 0, REG_SZ, reinterpret_cast<const BYTE*>(threadingModel),
                                static_cast<DWORD>((wcslen(threadingModel) + 1) * sizeof(wchar_t)));
    }
    RegCloseKey(key);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    if (category && *category) {
        const std::wstring categoryKey = baseKey + L"\\Implemented Categories\\" + category;
        status = RegCreateKeyExW(HKEY_CURRENT_USER, categoryKey.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr);
        if (status == ERROR_SUCCESS) {
            RegCloseKey(key);
        } else {
            return HRESULT_FROM_WIN32(status);
        }
    }

    return S_OK;
}

HRESULT RegisterDeskBandKey(const std::wstring& clsidString) {
    const std::wstring keyPath = L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Deskband\\" + clsidString;
    HKEY key = nullptr;
    LONG status = RegCreateKeyExW(HKEY_CURRENT_USER, keyPath.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    const wchar_t friendlyName[] = L"Shell Tabs";
    status = RegSetValueExW(key, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(friendlyName),
                            static_cast<DWORD>((wcslen(friendlyName) + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    return S_OK;
}

HRESULT UnregisterKeyTree(HKEY root, const std::wstring& path) {
    const LONG status = RegDeleteTreeW(root, path.c_str());
    if (status == ERROR_FILE_NOT_FOUND) {
        return S_OK;
    }
    return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
}

HRESULT RegisterApprovedExtension(const std::wstring& clsidString, const wchar_t* friendlyName) {
    const wchar_t* baseKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    HKEY key = nullptr;
    LONG status = RegCreateKeyExW(HKEY_CURRENT_USER, baseKey, 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    status = RegSetValueExW(key, clsidString.c_str(), 0, REG_SZ, reinterpret_cast<const BYTE*>(friendlyName),
                            static_cast<DWORD>((wcslen(friendlyName) + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
}

HRESULT RegisterColumnHandler(const std::wstring& clsidString) {
    const std::wstring keyPath =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ColumnHandlers\\ShellTabsTags";
    HKEY key = nullptr;
    LONG status = RegCreateKeyExW(HKEY_CURRENT_USER, keyPath.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    status = RegSetValueExW(key, nullptr, 0, REG_SZ, reinterpret_cast<const BYTE*>(clsidString.c_str()),
                            static_cast<DWORD>((clsidString.size() + 1) * sizeof(wchar_t)));
    RegCloseKey(key);
    return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
}

HRESULT UnregisterApprovedExtension(const std::wstring& clsidString) {
    const wchar_t* baseKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved";
    HKEY key = nullptr;
    LONG status = RegOpenKeyExW(HKEY_CURRENT_USER, baseKey, 0, KEY_WRITE, &key);
    if (status == ERROR_FILE_NOT_FOUND) {
        return S_OK;
    }
    if (status != ERROR_SUCCESS) {
        return HRESULT_FROM_WIN32(status);
    }

    status = RegDeleteValueW(key, clsidString.c_str());
    RegCloseKey(key);
    if (status == ERROR_FILE_NOT_FOUND || status == ERROR_SUCCESS) {
        return S_OK;
    }
    return HRESULT_FROM_WIN32(status);
}

HRESULT UnregisterColumnHandler() {
    const std::wstring keyPath =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ColumnHandlers\\ShellTabsTags";
    const LONG status = RegDeleteTreeW(HKEY_CURRENT_USER, keyPath.c_str());
    if (status == ERROR_SUCCESS || status == ERROR_FILE_NOT_FOUND) {
        return S_OK;
    }
    return HRESULT_FROM_WIN32(status);
}

}  // namespace

BOOL APIENTRY DllMain(HMODULE module, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        shelltabs::SetModuleHandleInstance(module);
        DisableThreadLibraryCalls(module);

        INITCOMMONCONTROLSEX icc{sizeof(INITCOMMONCONTROLSEX)};
        icc.dwICC = ICC_BAR_CLASSES | ICC_TAB_CLASSES;
        InitCommonControlsEx(&icc);
    }
    return TRUE;
}

extern "C" HRESULT __stdcall DllCanUnloadNow() {
    return ModuleCanUnload() ? S_OK : S_FALSE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    *object = nullptr;

    if (rclsid == CLSID_ShellTabsBand) {
        return CreateTabBandClassFactory(riid, object);
    }
    if (rclsid == CLSID_ShellTabsTagColumnProvider) {
        return CreateTagColumnProviderClassFactory(riid, object);
    }

    return CLASS_E_CLASSNOTAVAILABLE;
}

extern "C" HRESULT __stdcall DllRegisterServer() {
    wchar_t modulePath[MAX_PATH] = {};
    if (!GetModuleFileNameW(GetModuleHandleInstance(), modulePath, ARRAYSIZE(modulePath))) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const std::wstring clsidString = GuidToString(CLSID_ShellTabsBand);
    constexpr wchar_t kDeskBandCategory[] = L"{00021492-0000-0000-C000-000000000046}";
    HRESULT hr = RegisterInprocServer(modulePath, clsidString, L"Shell Tabs", kDeskBandCategory);
    if (FAILED(hr)) {
        return hr;
    }

    hr = RegisterDeskBandKey(clsidString);
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring columnClsid = GuidToString(CLSID_ShellTabsTagColumnProvider);
    hr = RegisterInprocServer(modulePath, columnClsid, L"Shell Tabs Tags Column", nullptr);
    if (FAILED(hr)) {
        return hr;
    }

    hr = RegisterApprovedExtension(columnClsid, L"Shell Tabs Tags Column");
    if (FAILED(hr)) {
        return hr;
    }

    hr = RegisterColumnHandler(columnClsid);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

extern "C" HRESULT __stdcall DllUnregisterServer() {
    const std::wstring clsidString = GuidToString(CLSID_ShellTabsBand);
    const std::wstring baseKey = L"Software\\Classes\\CLSID\\" + clsidString;
    const std::wstring deskbandKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Deskband\\" + clsidString;

    HRESULT hr = UnregisterKeyTree(HKEY_CURRENT_USER, baseKey);
    if (FAILED(hr)) {
        return hr;
    }

    hr = UnregisterKeyTree(HKEY_CURRENT_USER, deskbandKey);
    if (FAILED(hr)) {
        return hr;
    }

    const std::wstring columnClsid = GuidToString(CLSID_ShellTabsTagColumnProvider);
    const std::wstring columnBaseKey = L"Software\\Classes\\CLSID\\" + columnClsid;

    hr = UnregisterKeyTree(HKEY_CURRENT_USER, columnBaseKey);
    if (FAILED(hr)) {
        return hr;
    }

    hr = UnregisterApprovedExtension(columnClsid);
    if (FAILED(hr)) {
        return hr;
    }

    hr = UnregisterColumnHandler();
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

