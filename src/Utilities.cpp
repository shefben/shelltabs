#include "Utilities.h"

#include <shobjidl_core.h>
#include <shlwapi.h>
#include <urlmon.h>
#include <oleauto.h>

#include <string>

namespace shelltabs {

void PidlDeleter::operator()(ITEMIDLIST* pidl) const noexcept {
    if (pidl) {
        CoTaskMemFree(pidl);
    }
}

UniquePidl ClonePidl(PCIDLIST_ABSOLUTE source) {
    if (!source) {
        return nullptr;
    }
    return UniquePidl(ILCloneFull(source));
}

bool ArePidlsEqual(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right) {
    if (left == nullptr && right == nullptr) {
        return true;
    }
    if (left == nullptr || right == nullptr) {
        return false;
    }
    return ILIsEqual(left, right) == TRUE;
}

std::wstring GetDisplayName(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return {};
    }
    PWSTR name = nullptr;
    std::wstring displayName;
    if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, &name)) && name) {
        displayName.assign(name);
    }
    if (name) {
        CoTaskMemFree(name);
    }
    return displayName;
}

std::wstring GetParsingName(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return {};
    }
    PWSTR name = nullptr;
    std::wstring parsingName;
    if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_DESKTOPABSOLUTEPARSING, &name)) && name) {
        parsingName.assign(name);
    }
    if (name) {
        CoTaskMemFree(name);
    }
    return parsingName;
}

UniquePidl ParseDisplayName(const std::wstring& parsingName) {
    if (parsingName.empty()) {
        return nullptr;
    }
    PIDLIST_ABSOLUTE pidl = nullptr;
    SFGAOF attributes = 0;
    if (SUCCEEDED(SHParseDisplayName(parsingName.c_str(), nullptr, &pidl, attributes, nullptr)) && pidl) {
        return UniquePidl(pidl);
    }
    return nullptr;
}

UniquePidl GetCurrentFolderPidL(const Microsoft::WRL::ComPtr<IShellBrowser>& shellBrowser,
                                const Microsoft::WRL::ComPtr<IWebBrowser2>& webBrowser) {
    if (shellBrowser) {
        Microsoft::WRL::ComPtr<IShellView> shellView;
        if (SUCCEEDED(shellBrowser->QueryActiveShellView(&shellView)) && shellView) {
            Microsoft::WRL::ComPtr<IFolderView> folderView;
            if (SUCCEEDED(shellView.As(&folderView)) && folderView) {
                Microsoft::WRL::ComPtr<IPersistFolder2> persist;
                if (SUCCEEDED(folderView->GetFolder(IID_PPV_ARGS(&persist))) && persist) {
                    PIDLIST_ABSOLUTE pidl = nullptr;
                    if (SUCCEEDED(persist->GetCurFolder(&pidl)) && pidl) {
                        return UniquePidl(pidl);
                    }
                }
            }
        }
    }

    if (webBrowser) {
        BSTR location = nullptr;
        if (SUCCEEDED(webBrowser->get_LocationURL(&location)) && location) {
            std::wstring url(location, SysStringLen(location));
            SysFreeString(location);

            if (!url.empty()) {
                PIDLIST_ABSOLUTE pidl = nullptr;
                SFGAOF attributes = SFGAO_FOLDER;
                if (SUCCEEDED(SHParseDisplayName(url.c_str(), nullptr, &pidl, attributes, nullptr)) && pidl) {
                    return UniquePidl(pidl);
                }

                wchar_t path[MAX_PATH];
                DWORD pathLength = ARRAYSIZE(path);
                if (SUCCEEDED(PathCreateFromUrlW(url.c_str(), path, &pathLength, 0))) {
                    if (SUCCEEDED(SHParseDisplayName(path, nullptr, &pidl, attributes, nullptr)) && pidl) {
                        return UniquePidl(pidl);
                    }
                }
            }
        }
    }

    return nullptr;
}

}  // namespace shelltabs

