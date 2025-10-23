#include "OpenFolderCommand.h"

#include <shlguid.h>
#include <shlwapi.h>
#include <windows.h>

#include <string>
#include <utility>
#include <vector>

#include "Logging.h"
#include "Module.h"
#include "TabBandWindow.h"
#include "Utilities.h"

namespace shelltabs {

namespace {
constexpr wchar_t kCommandLabel[] = L"Open in new tab";
constexpr wchar_t kBandWindowClassName[] = L"ShellTabsBandWindow";
}

OpenFolderCommand::OpenFolderCommand() { ModuleAddRef(); }

OpenFolderCommand::~OpenFolderCommand() { ModuleRelease(); }

IFACEMETHODIMP OpenFolderCommand::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IExplorerCommand) {
        *object = static_cast<IExplorerCommand*>(this);
    } else if (riid == IID_IObjectWithSite) {
        *object = static_cast<IObjectWithSite*>(this);
    } else {
        *object = nullptr;
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) OpenFolderCommand::AddRef() { return ++m_refCount; }

IFACEMETHODIMP_(ULONG) OpenFolderCommand::Release() {
    const ULONG count = --m_refCount;
    if (count == 0) {
        delete this;
    }
    return count;
}

IFACEMETHODIMP OpenFolderCommand::GetTitle(IShellItemArray*, LPWSTR* name) {
    if (!name) {
        return E_POINTER;
    }
    *name = nullptr;
    return SHStrDupW(kCommandLabel, name);
}

IFACEMETHODIMP OpenFolderCommand::GetIcon(IShellItemArray*, LPWSTR* icon) {
    if (icon) {
        *icon = nullptr;
    }
    return E_NOTIMPL;
}

IFACEMETHODIMP OpenFolderCommand::GetToolTip(IShellItemArray*, LPWSTR* infoTip) {
    if (infoTip) {
        *infoTip = nullptr;
    }
    return E_NOTIMPL;
}

IFACEMETHODIMP OpenFolderCommand::GetCanonicalName(GUID* guidCommandName) {
    if (guidCommandName) {
        *guidCommandName = GUID_NULL;
    }
    return E_NOTIMPL;
}

IFACEMETHODIMP OpenFolderCommand::GetState(IShellItemArray* itemArray, BOOL, EXPCMDSTATE* state) {
    if (!state) {
        return E_POINTER;
    }
    *state = ECS_HIDDEN;

    UpdateFrameWindow();

    if (!HasOpenableFolder(itemArray)) {
        return S_OK;
    }

    *state = ECS_ENABLED;
    return S_OK;
}

IFACEMETHODIMP OpenFolderCommand::Invoke(IShellItemArray* itemArray, IBindCtx*) {
    UpdateFrameWindow();

    std::vector<std::wstring> paths;
    if (!CollectOpenablePaths(itemArray, &paths) || paths.empty()) {
        return S_OK;
    }

    if (!OpenPathsInNewTabs(paths)) {
        LogMessage(LogLevel::Warning, L"OpenFolderCommand::Invoke could not locate ShellTabs window");
    }

    return S_OK;
}

IFACEMETHODIMP OpenFolderCommand::GetFlags(EXPCMDFLAGS* flags) {
    if (!flags) {
        return E_POINTER;
    }
    *flags = ECF_DEFAULT;
    return S_OK;
}

IFACEMETHODIMP OpenFolderCommand::EnumSubCommands(IEnumExplorerCommand**) { return E_NOTIMPL; }

IFACEMETHODIMP OpenFolderCommand::SetSite(IUnknown* site) {
    m_site.Reset();
    m_frameWindow = nullptr;

    if (!site) {
        return S_OK;
    }

    Microsoft::WRL::ComPtr<IServiceProvider> provider;
    if (FAILED(site->QueryInterface(IID_PPV_ARGS(&provider))) || !provider) {
        return E_NOINTERFACE;
    }

    m_site = provider;
    UpdateFrameWindow();
    return S_OK;
}

IFACEMETHODIMP OpenFolderCommand::GetSite(REFIID riid, void** site) {
    if (!site) {
        return E_POINTER;
    }
    *site = nullptr;
    if (!m_site) {
        return E_FAIL;
    }
    return m_site->QueryInterface(riid, site);
}

void OpenFolderCommand::UpdateFrameWindow() {
    m_frameWindow = nullptr;
    if (!m_site) {
        return;
    }

    Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
    if (SUCCEEDED(m_site.As(&oleWindow)) && oleWindow) {
        HWND hwnd = nullptr;
        if (SUCCEEDED(oleWindow->GetWindow(&hwnd)) && hwnd) {
            m_frameWindow = hwnd;
            return;
        }
    }

    Microsoft::WRL::ComPtr<IShellBrowser> browser;
    if (SUCCEEDED(m_site->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser))) && browser) {
        HWND hwnd = nullptr;
        if (SUCCEEDED(browser->GetWindow(&hwnd)) && hwnd) {
            m_frameWindow = hwnd;
            return;
        }
    }

    browser.Reset();
    if (SUCCEEDED(m_site->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&browser))) && browser) {
        HWND hwnd = nullptr;
        if (SUCCEEDED(browser->GetWindow(&hwnd)) && hwnd) {
            m_frameWindow = hwnd;
        }
    }
}

bool OpenFolderCommand::HasOpenableFolder(IShellItemArray* items) const {
    if (!items) {
        return false;
    }

    DWORD count = 0;
    if (FAILED(items->GetCount(&count)) || count == 0) {
        return false;
    }

    for (DWORD i = 0; i < count; ++i) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(i, &item)) || !item) {
            continue;
        }
        SFGAOF attributes = 0;
        if (FAILED(item->GetAttributes(SFGAO_FILESYSTEM | SFGAO_FOLDER, &attributes))) {
            continue;
        }
        if ((attributes & (SFGAO_FILESYSTEM | SFGAO_FOLDER)) == (SFGAO_FILESYSTEM | SFGAO_FOLDER)) {
            return true;
        }
    }
    return false;
}

bool OpenFolderCommand::CollectOpenablePaths(IShellItemArray* items, std::vector<std::wstring>* paths) const {
    if (!items || !paths) {
        return false;
    }

    DWORD count = 0;
    if (FAILED(items->GetCount(&count)) || count == 0) {
        return false;
    }

    bool collected = false;
    paths->reserve(paths->size() + count);

    for (DWORD i = 0; i < count; ++i) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(i, &item)) || !item) {
            continue;
        }
        SFGAOF attributes = 0;
        if (FAILED(item->GetAttributes(SFGAO_FILESYSTEM | SFGAO_FOLDER, &attributes))) {
            continue;
        }
        if ((attributes & (SFGAO_FILESYSTEM | SFGAO_FOLDER)) != (SFGAO_FILESYSTEM | SFGAO_FOLDER)) {
            continue;
        }

        std::wstring path;
        if (!TryGetFileSystemPath(item.Get(), &path) || path.empty()) {
            continue;
        }

        paths->push_back(std::move(path));
        collected = true;
    }

    return collected;
}

bool OpenFolderCommand::OpenPathsInNewTabs(const std::vector<std::wstring>& paths) const {
    if (paths.empty()) {
        return false;
    }

    HWND bandWindow = FindBandWindow();
    if (!bandWindow) {
        return false;
    }

    for (const auto& path : paths) {
        if (path.empty()) {
            continue;
        }
        COPYDATASTRUCT copy{};
        copy.dwData = SHELLTABS_COPYDATA_OPEN_FOLDER;
        copy.cbData = static_cast<DWORD>((path.size() + 1) * sizeof(wchar_t));
        copy.lpData = const_cast<wchar_t*>(path.c_str());
        SendMessageW(bandWindow, WM_COPYDATA, reinterpret_cast<WPARAM>(m_frameWindow),
                     reinterpret_cast<LPARAM>(&copy));
    }

    return true;
}

HWND OpenFolderCommand::FindBandWindow() const {
    if (!m_frameWindow) {
        return nullptr;
    }
    return FindBandWindowRecursive(m_frameWindow);
}

HWND OpenFolderCommand::FindBandWindowRecursive(HWND parent) const {
    if (!parent) {
        return nullptr;
    }

    HWND child = FindWindowExW(parent, nullptr, nullptr, nullptr);
    while (child) {
        wchar_t className[64] = {};
        if (GetClassNameW(child, className, ARRAYSIZE(className))) {
            if (wcscmp(className, kBandWindowClassName) == 0) {
                return child;
            }
        }

        if (HWND nested = FindBandWindowRecursive(child)) {
            return nested;
        }

        child = FindWindowExW(parent, child, nullptr, nullptr);
    }

    return nullptr;
}

}  // namespace shelltabs

