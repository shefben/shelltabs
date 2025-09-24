#pragma once

#include <windows.h>

#include <memory>
#include <string>
#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <wrl/client.h>

namespace shelltabs {

struct PidlDeleter {
    void operator()(ITEMIDLIST* pidl) const noexcept;
};

using UniquePidl = std::unique_ptr<ITEMIDLIST, PidlDeleter>;

UniquePidl ClonePidl(PCIDLIST_ABSOLUTE source);
bool ArePidlsEqual(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right);
std::wstring GetDisplayName(PCIDLIST_ABSOLUTE pidl);
std::wstring GetParsingName(PCIDLIST_ABSOLUTE pidl);
UniquePidl ParseDisplayName(const std::wstring& parsingName);
UniquePidl GetCurrentFolderPidL(const Microsoft::WRL::ComPtr<IShellBrowser>& shellBrowser,
                                const Microsoft::WRL::ComPtr<IWebBrowser2>& webBrowser);

}  // namespace shelltabs

