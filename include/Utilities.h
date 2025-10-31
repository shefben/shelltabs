#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>

#include <exception>
#include <memory>
#include <string>
#include <string_view>
#include <type_traits>
#include <utility>
#include <vector>
#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <wrl/client.h>

namespace shelltabs {

using AbsolutePidl = std::remove_pointer_t<PIDLIST_ABSOLUTE>;

struct PidlDeleter {
    void operator()(AbsolutePidl* pidl) const noexcept;
};

using UniquePidl = std::unique_ptr<AbsolutePidl, PidlDeleter>;

struct FtpUrlParts {
    std::wstring canonicalUrl;
    std::wstring userName;
    std::wstring password;
    std::wstring host;
    std::wstring path;
    unsigned short port = 21;
};

UniquePidl ClonePidl(PCIDLIST_ABSOLUTE source);
UniquePidl CloneParent(PCIDLIST_ABSOLUTE source);
bool ArePidlsEqual(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right);
bool ArePidlsCanonicallyEqual(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right);
std::wstring GetDisplayName(PCIDLIST_ABSOLUTE pidl);
std::wstring GetParsingName(PCIDLIST_ABSOLUTE pidl);
std::wstring GetCanonicalParsingName(PCIDLIST_ABSOLUTE pidl);
UniquePidl ParseDisplayName(const std::wstring& parsingName);
UniquePidl ParseExplorerUrl(const std::wstring& url);
bool TryParseFtpUrl(const std::wstring& url, FtpUrlParts* parts);
UniquePidl CreateFtpPidlFromUrl(const FtpUrlParts& parts);
std::wstring NormalizeFileSystemPath(const std::wstring& path);
std::wstring GetShellTabsDataDirectory();
bool TryGetFileSystemPath(IShellItem* item, std::wstring* path);
bool IsLikelyFileSystemPath(const std::wstring& path);
UniquePidl GetCurrentFolderPidL(const Microsoft::WRL::ComPtr<IShellBrowser>& shellBrowser,
                                const Microsoft::WRL::ComPtr<IWebBrowser2>& webBrowser);
std::vector<UniquePidl> GetSelectedItemsPidL(const Microsoft::WRL::ComPtr<IShellBrowser>& shellBrowser);

std::wstring Utf8ToWide(std::string_view utf8);
std::string WideToUtf8(std::wstring_view wide);
bool ReadUtf8File(const std::wstring& path, std::wstring* contents, bool* fileExists = nullptr);
bool WriteUtf8File(const std::wstring& path, std::wstring_view contents);

bool PromptForTextInput(HWND parent, const std::wstring& title, const std::wstring& prompt, std::wstring* value,
                        COLORREF* color = nullptr);
bool PromptForColor(HWND parent, COLORREF initial, COLORREF* value);
bool BrowseForFolder(HWND parent, std::wstring* path);

void LogUnhandledException(const wchar_t* context, const wchar_t* details = nullptr);
void LogUnhandledExceptionNarrow(const wchar_t* context, const char* details);

template <typename Func>
auto GuardExplorerCall(const wchar_t* context, Func&& func) noexcept
    -> std::enable_if_t<std::is_void_v<std::invoke_result_t<Func>>> {
    try {
        std::forward<Func>(func)();
    } catch (const std::exception& ex) {
        LogUnhandledExceptionNarrow(context, ex.what());
    } catch (...) {
        LogUnhandledException(context);
    }
}

template <typename Func, typename Fallback>
auto GuardExplorerCall(const wchar_t* context, Func&& func, Fallback&& fallback) noexcept
    -> std::enable_if_t<!std::is_void_v<std::invoke_result_t<Func>>, std::invoke_result_t<Func>> {
    using Result = std::invoke_result_t<Func>;
    try {
        return std::forward<Func>(func)();
    } catch (const std::exception& ex) {
        LogUnhandledExceptionNarrow(context, ex.what());
    } catch (...) {
        LogUnhandledException(context);
    }
    return std::forward<Fallback>(fallback)();
}

}  // namespace shelltabs

