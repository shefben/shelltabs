#include "Utilities.h"

#include <shlobj.h>
#include <shobjidl_core.h>
#include <shlwapi.h>
#include <urlmon.h>
#include <oleauto.h>
#include <commdlg.h>

#include <array>
#include <string>
#include <string_view>
#include <vector>
#include <cwchar>
#include <algorithm>
#include <PathCch.h>

#include "Logging.h"
#include "Module.h"

namespace shelltabs {

std::wstring Utf8ToWide(std::string_view utf8) {
    if (utf8.empty()) {
        return {};
    }
    const int length = MultiByteToWideChar(CP_UTF8, 0, utf8.data(), static_cast<int>(utf8.size()), nullptr, 0);
    if (length <= 0) {
        return {};
    }
    std::wstring result(static_cast<size_t>(length), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.data(), static_cast<int>(utf8.size()), result.data(), length);
    return result;
}

std::string WideToUtf8(std::wstring_view wide) {
    if (wide.empty()) {
        return {};
    }
    const int length = WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0, nullptr, nullptr);
    if (length <= 0) {
        return {};
    }
    std::string result(static_cast<size_t>(length), '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), result.data(), length, nullptr, nullptr);
    return result;
}

void LogUnhandledException(const wchar_t* context, const wchar_t* details) {
    if (details && *details) {
        LogMessage(LogLevel::Error, L"%ls: unhandled exception - %ls", context && *context ? context : L"(unknown context)",
                   details);
        return;
    }
    LogMessage(LogLevel::Error, L"%ls: unhandled exception", context && *context ? context : L"(unknown context)");
}

void LogUnhandledExceptionNarrow(const wchar_t* context, const char* details) {
    const std::wstring wide = details ? Utf8ToWide(details) : std::wstring();
    if (wide.empty()) {
        LogUnhandledException(context, nullptr);
        return;
    }
    LogUnhandledException(context, wide.c_str());
}

void PidlDeleter::operator()(AbsolutePidl* pidl) const noexcept {
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

UniquePidl CloneParent(PCIDLIST_ABSOLUTE source) {
    if (!source) {
        return nullptr;
    }
    UniquePidl clone = ClonePidl(source);
    if (!clone) {
        return nullptr;
    }
    if (!ILRemoveLastID(reinterpret_cast<PIDLIST_RELATIVE>(clone.get()))) {
        return nullptr;
    }
    return clone;
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
    FtpUrlParts ftpParts;
    if (TryParseFtpUrl(parsingName, &ftpParts)) {
        if (auto ftpPidl = CreateFtpPidlFromUrl(ftpParts)) {
            return ftpPidl;
        }
    }
    PIDLIST_ABSOLUTE pidl = nullptr;
    SFGAOF attributes = 0;
    if (SUCCEEDED(SHParseDisplayName(parsingName.c_str(), nullptr, &pidl, attributes, nullptr)) && pidl) {
        return UniquePidl(pidl);
    }
    return nullptr;
}

UniquePidl ParseExplorerUrl(const std::wstring& url) {
    if (url.empty()) {
        return nullptr;
    }

    FtpUrlParts ftpParts;
    if (TryParseFtpUrl(url, &ftpParts)) {
        if (auto ftpPidl = CreateFtpPidlFromUrl(ftpParts)) {
            return ftpPidl;
        }
    }

    PIDLIST_ABSOLUTE pidl = nullptr;
    const auto tryParse = [&pidl](const wchar_t* source, SFGAOF attributes) -> bool {
        PIDLIST_ABSOLUTE candidate = nullptr;
        if (FAILED(SHParseDisplayName(source, nullptr, &candidate, attributes, nullptr)) || !candidate) {
            return false;
        }
        pidl = candidate;
        return true;
    };

    if (tryParse(url.c_str(), SFGAO_FOLDER)) {
        return UniquePidl(pidl);
    }

    // Some shell URLs (e.g. shell:::{CLSID}) refuse the SFGAO_FOLDER hint but still resolve when
    // parsed without attribute filtering. Retry without restrictions before falling back to
    // translating file:// URLs.
    if (tryParse(url.c_str(), 0)) {
        return UniquePidl(pidl);
    }

    wchar_t path[MAX_PATH];
    DWORD pathLength = ARRAYSIZE(path);
    if (SUCCEEDED(PathCreateFromUrlW(url.c_str(), path, &pathLength, 0))) {
        if (tryParse(path, SFGAO_FOLDER) || tryParse(path, 0)) {
            return UniquePidl(pidl);
        }
    }

    return nullptr;
}

std::wstring NormalizeFileSystemPath(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    std::wstring normalized = path;
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    // Ensure the buffer is null-terminated for PathCchRemoveBackslashEx.
    normalized.push_back(L'\0');
    PWSTR end = nullptr;
    size_t remaining = 0;
    if (SUCCEEDED(PathCchRemoveBackslashEx(normalized.data(), normalized.size(), &end, &remaining)) && end) {
        *end = L'\0';
    }

    normalized.resize(wcslen(normalized.c_str()));
    return normalized;
}

bool TryGetFileSystemPath(IShellItem* item, std::wstring* path) {
    if (!item || !path) {
        return false;
    }

    PWSTR buffer = nullptr;
    const HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, &buffer);
    if (FAILED(hr) || !buffer) {
        if (buffer) {
            CoTaskMemFree(buffer);
        }
        return false;
    }

    std::wstring normalized = NormalizeFileSystemPath(buffer);
    CoTaskMemFree(buffer);

    if (normalized.empty()) {
        return false;
    }

    *path = std::move(normalized);
    return true;
}

bool IsLikelyFileSystemPath(const std::wstring& path) {
    if (path.empty()) {
        return false;
    }
    if (PathIsURLW(path.c_str())) {
        return false;
    }
    if (path.size() >= 2 && path[1] == L':') {
        return true;
    }
    if (path.rfind(L"\\\\?\\", 0) == 0) {
        return true;
    }
    if (PathIsUNCW(path.c_str())) {
        return true;
    }
    return false;
}

namespace {

UniquePidl ShellItemToPidl(IShellItem* item) {
    if (!item) {
        return nullptr;
    }

    PIDLIST_ABSOLUTE pidl = nullptr;
    if (SUCCEEDED(SHGetIDListFromObject(item, &pidl)) && pidl) {
        return UniquePidl(pidl);
    }

    PWSTR parsingName = nullptr;
    if (SUCCEEDED(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &parsingName)) && parsingName) {
        UniquePidl parsedPidl = ParseDisplayName(parsingName);
        CoTaskMemFree(parsingName);
        if (parsedPidl) {
            return parsedPidl;
        }
    } else if (parsingName) {
        CoTaskMemFree(parsingName);
    }

    return nullptr;
}

}  // namespace

std::vector<UniquePidl> GetSelectedItemsPidL(const Microsoft::WRL::ComPtr<IShellBrowser>& shellBrowser) {
    std::vector<UniquePidl> result;
    if (!shellBrowser) {
        return result;
    }

    IShellView* rawShellView = nullptr;
    if (FAILED(shellBrowser->QueryActiveShellView(&rawShellView)) || !rawShellView) {
        return result;
    }

    Microsoft::WRL::ComPtr<IShellView> shellView;
    shellView.Attach(rawShellView);

    Microsoft::WRL::ComPtr<IFolderView2> folderView2;
    shellView.As(&folderView2);
    if (folderView2) {
        Microsoft::WRL::ComPtr<IShellItemArray> selection;
        if (SUCCEEDED(folderView2->GetSelection(TRUE, &selection)) && selection) {
            DWORD count = 0;
            if (SUCCEEDED(selection->GetCount(&count)) && count > 0) {
                result.reserve(count);
                for (DWORD i = 0; i < count; ++i) {
                    Microsoft::WRL::ComPtr<IShellItem> item;
                    if (SUCCEEDED(selection->GetItemAt(i, &item)) && item) {
                        if (auto pidl = ShellItemToPidl(item.Get())) {
                            result.emplace_back(std::move(pidl));
                        }
                    }
                }
            }
        }
    }

    if (!result.empty()) {
        return result;
    }

    Microsoft::WRL::ComPtr<IFolderView> folderView;
    shellView.As(&folderView);
    if (!folderView) {
        return result;
    }

    Microsoft::WRL::ComPtr<IPersistFolder2> persist;
    if (FAILED(folderView->GetFolder(IID_PPV_ARGS(&persist))) || !persist) {
        return result;
    }

    PIDLIST_ABSOLUTE parent = nullptr;
    if (FAILED(persist->GetCurFolder(&parent)) || !parent) {
        return result;
    }
    UniquePidl parentHolder(parent);

    auto resolveIndex = [&](int index) -> UniquePidl {
        if (index < 0) {
            return nullptr;
        }
        PIDLIST_RELATIVE child = nullptr;
        if (FAILED(folderView->Item(index, &child)) || !child) {
            return nullptr;
        }
        PIDLIST_ABSOLUTE combined = ILCombine(parentHolder.get(), child);
        CoTaskMemFree(child);
        if (!combined) {
            return nullptr;
        }
        return UniquePidl(combined);
    };

    int index = -1;
    if (SUCCEEDED(folderView->GetSelectionMarkedItem(&index)) && index >= 0) {
        if (auto pidl = resolveIndex(index)) {
            result.emplace_back(std::move(pidl));
            return result;
        }
    }

    index = -1;
    if (SUCCEEDED(folderView->GetFocusedItem(&index)) && index >= 0) {
        if (auto pidl = resolveIndex(index)) {
            result.emplace_back(std::move(pidl));
            return result;
        }
    }

    Microsoft::WRL::ComPtr<IEnumIDList> selection;
    if (SUCCEEDED(folderView->Items(SVGIO_SELECTION, IID_PPV_ARGS(&selection))) && selection) {
        ULONG fetched = 0;
        PIDLIST_RELATIVE child = nullptr;
        if (SUCCEEDED(selection->Next(1, &child, &fetched)) && fetched == 1 && child) {
            PIDLIST_ABSOLUTE combined = ILCombine(parentHolder.get(), child);
            CoTaskMemFree(child);
            if (combined) {
                result.emplace_back(UniquePidl(combined));
            }
        }
    }

    return result;
}

UniquePidl GetCurrentFolderPidL(const Microsoft::WRL::ComPtr<IShellBrowser>& shellBrowser,
                                const Microsoft::WRL::ComPtr<IWebBrowser2>& webBrowser) {
    if (shellBrowser) {
        IShellView* rawShellView = nullptr;
        if (SUCCEEDED(shellBrowser->QueryActiveShellView(&rawShellView)) && rawShellView) {
            Microsoft::WRL::ComPtr<IShellView> shellView;
            shellView.Attach(rawShellView);

            IFolderView* rawFolderView = nullptr;
            if (SUCCEEDED(shellView->QueryInterface(IID_PPV_ARGS(&rawFolderView))) && rawFolderView) {
                Microsoft::WRL::ComPtr<IFolderView> folderView;
                folderView.Attach(rawFolderView);

                IPersistFolder2* rawPersist = nullptr;
                if (SUCCEEDED(folderView->GetFolder(IID_PPV_ARGS(&rawPersist))) && rawPersist) {
                    Microsoft::WRL::ComPtr<IPersistFolder2> persist;
                    persist.Attach(rawPersist);

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
                if (auto pidl = ParseExplorerUrl(url)) {
                    return pidl;
                }
            }
        }
    }

    return nullptr;
}

namespace {
constexpr WORD kPromptControlId = 3001;
constexpr WORD kEditControlId = 3002;
constexpr WORD kColorButtonId = 3003;
constexpr WORD kColorPreviewId = 3004;
constexpr WORD kColorLabelId = 3005;

void AppendWord(std::vector<BYTE>& buffer, WORD value) {
    buffer.push_back(static_cast<BYTE>(value & 0xFF));
    buffer.push_back(static_cast<BYTE>((value >> 8) & 0xFF));
}

void AppendString(std::vector<BYTE>& buffer, const std::wstring& text) {
    for (wchar_t ch : text) {
        AppendWord(buffer, static_cast<WORD>(ch));
    }
    AppendWord(buffer, 0);
}

void AlignDialogBuffer(std::vector<BYTE>& buffer) {
    while (buffer.size() % 4 != 0) {
        buffer.push_back(0);
    }
}

std::vector<BYTE> BuildInputDialogTemplate() {
    std::vector<BYTE> data;
    data.resize(sizeof(DLGTEMPLATE));
    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(data.data());
    dlg->style = DS_SETFONT | DS_MODALFRAME | WS_POPUP | WS_CAPTION | WS_SYSMENU;
    dlg->dwExtendedStyle = 0;
    dlg->cdit = 7;
    dlg->x = 10;
    dlg->y = 10;
    dlg->cx = 220;
    dlg->cy = 100;

    AppendWord(data, 0);  // menu
    AppendWord(data, 0);  // class
    AppendWord(data, 0);  // title
    AppendWord(data, 9);  // font size
    AppendString(data, L"Segoe UI");

    AlignDialogBuffer(data);
    size_t offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* prompt = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    prompt->style = WS_CHILD | WS_VISIBLE;
    prompt->dwExtendedStyle = 0;
    prompt->x = 7;
    prompt->y = 8;
    prompt->cx = 206;
    prompt->cy = 12;
    prompt->id = kPromptControlId;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* edit = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    edit->style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL;
    edit->dwExtendedStyle = 0;
    edit->x = 7;
    edit->y = 24;
    edit->cx = 206;
    edit->cy = 14;
    edit->id = kEditControlId;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0081);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* colorLabel = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    colorLabel->style = WS_CHILD | WS_VISIBLE;
    colorLabel->dwExtendedStyle = 0;
    colorLabel->x = 7;
    colorLabel->y = 44;
    colorLabel->cx = 30;
    colorLabel->cy = 12;
    colorLabel->id = kColorLabelId;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"Color:");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* colorPreview = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    colorPreview->style = WS_CHILD | WS_VISIBLE | SS_SUNKEN;
    colorPreview->dwExtendedStyle = 0;
    colorPreview->x = 42;
    colorPreview->y = 42;
    colorPreview->cx = 28;
    colorPreview->cy = 16;
    colorPreview->id = kColorPreviewId;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0082);
    AppendString(data, L"");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* colorButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    colorButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    colorButton->dwExtendedStyle = 0;
    colorButton->x = 78;
    colorButton->y = 42;
    colorButton->cx = 60;
    colorButton->cy = 16;
    colorButton->id = kColorButtonId;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Color...");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* okButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    okButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON;
    okButton->dwExtendedStyle = 0;
    okButton->x = 60;
    okButton->y = 70;
    okButton->cx = 50;
    okButton->cy = 14;
    okButton->id = IDOK;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"OK");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    offset = data.size();
    data.resize(offset + sizeof(DLGITEMTEMPLATE));
    auto* cancelButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    cancelButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    cancelButton->dwExtendedStyle = 0;
    cancelButton->x = 120;
    cancelButton->y = 70;
    cancelButton->cx = 60;
    cancelButton->cy = 14;
    cancelButton->id = IDCANCEL;
    AppendWord(data, 0xFFFF);
    AppendWord(data, 0x0080);
    AppendString(data, L"Cancel");
    AppendWord(data, 0);

    AlignDialogBuffer(data);
    return data;
}

struct InputDialogContext {
    std::wstring title;
    std::wstring prompt;
    std::wstring value;
    bool allowColor = false;
    COLORREF color = RGB(0, 120, 215);
    HBRUSH colorBrush = nullptr;
    std::array<COLORREF, 16>* customColors = nullptr;
};

INT_PTR CALLBACK InputDialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
        case WM_INITDIALOG: {
            auto* context = reinterpret_cast<InputDialogContext*>(lParam);
            SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(context));
            if (context) {
                SetWindowTextW(hwnd, context->title.c_str());
                SetDlgItemTextW(hwnd, kPromptControlId, context->prompt.c_str());
                SetDlgItemTextW(hwnd, kEditControlId, context->value.c_str());
                if (!context->allowColor) {
                    if (HWND label = GetDlgItem(hwnd, kColorLabelId)) {
                        ShowWindow(label, SW_HIDE);
                    }
                    if (HWND preview = GetDlgItem(hwnd, kColorPreviewId)) {
                        ShowWindow(preview, SW_HIDE);
                    }
                    if (HWND button = GetDlgItem(hwnd, kColorButtonId)) {
                        ShowWindow(button, SW_HIDE);
                    }
                } else {
                    if (context->colorBrush) {
                        DeleteObject(context->colorBrush);
                        context->colorBrush = nullptr;
                    }
                    context->colorBrush = CreateSolidBrush(context->color);
                    if (HWND preview = GetDlgItem(hwnd, kColorPreviewId)) {
                        InvalidateRect(preview, nullptr, TRUE);
                    }
                }
                HWND editControl = GetDlgItem(hwnd, kEditControlId);
                if (editControl) {
                    SendMessageW(editControl, EM_SETSEL, 0, -1);
                    SetFocus(editControl);
                    return FALSE;
                }
            }
            return TRUE;
        }
        case WM_COMMAND: {
            switch (LOWORD(wParam)) {
                case IDOK: {
                    auto* context = reinterpret_cast<InputDialogContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                    if (context) {
                        HWND editControl = GetDlgItem(hwnd, kEditControlId);
                        if (editControl) {
                            const int length = GetWindowTextLengthW(editControl);
                            std::wstring text(static_cast<size_t>(length) + 1, L'\0');
                            GetWindowTextW(editControl, text.data(), length + 1);
                            text.resize(wcslen(text.c_str()));
                            context->value = std::move(text);
                        }
                        if (!context->allowColor && context->colorBrush) {
                            DeleteObject(context->colorBrush);
                            context->colorBrush = nullptr;
                        }
                    }
                    EndDialog(hwnd, IDOK);
                    return TRUE;
                }
                case IDCANCEL:
                    if (auto* context = reinterpret_cast<InputDialogContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                        context && context->colorBrush) {
                        DeleteObject(context->colorBrush);
                        context->colorBrush = nullptr;
                    }
                    EndDialog(hwnd, IDCANCEL);
                    return TRUE;
                case kColorButtonId: {
                    auto* context = reinterpret_cast<InputDialogContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                    if (!context || !context->allowColor) {
                        return TRUE;
                    }
                    CHOOSECOLORW cc{};
                    cc.lStructSize = sizeof(cc);
                    cc.hwndOwner = hwnd;
                    cc.rgbResult = context->color;
                    cc.lpCustColors = context->customColors ? context->customColors->data() : nullptr;
                    cc.Flags = CC_FULLOPEN | CC_RGBINIT;
                    if (ChooseColorW(&cc)) {
                        context->color = cc.rgbResult;
                        if (context->colorBrush) {
                            DeleteObject(context->colorBrush);
                        }
                        context->colorBrush = CreateSolidBrush(context->color);
                        if (HWND preview = GetDlgItem(hwnd, kColorPreviewId)) {
                            InvalidateRect(preview, nullptr, TRUE);
                        }
                    }
                    return TRUE;
                }
            }
            break;
        }
        case WM_CTLCOLORSTATIC: {
            auto* context = reinterpret_cast<InputDialogContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (context && context->allowColor) {
                HWND control = reinterpret_cast<HWND>(lParam);
                if (control && GetDlgCtrlID(control) == kColorPreviewId) {
                    HDC dc = reinterpret_cast<HDC>(wParam);
                    SetBkColor(dc, context->color);
                    SetTextColor(dc, RGB(0, 0, 0));
                    if (!context->colorBrush) {
                        context->colorBrush = CreateSolidBrush(context->color);
                    }
                    return reinterpret_cast<INT_PTR>(context->colorBrush);
                }
            }
            break;
        }
        case WM_DESTROY: {
            auto* context = reinterpret_cast<InputDialogContext*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (context && context->colorBrush) {
                DeleteObject(context->colorBrush);
                context->colorBrush = nullptr;
            }
            break;
        }
    }
    return FALSE;
}
}  // namespace

bool PromptForTextInput(HWND parent, const std::wstring& title, const std::wstring& prompt, std::wstring* value,
                        COLORREF* color) {
    if (!value) {
        return false;
    }
    InputDialogContext context;
    context.title = title;
    context.prompt = prompt;
    context.value = *value;
    static std::array<COLORREF, 16> customColors{};
    if (color) {
        context.allowColor = true;
        context.color = *color;
        context.customColors = &customColors;
    }

    std::vector<BYTE> dialogTemplate = BuildInputDialogTemplate();
    INT_PTR result = DialogBoxIndirectParamW(GetModuleHandleInstance(),
                                             reinterpret_cast<DLGTEMPLATE*>(dialogTemplate.data()), parent,
                                             InputDialogProc, reinterpret_cast<LPARAM>(&context));
    if (result == IDOK) {
        *value = context.value;
        if (color) {
            *color = context.color;
        }
        return true;
    }
    return false;
}

bool PromptForColor(HWND parent, COLORREF initial, COLORREF* value) {
    if (!value) {
        return false;
    }
    COLORREF customColors[16] = {};
    CHOOSECOLORW cc{};
    cc.lStructSize = sizeof(cc);
    cc.hwndOwner = parent;
    cc.rgbResult = initial;
    cc.lpCustColors = customColors;
    cc.Flags = CC_FULLOPEN | CC_RGBINIT;
    if (ChooseColorW(&cc)) {
        *value = cc.rgbResult;
        return true;
    }
    return false;
}

namespace {

int CALLBACK BrowseCallbackProc(HWND hwnd, UINT message, LPARAM, LPARAM data) {
    if (message == BFFM_INITIALIZED && data) {
        const auto* initial = reinterpret_cast<const std::wstring*>(data);
        if (initial && !initial->empty()) {
            SendMessageW(hwnd, BFFM_SETSELECTIONW, TRUE, reinterpret_cast<LPARAM>(initial->c_str()));
        }
    }
    return 0;
}

}  // namespace

bool BrowseForFolder(HWND parent, std::wstring* path) {
    if (!path) {
        return false;
    }

    Microsoft::WRL::ComPtr<IFileDialog> dialog;
    if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog))) &&
        dialog) {
        DWORD options = 0;
        if (SUCCEEDED(dialog->GetOptions(&options))) {
            dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
        }
        if (!path->empty()) {
            Microsoft::WRL::ComPtr<IShellItem> folder;
            if (SUCCEEDED(SHCreateItemFromParsingName(path->c_str(), nullptr, IID_PPV_ARGS(&folder))) && folder) {
                dialog->SetFolder(folder.Get());
            }
        }
        if (SUCCEEDED(dialog->Show(parent))) {
            Microsoft::WRL::ComPtr<IShellItem> result;
            if (SUCCEEDED(dialog->GetResult(&result)) && result) {
                PWSTR buffer = nullptr;
                if (SUCCEEDED(result->GetDisplayName(SIGDN_FILESYSPATH, &buffer)) && buffer) {
                    *path = buffer;
                    CoTaskMemFree(buffer);
                    return true;
                }
                if (buffer) {
                    CoTaskMemFree(buffer);
                }
            }
            return false;
        }
    }

    std::wstring initial = *path;
    BROWSEINFOW bi{};
    bi.hwndOwner = parent;
    bi.lpszTitle = L"Select Folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_USENEWUI | BIF_NONEWFOLDERBUTTON;
    bi.lpfn = BrowseCallbackProc;
    bi.lParam = reinterpret_cast<LPARAM>(&initial);

    PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
    if (!pidl) {
        return false;
    }

    wchar_t buffer[MAX_PATH];
    bool success = SHGetPathFromIDListW(pidl, buffer) != FALSE;
    CoTaskMemFree(pidl);
    if (success) {
        *path = buffer;
    }
    return success;
}

}  // namespace shelltabs

