#include "Utilities.h"

#include <shobjidl_core.h>
#include <shlwapi.h>
#include <urlmon.h>
#include <oleauto.h>
#include <commdlg.h>

#include <string>
#include <vector>
#include <cwchar>

#include "Module.h"

namespace shelltabs {

namespace {
std::wstring BuildLogPrefix(const wchar_t* context) {
    std::wstring prefix = L"[ShellTabs] ";
    if (context && *context) {
        prefix += context;
    } else {
        prefix += L"(unknown context)";
    }
    prefix += L": ";
    return prefix;
}

std::wstring NarrowToWide(const char* value) {
    if (!value || !*value) {
        return {};
    }
    int length = MultiByteToWideChar(CP_UTF8, 0, value, -1, nullptr, 0);
    if (length <= 0) {
        return {};
    }
    std::wstring result;
    result.resize(static_cast<size_t>(length - 1));
    MultiByteToWideChar(CP_UTF8, 0, value, -1, result.data(), length);
    return result;
}
}  // namespace

void LogUnhandledException(const wchar_t* context, const wchar_t* details) {
    std::wstring message = BuildLogPrefix(context);
    message += L"unhandled exception";
    if (details && *details) {
        message += L" - ";
        message += details;
    }
    message += L"\r\n";
    OutputDebugStringW(message.c_str());
}

void LogUnhandledExceptionNarrow(const wchar_t* context, const char* details) {
    const std::wstring wide = NarrowToWide(details);
    if (wide.empty()) {
        LogUnhandledException(context, nullptr);
        return;
    }
    LogUnhandledException(context, wide.c_str());
}

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

UniquePidl ParseExplorerUrl(const std::wstring& url) {
    if (url.empty()) {
        return nullptr;
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
    dlg->cdit = 4;
    dlg->x = 10;
    dlg->y = 10;
    dlg->cx = 220;
    dlg->cy = 80;

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
    auto* okButton = reinterpret_cast<DLGITEMTEMPLATE*>(data.data() + offset);
    okButton->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON;
    okButton->dwExtendedStyle = 0;
    okButton->x = 60;
    okButton->y = 50;
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
    cancelButton->y = 50;
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
                    }
                    EndDialog(hwnd, IDOK);
                    return TRUE;
                }
                case IDCANCEL:
                    EndDialog(hwnd, IDCANCEL);
                    return TRUE;
            }
            break;
        }
    }
    return FALSE;
}
}  // namespace

bool PromptForTextInput(HWND parent, const std::wstring& title, const std::wstring& prompt, std::wstring* value) {
    if (!value) {
        return false;
    }
    InputDialogContext context;
    context.title = title;
    context.prompt = prompt;
    context.value = *value;

    std::vector<BYTE> dialogTemplate = BuildInputDialogTemplate();
    INT_PTR result = DialogBoxIndirectParamW(GetModuleHandleInstance(),
                                             reinterpret_cast<DLGTEMPLATE*>(dialogTemplate.data()), parent,
                                             InputDialogProc, reinterpret_cast<LPARAM>(&context));
    if (result == IDOK) {
        *value = context.value;
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

}  // namespace shelltabs

