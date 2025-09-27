#include "Utilities.h"

#include <shobjidl_core.h>
#include <shlwapi.h>
#include <urlmon.h>
#include <oleauto.h>
#include <commdlg.h>

#include <array>
#include <string>
#include <vector>
#include <cwchar>

#include "Logging.h"
#include "Module.h"

namespace shelltabs {

namespace {
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
    if (details && *details) {
        LogMessage(LogLevel::Error, L"%ls: unhandled exception - %ls", context && *context ? context : L"(unknown context)",
                   details);
        return;
    }
    LogMessage(LogLevel::Error, L"%ls: unhandled exception", context && *context ? context : L"(unknown context)");
}

void LogUnhandledExceptionNarrow(const wchar_t* context, const char* details) {
    const std::wstring wide = NarrowToWide(details);
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

}  // namespace shelltabs

