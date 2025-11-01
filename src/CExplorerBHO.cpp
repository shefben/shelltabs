#include "CExplorerBHO.h"
#include "ShellTabsListView.h"

#include <combaseapi.h>
#include <exdispid.h>
#include <oleauto.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlguid.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <CommCtrl.h>
#include <KnownFolders.h>
#include <windowsx.h>
#include <uxtheme.h>
#include <OleIdl.h>
#include <gdiplus.h>
#include <new>
#include <exception>

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdarg>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <mutex>
#include <limits>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>
#include <wrl/client.h>
#include <shobjidl_core.h>
#include <optional>

#include "BackgroundCache.h"
#include "BreadcrumbGradient.h"
#include "ComUtils.h"
#include "Guids.h"
#include "Logging.h"
#include "Module.h"
#include "Notifications.h"
#include "OptionsStore.h"
#include "ShellTabsTreeView.h"
#include "ShellTabsMessages.h"
#include "Utilities.h"
#include "ExplorerThemeUtils.h"
#include "TabManager.h"
#include "IconCache.h"

#ifndef TBSTATE_HOT
#define TBSTATE_HOT 0x80
#endif

#ifndef WM_DWMCOLORIZATIONCOLORCHANGED
#define WM_DWMCOLORIZATIONCOLORCHANGED 0x0320
#endif

#ifndef SFVIDM_CLIENT_OPENWINDOW
#define SFVIDM_CLIENT_OPENWINDOW 0x705B
#endif

// Add flag for transparent list view backgrounds if it's missing from the SDK
#ifndef LVS_EX_TRANSPARENTBKGND
#define LVS_EX_TRANSPARENTBKGND 0x00400000
#endif


namespace {

using shelltabs::LogLevel;
using shelltabs::LogMessage;
using shelltabs::LogMessageV;
using shelltabs::ArePidlsEqual;
using shelltabs::GetCanonicalParsingName;
using shelltabs::GetParsingName;
using shelltabs::UniquePidl;
using shelltabs::IconCache;
using shelltabs::BreadcrumbGradientPalette;
using shelltabs::EvaluateBreadcrumbGradientColor;
using shelltabs::ResolveBreadcrumbGradientPalette;

constexpr DWORD kEnsureRetryInitialDelayMs = 500;
constexpr DWORD kEnsureRetryMaxDelayMs = 4000;
constexpr DWORD kOpenInNewTabRetryDelayMs = 250;
constexpr wchar_t kUniversalBackgroundCacheKey[] = L"__shelltabs_universal_background";
constexpr UINT_PTR kAddressEditRedrawTimerId = 0x53445257;  // 'SRDW'
constexpr UINT kAddressEditRedrawCoalesceDelayMs = 30;

void ConfigureToolbarForCustomSeparators(HWND toolbar) {
    if (!toolbar || !IsWindow(toolbar)) {
        return;
    }

    const LRESULT currentStyle = SendMessageW(toolbar, TB_GETEXTENDEDSTYLE, 0, 0);
    const LRESULT desiredStyle = currentStyle | TBSTYLE_EX_HIDECLIPPEDBUTTONS;
    if (desiredStyle != currentStyle) {
        SendMessageW(toolbar, TB_SETEXTENDEDSTYLE, 0, desiredStyle);
    }
}

void ConfigureHeaderForCustomDividers(HWND header) {
    if (!header || !IsWindow(header)) {
        return;
    }

    const int itemCount = Header_GetItemCount(header);
    if (itemCount <= 0) {
        return;
    }

    const UINT dpi = GetDpiForWindow(header);
    const int kBaseThreshold = 4;
    const int threshold = std::max(kBaseThreshold, MulDiv(kBaseThreshold, static_cast<int>(dpi), 96));
    for (int index = 0; index < itemCount; ++index) {
        RECT itemRect{};
        if (!Header_GetItemRect(header, index, &itemRect)) {
            continue;
        }

        const int width = itemRect.right - itemRect.left;
        if (width > threshold) {
            continue;
        }

        HDITEMW item{};
        item.mask = HDI_FORMAT;
        if (!Header_GetItem(header, index, &item)) {
            continue;
        }

        if ((item.fmt & HDF_OWNERDRAW) != 0) {
            continue;
        }

        item.fmt |= HDF_OWNERDRAW;
        Header_SetItem(header, index, &item);
    }
}

std::optional<std::wstring> TranslateVirtualLocation(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return std::nullopt;
    }

    struct KnownVirtualFolderMapping {
        const KNOWNFOLDERID* folderId;
        const wchar_t* canonical;
    };

    static const KnownVirtualFolderMapping kKnownVirtualFolders[] = {
        {&FOLDERID_ComputerFolder, L"shell:MyComputerFolder"},
        {&FOLDERID_NetworkFolder, L"shell:NetworkPlacesFolder"},
        {&FOLDERID_ControlPanelFolder, L"shell:ControlPanelFolder"},
        {&FOLDERID_RecycleBinFolder, L"shell:RecycleBinFolder"},
        {&FOLDERID_Libraries, L"shell:Libraries"},
#ifdef FOLDERID_QuickAccess
        {&FOLDERID_QuickAccess, L"shell:QuickAccess"},
#endif
#ifdef FOLDERID_AppsFolder
        {&FOLDERID_AppsFolder, L"shell:AppsFolder"},
#endif
    };

    for (const auto& mapping : kKnownVirtualFolders) {
        PIDLIST_ABSOLUTE knownFolder = nullptr;
        if (FAILED(SHGetKnownFolderIDList(*mapping.folderId, KF_FLAG_DEFAULT | KF_FLAG_NO_ALIAS, nullptr, &knownFolder)) ||
            !knownFolder) {
            continue;
        }

        UniquePidl known(knownFolder);
        if (ArePidlsEqual(pidl, known.get())) {
            return std::wstring(mapping.canonical);
        }
    }

    auto canonical = GetCanonicalParsingName(pidl);
    if (!canonical.empty()) {
        if (canonical.rfind(L"shell:", 0) == 0) {
            return canonical;
        }
        if (canonical.rfind(L"::", 0) == 0) {
            return L"shell:" + canonical;
        }
    }

    auto parsing = GetParsingName(pidl);
    if (!parsing.empty()) {
        if (parsing.rfind(L"shell:", 0) == 0) {
            return parsing;
        }
        if (parsing.rfind(L"::", 0) == 0) {
            return L"shell:" + parsing;
        }
    }

    return std::nullopt;
}

Microsoft::WRL::ComPtr<ITypeInfo> LoadBrowserEventsTypeInfo() {
    static std::once_flag once;
    static Microsoft::WRL::ComPtr<ITypeInfo> cachedTypeInfo;
    static HRESULT cachedResult = E_FAIL;

    std::call_once(once, []() {
        Microsoft::WRL::ComPtr<ITypeLib> typeLibrary;
        HRESULT hr = LoadRegTypeLib(LIBID_SHDocVw, 1, 1, LOCALE_USER_DEFAULT, &typeLibrary);
        if (FAILED(hr)) {
            hr = LoadRegTypeLib(LIBID_SHDocVw, 1, 1, LOCALE_SYSTEM_DEFAULT, &typeLibrary);
        }
        if (FAILED(hr)) {
            hr = LoadTypeLibEx(L"shdocvw.dll", REGKIND_NONE, &typeLibrary);
        }

        if (SUCCEEDED(hr)) {
            Microsoft::WRL::ComPtr<ITypeInfo> typeInfo;
            hr = typeLibrary->GetTypeInfoOfGuid(DIID_DWebBrowserEvents2, &typeInfo);
            if (SUCCEEDED(hr)) {
                cachedTypeInfo = typeInfo;
            }
        }

        if (FAILED(hr)) {
            LogMessage(LogLevel::Warning,
                       L"CExplorerBHO failed to load DWebBrowserEvents2 type information hr=0x%08X", hr);
        }

        cachedResult = hr;
    });

    if (SUCCEEDED(cachedResult) && cachedTypeInfo) {
        Microsoft::WRL::ComPtr<ITypeInfo> result = cachedTypeInfo;
        return result;
    }

    return nullptr;
}

#ifndef ERROR_AUTOMATION_DISABLED
#define ERROR_AUTOMATION_DISABLED 430L
#endif

#ifndef ERROR_ACCESS_DISABLED_BY_POLICY
#define ERROR_ACCESS_DISABLED_BY_POLICY 1260L
#endif

#ifndef ERROR_ACCESS_DISABLED_BY_POLICY_ADMIN
#define ERROR_ACCESS_DISABLED_BY_POLICY_ADMIN 1262L
#endif

#ifndef ERROR_ACCESS_DISABLED_BY_POLICY_DEFAULT
#define ERROR_ACCESS_DISABLED_BY_POLICY_DEFAULT 1261L
#endif

#ifndef ERROR_ACCESS_DISABLED_BY_POLICY_OTHER
#define ERROR_ACCESS_DISABLED_BY_POLICY_OTHER 1263L
#endif

#ifndef RPC_E_CALL_REJECTED
#define RPC_E_CALL_REJECTED _HRESULT_TYPEDEF_(0x80010001L)
#endif

#ifndef RPC_E_SERVERCALL_RETRYLATER
#define RPC_E_SERVERCALL_RETRYLATER _HRESULT_TYPEDEF_(0x8001010AL)
#endif

#ifndef RPC_E_SERVERCALL_REJECTED
#define RPC_E_SERVERCALL_REJECTED _HRESULT_TYPEDEF_(0x8001010BL)
#endif

#ifndef SID_STopLevelBrowserFrame
EXTERN_C const GUID SID_STopLevelBrowserFrame;
#endif

#ifndef SID_NamespaceTreeControl
EXTERN_C const GUID SID_NamespaceTreeControl;
#endif

bool IsShowBrowserBarThrottled(HRESULT hr) {
    if (hr == S_FALSE) {
        return true;
    }

    switch (hr) {
        case RPC_E_CALL_REJECTED:
        case RPC_E_SERVERCALL_RETRYLATER:
        case RPC_E_SERVERCALL_REJECTED:
            return true;
        default:
            break;
    }

    switch (HRESULT_CODE(hr)) {
        case ERROR_RETRY:
        case ERROR_BUSY:
        case ERROR_TIMEOUT:
            return HRESULT_FACILITY(hr) == FACILITY_WIN32;
        default:
            break;
    }

    return false;
}

bool IsAutomationDisabledResult(HRESULT hr) {
    if (hr == HRESULT_FROM_WIN32(ERROR_AUTOMATION_DISABLED)) {
        return true;
    }

    if (HRESULT_FACILITY(hr) != FACILITY_WIN32) {
        return false;
    }

    switch (HRESULT_CODE(hr)) {
        case ERROR_ACCESS_DISABLED_BY_POLICY:
        case ERROR_ACCESS_DISABLED_BY_POLICY_ADMIN:
        case ERROR_ACCESS_DISABLED_BY_POLICY_DEFAULT:
        case ERROR_ACCESS_DISABLED_BY_POLICY_OTHER:
            return true;
        default:
            break;
    }

    return false;
}

#ifndef ListView_GetItemW
BOOL ListView_GetItemW(HWND hwnd, LVITEMW* item) {
    return static_cast<BOOL>(SendMessageW(hwnd, LVM_GETITEMW, 0, reinterpret_cast<LPARAM>(item)));
}
#endif

#ifndef TreeView_GetItemW
BOOL TreeView_GetItemW(HWND hwnd, TVITEMEXW* item) {
    return static_cast<BOOL>(SendMessageW(hwnd, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(item)));
}
#endif

bool MatchesClass(HWND hwnd, const wchar_t* className) {
    if (!hwnd || !className) {
        return false;
    }
    wchar_t buffer[256];
    const int length = GetClassNameW(hwnd, buffer, ARRAYSIZE(buffer));
    if (length <= 0) {
        return false;
    }
    return _wcsicmp(buffer, className) == 0;
}

HWND FindDescendantWindow(HWND parent, const wchar_t* className) {
    if (!parent || !className) {
        return nullptr;
    }
    for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT)) {
        if (MatchesClass(child, className)) {
            return child;
        }
        if (HWND found = FindDescendantWindow(child, className)) {
            return found;
        }
    }
    return nullptr;
}

std::wstring NormalizeMenuText(const std::wstring& value) {
    if (value.empty()) {
        return {};
    }

    std::wstring normalized;
    normalized.reserve(value.size());
    for (wchar_t ch : value) {
        if (ch == L'&' || ch == L'.' || ch == 0x2026) {  // 0x2026 = ellipsis
            continue;
        }
        normalized.push_back(static_cast<wchar_t>(towlower(ch)));
    }

    const size_t first = normalized.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos) {
        return {};
    }
    const size_t last = normalized.find_last_not_of(L" \t\r\n");
    return normalized.substr(first, last - first + 1);
}

using GetDpiForWindowFunction = UINT(WINAPI*)(HWND);
using SetThreadDpiAwarenessContextFunction = DPI_AWARENESS_CONTEXT(WINAPI*)(DPI_AWARENESS_CONTEXT);

GetDpiForWindowFunction ResolveGetDpiForWindow() {
    static const GetDpiForWindowFunction function =
        reinterpret_cast<GetDpiForWindowFunction>(GetProcAddress(GetModuleHandleW(L"user32"), "GetDpiForWindow"));
    return function;
}

SetThreadDpiAwarenessContextFunction ResolveSetThreadDpiAwarenessContext() {
    static const SetThreadDpiAwarenessContextFunction function =
        reinterpret_cast<SetThreadDpiAwarenessContextFunction>(
            GetProcAddress(GetModuleHandleW(L"user32"), "SetThreadDpiAwarenessContext"));
    return function;
}

UINT GetWindowDpi(HWND hwnd) {
    if (auto* function = ResolveGetDpiForWindow()) {
        const UINT dpi = function(hwnd);
        if (dpi != 0) {
            return dpi;
        }
    }

    HDC localDc = hwnd ? GetDC(hwnd) : GetDC(nullptr);
    if (!localDc) {
        return 96u;
    }

    const int dpi = GetDeviceCaps(localDc, LOGPIXELSX);
    if (hwnd) {
        ReleaseDC(hwnd, localDc);
    } else {
        ReleaseDC(nullptr, localDc);
    }

    return dpi > 0 ? static_cast<UINT>(dpi) : 96u;
}

class ScopedThreadDpiAwarenessContext {
public:
    ScopedThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT desiredContext, bool enabled)
        : m_function(ResolveSetThreadDpiAwarenessContext()) {
        if (!enabled || !m_function || !desiredContext) {
            return;
        }

        m_previousContext = m_function(desiredContext);
        m_applied = (m_previousContext != nullptr);
    }

    ScopedThreadDpiAwarenessContext(const ScopedThreadDpiAwarenessContext&) = delete;
    ScopedThreadDpiAwarenessContext& operator=(const ScopedThreadDpiAwarenessContext&) = delete;

    ~ScopedThreadDpiAwarenessContext() {
        if (m_applied && m_function) {
            m_function(m_previousContext);
        }
    }

private:
    SetThreadDpiAwarenessContextFunction m_function{};
    DPI_AWARENESS_CONTEXT m_previousContext{};
    bool m_applied{false};
};

bool TryGetMenuItemText(HMENU menu, UINT position, std::wstring& text) {
    text.clear();
    if (!menu) {
        return false;
    }

    MENUITEMINFOW info{};
    info.cbSize = sizeof(info);
    info.fMask = MIIM_STRING;
    info.dwTypeData = nullptr;
    info.cch = 0;
    if (!GetMenuItemInfoW(menu, position, TRUE, &info)) {
        return false;
    }

    if (info.cch == 0) {
        return true;
    }

    std::wstring buffer;
    buffer.resize(info.cch);
    info.dwTypeData = buffer.data();
    info.cch = static_cast<UINT>(buffer.size());
    if (!GetMenuItemInfoW(menu, position, TRUE, &info)) {
        return false;
    }

    buffer.resize(info.cch);
    text = std::move(buffer);
    return true;
}

bool FindMenuItemById(HMENU menu, UINT commandId, UINT* position) {
    if (!menu) {
        return false;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0) {
        return false;
    }

    for (int i = 0; i < count; ++i) {
        const UINT id = GetMenuItemID(menu, i);
        if (id == commandId) {
            if (position) {
                *position = static_cast<UINT>(i);
            }
            return true;
        }
    }

    return false;
}

bool FindOpenInNewWindowMenuItem(HMENU menu, UINT* position, UINT* commandId) {
    if (!menu) {
        return false;
    }

    const UINT candidates[] = {SFVIDM_CLIENT_OPENWINDOW, 0x705A, 0x7059, 0x7020};
    for (UINT candidate : candidates) {
        UINT pos = 0;
        if (FindMenuItemById(menu, candidate, &pos)) {
            if (position) {
                *position = pos;
            }
            if (commandId) {
                *commandId = candidate;
            }
            return true;
        }
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0) {
        return false;
    }

    static const wchar_t* kTargets[] = {L"open in new window", L"open new window"};

    for (int i = 0; i < count; ++i) {
        const UINT id = GetMenuItemID(menu, i);
        if (id == UINT_MAX) {
            continue;
        }

        std::wstring text;
        if (!TryGetMenuItemText(menu, i, text)) {
            continue;
        }

        const std::wstring normalized = NormalizeMenuText(text);
        if (normalized.empty()) {
            continue;
        }

        for (const wchar_t* target : kTargets) {
            if (normalized == target) {
                if (position) {
                    *position = static_cast<UINT>(i);
                }
                if (commandId) {
                    *commandId = id;
                }
                return true;
            }
        }
    }

    return false;
}

bool IsSeparatorItem(HMENU menu, UINT position) {
    if (!menu) {
        return false;
    }

    MENUITEMINFOW info{};
    info.cbSize = sizeof(info);
    info.fMask = MIIM_FTYPE;
    if (!GetMenuItemInfoW(menu, position, TRUE, &info)) {
        return false;
    }

    return (info.fType & MFT_SEPARATOR) != 0;
}

std::wstring ExtractLowercaseExtension(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    const size_t slash = path.find_last_of(L"\\/");
    const size_t dot = path.find_last_of(L'.');
    if (dot == std::wstring::npos || (slash != std::wstring::npos && dot < slash + 1)) {
        return {};
    }

    std::wstring extension = path.substr(dot);
    std::transform(extension.begin(), extension.end(), extension.begin(), [](wchar_t ch) {
        return static_cast<wchar_t>(std::towlower(static_cast<unsigned int>(ch)));
    });
    return extension;
}

std::wstring ExtractParentDirectory(const std::wstring& path) {
    if (path.empty()) {
        return {};
    }

    size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring::npos) {
        return {};
    }

    if (slash == 0 && path.size() > 1 && path[1] == L':') {
        return path.substr(0, 2);
    }

    return path.substr(0, slash);
}

bool ContainsToken(const std::wstring& command, std::wstring_view token) {
    return command.find(std::wstring(token)) != std::wstring::npos;
}

std::wstring ReplaceToken(const std::wstring& input, std::wstring_view token,
                          const std::wstring& replacement) {
    if (token.empty()) {
        return input;
    }

    std::wstring result = input;
    const std::wstring pattern(token);
    size_t position = 0;
    while ((position = result.find(pattern, position)) != std::wstring::npos) {
        result.replace(position, pattern.size(), replacement);
        position += replacement.size();
    }
    return result;
}

std::wstring QuoteArgument(const std::wstring& argument) {
    if (argument.empty()) {
        return L"\"\"";
    }

    bool needsQuotes = false;
    for (wchar_t ch : argument) {
        if (std::iswspace(ch) || ch == L'"') {
            needsQuotes = true;
            break;
        }
    }

    if (!needsQuotes) {
        return argument;
    }

    std::wstring result;
    result.reserve(argument.size() + 2);
    result.push_back(L'"');
    size_t consecutiveBackslashes = 0;
    for (wchar_t ch : argument) {
        if (ch == L'\\') {
            ++consecutiveBackslashes;
            result.push_back(ch);
            continue;
        }

        if (ch == L'"') {
            result.append(consecutiveBackslashes + 1, L'\\');
            consecutiveBackslashes = 0;
            result.push_back(ch);
            continue;
        }

        consecutiveBackslashes = 0;
        result.push_back(ch);
    }

    if (consecutiveBackslashes > 0) {
        result.append(consecutiveBackslashes, L'\\');
    }

    result.push_back(L'"');
    return result;
}

SIZE ResolveMenuIconSize(const IconCache::Reference& iconReference) {
    const int defaultWidth = GetSystemMetrics(SM_CXSMICON);
    const int defaultHeight = GetSystemMetrics(SM_CYSMICON);
    SIZE size{defaultWidth > 0 ? defaultWidth : 16, defaultHeight > 0 ? defaultHeight : 16};
    if (iconReference) {
        if (auto metrics = iconReference.GetMetrics()) {
            if (metrics->cx > 0 && metrics->cy > 0) {
                size = *metrics;
            }
        }
    }
    return size;
}

BYTE AverageColorChannel(BYTE a, BYTE b) {
    return static_cast<BYTE>((static_cast<int>(a) + static_cast<int>(b)) / 2);
}

Gdiplus::Color BrightenBreadcrumbColor(const Gdiplus::Color& color,
                                       bool isHot,
                                       bool isPressed,
                                       COLORREF highlightBackgroundColor) {
    if (!isHot && !isPressed) {
        return color;
    }

    const float blendFactor = isPressed ? 0.75f : 0.55f;
    const BYTE blendRed = GetRValue(highlightBackgroundColor);
    const BYTE blendGreen = GetGValue(highlightBackgroundColor);
    const BYTE blendBlue = GetBValue(highlightBackgroundColor);

    auto blendChannel = [&](BYTE base, BYTE blend) -> BYTE {
        const double result = static_cast<double>(base) +
                              (static_cast<double>(blend) - static_cast<double>(base)) * blendFactor;
        return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(result)), 0, 255));
    };

    return Gdiplus::Color(color.GetA(), blendChannel(color.GetR(), blendRed),
                          blendChannel(color.GetG(), blendGreen), blendChannel(color.GetB(), blendBlue));
}

double SrgbChannelToLinear(BYTE channel) {
    const double srgb = static_cast<double>(channel) / 255.0;
    if (srgb <= 0.04045) {
        return srgb / 12.92;
    }

    return std::pow((srgb + 0.055) / 1.055, 2.4);
}

double ComputeColorLuminance(COLORREF color) {
    const double r = SrgbChannelToLinear(GetRValue(color));
    const double g = SrgbChannelToLinear(GetGValue(color));
    const double b = SrgbChannelToLinear(GetBValue(color));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

double ComputeContrastRatio(double a, double b) {
    const double lighter = std::max(a, b);
    const double darker = std::min(a, b);
    return (lighter + 0.05) / (darker + 0.05);
}

COLORREF ChooseStatusBarTextColor(COLORREF topColor, COLORREF bottomColor) {
    const double topLuminance = ComputeColorLuminance(topColor);
    const double bottomLuminance = ComputeColorLuminance(bottomColor);

    const double blackLuminance = ComputeColorLuminance(RGB(0, 0, 0));
    const double whiteLuminance = ComputeColorLuminance(RGB(255, 255, 255));

    const double contrastBlackTop = ComputeContrastRatio(topLuminance, blackLuminance);
    const double contrastBlackBottom = ComputeContrastRatio(bottomLuminance, blackLuminance);
    const double contrastWhiteTop = ComputeContrastRatio(topLuminance, whiteLuminance);
    const double contrastWhiteBottom = ComputeContrastRatio(bottomLuminance, whiteLuminance);

    const double minContrastBlack = std::min(contrastBlackTop, contrastBlackBottom);
    const double minContrastWhite = std::min(contrastWhiteTop, contrastWhiteBottom);

    constexpr double kMinimumReadableContrast = 4.5;
    const bool blackReadable = minContrastBlack >= kMinimumReadableContrast;
    const bool whiteReadable = minContrastWhite >= kMinimumReadableContrast;

    if (blackReadable != whiteReadable) {
        return blackReadable ? RGB(0, 0, 0) : RGB(255, 255, 255);
    }

    if (blackReadable && whiteReadable) {
        return minContrastBlack >= minContrastWhite ? RGB(0, 0, 0) : RGB(255, 255, 255);
    }

    return minContrastBlack >= minContrastWhite ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

COLORREF ChooseAccentTextColor(COLORREF accent) {
    return ComputeColorLuminance(accent) > 0.55 ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

constexpr wchar_t kOpenInNewTabLabel[] = L"Open in new tab";
constexpr int kProgressGradientSampleWidth = 256;

struct BreadcrumbHookEntry {
    HHOOK hook = nullptr;
    std::vector<shelltabs::CExplorerBHO*> observers;
};

std::mutex g_breadcrumbHookMutex;
std::unordered_map<DWORD, BreadcrumbHookEntry> g_breadcrumbHooks;

const wchar_t* DescribeSurfaceKind(shelltabs::ExplorerSurfaceKind kind) {
    using shelltabs::ExplorerSurfaceKind;
    switch (kind) {
        case ExplorerSurfaceKind::ListView:
            return L"list view";
        case ExplorerSurfaceKind::Header:
            return L"header";
        case ExplorerSurfaceKind::Rebar:
            return L"rebar";
        case ExplorerSurfaceKind::Toolbar:
            return L"toolbar";
        case ExplorerSurfaceKind::Edit:
            return L"edit";
        case ExplorerSurfaceKind::Scrollbar:
            return L"scrollbar";
        case ExplorerSurfaceKind::DirectUi:
            return L"DirectUI host";
        default:
            return L"surface";
    }
}

}  // namespace

// --- CExplorerBHO private state (treat these as class members) ---

namespace shelltabs {

std::mutex CExplorerBHO::s_ensureTimerLock;
std::unordered_map<UINT_PTR, CExplorerBHO*> CExplorerBHO::s_ensureTimers;
std::mutex CExplorerBHO::s_openInNewTabTimerLock;
std::unordered_map<UINT_PTR, CExplorerBHO*> CExplorerBHO::s_openInNewTabTimers;

CExplorerBHO::CExplorerBHO() : m_refCount(1), m_paneHooks() {
    ModuleAddRef();
    m_bufferedPaintInitialized = SUCCEEDED(BufferedPaintInit());
    m_glowCoordinator.Configure(OptionsStore::Instance().Get());

    Gdiplus::GdiplusStartupInput gdiplusInput;
    if (Gdiplus::GdiplusStartup(&m_gdiplusToken, &gdiplusInput, nullptr) == Gdiplus::Ok) {
        m_gdiplusInitialized = true;
    } else {
        m_gdiplusToken = 0;
        LogMessage(LogLevel::Warning, L"Failed to initialize GDI+; breadcrumb gradient disabled");
    }
}

CExplorerBHO::~CExplorerBHO() {
    Disconnect();
    DestroyProgressGradientResources();
    m_glowSurfaces.clear();
    if (m_bufferedPaintInitialized) {
        BufferedPaintUnInit();
        m_bufferedPaintInitialized = false;
    }
    if (m_gdiplusInitialized) {
        Gdiplus::GdiplusShutdown(m_gdiplusToken);
        m_gdiplusInitialized = false;
        m_gdiplusToken = 0;
    }
    ModuleRelease();
}

void CExplorerBHO::CancelEnsureRetry(BandEnsureState& state) {
    if (!state.retryScheduled || state.timerId == 0) {
        state.retryScheduled = false;
        state.timerId = 0;
        return;
    }

    {
        std::lock_guard<std::mutex> lock(s_ensureTimerLock);
        s_ensureTimers.erase(state.timerId);
    }

    KillTimer(nullptr, state.timerId);
    state.timerId = 0;
    state.retryScheduled = false;
}

void CExplorerBHO::CancelAllEnsureRetries() {
    std::vector<UINT_PTR> timers;
    timers.reserve(m_bandEnsureStates.size());

    for (auto& entry : m_bandEnsureStates) {
        BandEnsureState& state = entry.second;
        if (state.timerId != 0) {
            timers.push_back(state.timerId);
            state.timerId = 0;
        }
        state.retryScheduled = false;
        state.retryDelayMs = 0;
    }

    if (!timers.empty()) {
        std::lock_guard<std::mutex> lock(s_ensureTimerLock);
        for (UINT_PTR timerId : timers) {
            s_ensureTimers.erase(timerId);
        }
    }

    for (UINT_PTR timerId : timers) {
        KillTimer(nullptr, timerId);
    }

    for (auto& entry : m_bandEnsureStates) {
        BandEnsureState& state = entry.second;
        state.lastOutcome = BandEnsureOutcome::Unknown;
        state.lastHresult = S_OK;
    }
}

void CExplorerBHO::ScheduleEnsureRetry(HWND hostWindow, BandEnsureState& state, HRESULT lastHr,
                                       BandEnsureOutcome outcome, const wchar_t* reason) {
    CancelEnsureRetry(state);

    DWORD nextDelay = state.retryDelayMs == 0 ? kEnsureRetryInitialDelayMs : state.retryDelayMs * 2;
    if (nextDelay > kEnsureRetryMaxDelayMs) {
        nextDelay = kEnsureRetryMaxDelayMs;
    }

    state.retryDelayMs = nextDelay;
    state.lastOutcome = outcome;
    state.lastHresult = lastHr;

    UINT_PTR timerId = SetTimer(nullptr, 0, nextDelay, &CExplorerBHO::EnsureBandTimerProc);
    if (timerId == 0) {
        const DWORD error = GetLastError();
        LogMessage(LogLevel::Error,
                   L"EnsureBandVisible: failed to schedule retry timer (delay=%u ms, error=%lu) for host=%p after hr=0x%08X",
                   nextDelay, error, hostWindow, lastHr);
        state.retryScheduled = false;
        state.timerId = 0;
        m_shouldRetryEnsure = true;
        return;
    }

    {
        std::lock_guard<std::mutex> lock(s_ensureTimerLock);
        s_ensureTimers[timerId] = this;
    }

    state.retryScheduled = true;
    state.timerId = timerId;

    const wchar_t* description = reason ? reason : L"ShowBrowserBar failure";
    LogMessage(LogLevel::Warning,
               L"EnsureBandVisible: %s (hr=0x%08X code=%lu) for host=%p; retry #%zu scheduled in %u ms",
               description, lastHr, HRESULT_CODE(lastHr), hostWindow, state.attemptCount + 1, nextDelay);
}

void CExplorerBHO::HandleEnsureBandTimer(UINT_PTR timerId) {
    HWND targetWindow = nullptr;

    for (auto& entry : m_bandEnsureStates) {
        BandEnsureState& state = entry.second;
        if (state.timerId == timerId) {
            state.timerId = 0;
            state.retryScheduled = false;
            targetWindow = entry.first;
            break;
        }
    }

    m_shouldRetryEnsure = true;

    if (!targetWindow) {
        return;
    }

    LogMessage(LogLevel::Info, L"EnsureBandVisible retry timer fired for host=%p", targetWindow);
    EnsureBandVisible();
}

void CALLBACK CExplorerBHO::EnsureBandTimerProc(HWND, UINT, UINT_PTR timerId, DWORD) {
    CExplorerBHO* instance = nullptr;

    {
        std::lock_guard<std::mutex> lock(s_ensureTimerLock);
        auto it = s_ensureTimers.find(timerId);
        if (it != s_ensureTimers.end()) {
            instance = it->second;
            s_ensureTimers.erase(it);
        }
    }

    KillTimer(nullptr, timerId);

    if (instance) {
        instance->HandleEnsureBandTimer(timerId);
    }
}

void CExplorerBHO::HandleOpenInNewTabTimer(UINT_PTR timerId) {
    if (m_openInNewTabTimerId != timerId) {
        return;
    }

    m_openInNewTabTimerId = 0;
    m_openInNewTabRetryScheduled = false;
    TryDispatchQueuedOpenInNewTabRequests();
}

void CExplorerBHO::ScheduleOpenInNewTabRetry() {
    if (m_openInNewTabRetryScheduled || m_openInNewTabQueue.empty()) {
        return;
    }

    UINT_PTR timerId = SetTimer(nullptr, 0, kOpenInNewTabRetryDelayMs, &CExplorerBHO::OpenInNewTabTimerProc);
    if (timerId == 0) {
        const DWORD error = GetLastError();
        LogMessage(LogLevel::Error,
                   L"Open In New Tab: failed to schedule retry timer (delay=%u ms, error=%lu)",
                   kOpenInNewTabRetryDelayMs, error);
        return;
    }

    {
        std::lock_guard<std::mutex> lock(s_openInNewTabTimerLock);
        s_openInNewTabTimers[timerId] = this;
    }

    m_openInNewTabRetryScheduled = true;
    m_openInNewTabTimerId = timerId;
}

void CExplorerBHO::CancelOpenInNewTabRetry() {
    if (!m_openInNewTabRetryScheduled || m_openInNewTabTimerId == 0) {
        m_openInNewTabRetryScheduled = false;
        m_openInNewTabTimerId = 0;
        return;
    }

    {
        std::lock_guard<std::mutex> lock(s_openInNewTabTimerLock);
        s_openInNewTabTimers.erase(m_openInNewTabTimerId);
    }

    KillTimer(nullptr, m_openInNewTabTimerId);
    m_openInNewTabRetryScheduled = false;
    m_openInNewTabTimerId = 0;
}

void CALLBACK CExplorerBHO::OpenInNewTabTimerProc(HWND, UINT, UINT_PTR timerId, DWORD) {
    CExplorerBHO* instance = nullptr;

    {
        std::lock_guard<std::mutex> lock(s_openInNewTabTimerLock);
        auto it = s_openInNewTabTimers.find(timerId);
        if (it != s_openInNewTabTimers.end()) {
            instance = it->second;
            s_openInNewTabTimers.erase(it);
        }
    }

    KillTimer(nullptr, timerId);

    if (instance) {
        instance->HandleOpenInNewTabTimer(timerId);
    }
}
IFACEMETHODIMP CExplorerBHO::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IObjectWithSite) {
        *object = static_cast<IObjectWithSite*>(this);
    } else if (riid == IID_IDispatch) {
        *object = static_cast<IDispatch*>(this);
    } else {
        *object = nullptr;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) CExplorerBHO::AddRef() {
    return static_cast<ULONG>(++m_refCount);
}

IFACEMETHODIMP_(ULONG) CExplorerBHO::Release() {
    const ULONG count = static_cast<ULONG>(--m_refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}

IFACEMETHODIMP CExplorerBHO::GetTypeInfoCount(UINT* pctinfo) {
    if (!pctinfo) {
        return E_POINTER;
    }

    const auto typeInfo = LoadBrowserEventsTypeInfo();
    *pctinfo = typeInfo ? 1 : 0;
    return S_OK;
}

IFACEMETHODIMP CExplorerBHO::GetTypeInfo(UINT iTInfo, LCID, ITypeInfo** ppTInfo) {
    if (!ppTInfo) {
        return E_POINTER;
    }

    if (iTInfo != 0) {
        return DISP_E_BADINDEX;
    }

    const auto typeInfo = LoadBrowserEventsTypeInfo();
    if (!typeInfo) {
        *ppTInfo = nullptr;
        return TYPE_E_ELEMENTNOTFOUND;
    }

    return typeInfo.CopyTo(ppTInfo);
}

IFACEMETHODIMP CExplorerBHO::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID, DISPID* rgDispId) {
    if (riid != IID_NULL) {
        return DISP_E_UNKNOWNINTERFACE;
    }
    if (!rgszNames || !rgDispId) {
        return E_POINTER;
    }

    const auto typeInfo = LoadBrowserEventsTypeInfo();
    if (!typeInfo) {
        return TYPE_E_ELEMENTNOTFOUND;
    }

    return typeInfo->GetIDsOfNames(rgszNames, cNames, reinterpret_cast<MEMBERID*>(rgDispId));
}

void CExplorerBHO::Disconnect() {
    CancelAllEnsureRetries();
    CancelOpenInNewTabRetry();
    m_bandEnsureStates.clear();
    m_openInNewTabQueue.clear();
    RemoveBreadcrumbHook();
    RemoveBreadcrumbSubclass();
    RemoveProgressSubclass();
    RemoveAddressEditSubclass();
    RemoveExplorerViewSubclass();
    RemoveStatusBarSubclass();
    ResetStatusBarTheme();
    m_statusBar = nullptr;
    DisconnectEvents();
    m_webBrowser.Reset();
    m_shellBrowser.Reset();
    m_site.Reset();
    m_bandVisible = false;
    m_shouldRetryEnsure = true;
    m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    m_loggedBreadcrumbToolbarMissing = false;
    m_lastBreadcrumbStage = BreadcrumbDiscoveryStage::None;
    ClearFolderBackgrounds();
    m_currentFolderKey.clear();
}


HRESULT CExplorerBHO::EnsureBandVisible() {
    return GuardExplorerCall(
        L"CExplorerBHO::EnsureBandVisible",
        [&]() -> HRESULT {
            if (!m_webBrowser) {
                return S_OK;
            }

            HWND hostWindow = GetTopLevelExplorerWindow();
            BandEnsureState& state = m_bandEnsureStates[hostWindow];

            if (!m_shouldRetryEnsure) {
                return S_OK;
            }

            if (state.lastOutcome == BandEnsureOutcome::Success ||
                state.lastOutcome == BandEnsureOutcome::PermanentFailure) {
                m_shouldRetryEnsure = false;
                return S_OK;
            }

            if (state.retryScheduled) {
                m_shouldRetryEnsure = false;
                return S_OK;
            }

            m_shouldRetryEnsure = false;

            Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
            HRESULT hr = m_webBrowser.As(&serviceProvider);
            if ((!serviceProvider || FAILED(hr)) && m_site) {
                serviceProvider = nullptr;
                hr = m_site.As(&serviceProvider);
            }
            if (FAILED(hr) || !serviceProvider) {
                HRESULT failure = FAILED(hr) ? hr : E_NOINTERFACE;
                LogMessage(LogLevel::Warning,
                           L"EnsureBandVisible: IServiceProvider unavailable for host=%p (hr=0x%08X)",
                           hostWindow, failure);
                m_bandVisible = false;
                ScheduleEnsureRetry(hostWindow, state, failure, BandEnsureOutcome::TemporaryFailure,
                                     L"IServiceProvider unavailable");
                return failure;
            }

            Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
            hr = serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&shellBrowser));
            if ((FAILED(hr) || !shellBrowser)) {
                hr = serviceProvider->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&shellBrowser));
                if (FAILED(hr) || !shellBrowser) {
                    HRESULT failure = FAILED(hr) ? hr : E_NOINTERFACE;
                    LogMessage(LogLevel::Warning,
                               L"EnsureBandVisible: IShellBrowser unavailable for host=%p (hr=0x%08X)",
                               hostWindow, failure);
                    m_bandVisible = false;
                    ScheduleEnsureRetry(hostWindow, state, failure, BandEnsureOutcome::TemporaryFailure,
                                         L"IShellBrowser unavailable");
                    return failure;
                }
            }

            bool supportedHost = false;
            wchar_t className[256] = {};
            bool hasClassInformation = false;
            if (hostWindow && IsWindow(hostWindow)) {
                const int length = GetClassNameW(hostWindow, className, ARRAYSIZE(className));
                if (length > 0) {
                    hasClassInformation = true;
                    supportedHost = (_wcsicmp(className, L"CabinetWClass") == 0);
                }
            }

            HRESULT explorerBrowserHr = E_FAIL;
            bool explorerBrowserAvailable = false;
            Microsoft::WRL::ComPtr<IUnknown> explorerBrowser;
            explorerBrowserHr = serviceProvider->QueryService(CLSID_ExplorerBrowser, IID_PPV_ARGS(&explorerBrowser));
            if (SUCCEEDED(explorerBrowserHr) && explorerBrowser) {
                explorerBrowserAvailable = true;
                supportedHost = true;
            }

            if (!supportedHost) {
                if (!hasClassInformation && !explorerBrowserAvailable) {
                    if (!state.unsupportedHost) {
                        LogMessage(LogLevel::Warning,
                                   L"EnsureBandVisible: delaying band activation; host=%p not yet classified (ExplorerBrowser hr=0x%08X)",
                                   hostWindow, explorerBrowserHr);
                    }
                    m_bandVisible = false;
                    const HRESULT classificationHr = FAILED(explorerBrowserHr) ? explorerBrowserHr : E_FAIL;
                    ScheduleEnsureRetry(hostWindow, state, classificationHr, BandEnsureOutcome::TemporaryFailure,
                                         L"Explorer host classification pending");
                    return classificationHr;
                }

                if (!state.unsupportedHost) {
                    if (hasClassInformation) {
                        LogMessage(LogLevel::Warning,
                                   L"EnsureBandVisible: host=%p uses unsupported class '%s'; ExplorerBrowser hr=0x%08X",
                                   hostWindow, className, explorerBrowserHr);
                    } else {
                        LogMessage(LogLevel::Warning,
                                   L"EnsureBandVisible: host=%p exposes ExplorerBrowser hr=0x%08X but remains unsupported",
                                   hostWindow, explorerBrowserHr);
                    }
                }

                CancelEnsureRetry(state);
                state.unsupportedHost = true;
                state.lastOutcome = BandEnsureOutcome::PermanentFailure;
                state.lastHresult = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                state.retryDelayMs = 0;
                m_bandVisible = false;
                return state.lastHresult;
            }

            state.unsupportedHost = false;

            const std::wstring clsidString = GuidToString(CLSID_ShellTabsBand);
            if (clsidString.empty()) {
                LogMessage(LogLevel::Error, L"EnsureBandVisible: failed to stringify CLSID_ShellTabsBand");
                m_bandVisible = false;
                ScheduleEnsureRetry(hostWindow, state, E_FAIL, BandEnsureOutcome::TemporaryFailure,
                                     L"Failed to format band CLSID");
                return E_FAIL;
            }

            VARIANT bandId;
            VariantInit(&bandId);
            bandId.vt = VT_BSTR;
            bandId.bstrVal = SysAllocString(clsidString.c_str());
            if (!bandId.bstrVal) {
                LogMessage(LogLevel::Error, L"EnsureBandVisible: SysAllocString failed for band CLSID");
                m_bandVisible = false;
                ScheduleEnsureRetry(hostWindow, state, E_OUTOFMEMORY, BandEnsureOutcome::TemporaryFailure,
                                     L"SysAllocString failed for band CLSID");
                return E_OUTOFMEMORY;
            }

            VARIANT show;
            VariantInit(&show);
            show.vt = VT_BOOL;
            show.boolVal = VARIANT_TRUE;

            const size_t attempt = ++state.attemptCount;
            LogMessage(LogLevel::Info, L"EnsureBandVisible: invoking ShowBrowserBar for host=%p (attempt %zu)",
                       hostWindow, attempt);
            hr = m_webBrowser->ShowBrowserBar(&bandId, &show, nullptr);

            VariantClear(&bandId);
            VariantClear(&show);

            if (SUCCEEDED(hr)) {
                m_bandVisible = true;
                CancelEnsureRetry(state);
                state.retryDelayMs = 0;
                state.lastOutcome = BandEnsureOutcome::Success;
                state.lastHresult = hr;
                LogMessage(LogLevel::Info,
                           L"EnsureBandVisible: ShowBrowserBar succeeded for host=%p on attempt %zu", hostWindow,
                           attempt);
                UpdateBreadcrumbSubclass();
                TryDispatchQueuedOpenInNewTabRequests();
            } else {
                m_bandVisible = false;
                const bool throttled = IsShowBrowserBarThrottled(hr);

                if (throttled) {
                    LogMessage(LogLevel::Warning,
                               L"EnsureBandVisible: ShowBrowserBar throttled for host=%p on attempt %zu (hr=0x%08X code=%lu)",
                               hostWindow, attempt, hr, HRESULT_CODE(hr));
                    ScheduleEnsureRetry(hostWindow, state, hr, BandEnsureOutcome::Throttled,
                                         L"ShowBrowserBar throttled");
                } else if (hr == E_ACCESSDENIED || HRESULT_CODE(hr) == ERROR_ACCESS_DENIED) {
                    CancelEnsureRetry(state);
                    state.retryDelayMs = 0;
                    state.lastOutcome = BandEnsureOutcome::PermanentFailure;
                    state.lastHresult = hr;
                    LogMessage(LogLevel::Error,
                               L"EnsureBandVisible: ShowBrowserBar denied access for host=%p (hr=0x%08X); stopping retries",
                               hostWindow, hr);
                } else if (IsAutomationDisabledResult(hr)) {
                    CancelEnsureRetry(state);
                    state.retryDelayMs = 0;
                    state.lastOutcome = BandEnsureOutcome::PermanentFailure;
                    state.lastHresult = hr;
                    LogMessage(LogLevel::Error,
                               L"EnsureBandVisible: automation disabled by policy for host=%p (hr=0x%08X code=%lu)",
                               hostWindow, hr, HRESULT_CODE(hr));
                    NotifyAutomationDisabledByPolicy(hr);
                } else {
                    ScheduleEnsureRetry(hostWindow, state, hr, BandEnsureOutcome::TemporaryFailure,
                                         L"ShowBrowserBar failed");
                }
            }

            return hr;
        },
        []() -> HRESULT { return E_FAIL; });
}

IFACEMETHODIMP CExplorerBHO::SetSite(IUnknown* site) {
    return GuardExplorerCall(
        L"CExplorerBHO::SetSite",
        [&]() -> HRESULT {
            if (!site) {
                LogMessage(LogLevel::Info, L"CExplorerBHO::SetSite detaching from site");
                Disconnect();
                return S_OK;
            }

            LogMessage(LogLevel::Info, L"CExplorerBHO::SetSite attaching to site=%p", site);
            Disconnect();

            Microsoft::WRL::ComPtr<IWebBrowser2> browser;
            HRESULT hr = ResolveBrowserFromSite(site, &browser);
            if (FAILED(hr) || !browser) {
                return S_OK;
            }

            m_site = site;
            m_webBrowser = browser;
            m_shouldRetryEnsure = true;

            m_shellBrowser.Reset();

            ConnectEvents();

            Microsoft::WRL::ComPtr<IServiceProvider> siteProvider;
            if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&siteProvider))) && siteProvider) {
                if (!m_shellBrowser) {
                    siteProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&m_shellBrowser));
                }
                if (!m_shellBrowser) {
                    siteProvider->QueryService(SID_SShellBrowser, IID_PPV_ARGS(&m_shellBrowser));
                }
            } else {
                Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
                if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&shellBrowser))) && shellBrowser) {
                    m_shellBrowser = shellBrowser;
                    m_shellBrowser.As(&siteProvider);
                }
            }

            if (!m_shellBrowser) {
                Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
                if (SUCCEEDED(site->QueryInterface(IID_PPV_ARGS(&shellBrowser))) && shellBrowser) {
                    m_shellBrowser = shellBrowser;
                }
            }

            if (!siteProvider && m_shellBrowser) {
                m_shellBrowser.As(&siteProvider);
            }
            EnsureBandVisible();
            UpdateBreadcrumbSubclass();
            UpdateExplorerViewSubclass();
            return S_OK;

        },
        []() -> HRESULT { return E_FAIL; });
}

HRESULT CExplorerBHO::ResolveBrowserFromSite(IUnknown* site, IWebBrowser2** browser) {
    if (!browser) {
        return E_POINTER;
    }

    *browser = nullptr;

    if (!site) {
        return E_POINTER;
    }

    Microsoft::WRL::ComPtr<IWebBrowser2> candidate;
    HRESULT hr = site->QueryInterface(IID_PPV_ARGS(&candidate));
    if (SUCCEEDED(hr) && candidate) {
        *browser = candidate.Detach();
        return S_OK;
    }

    Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
    hr = site->QueryInterface(IID_PPV_ARGS(&serviceProvider));
    if (SUCCEEDED(hr) && serviceProvider) {
        hr = serviceProvider->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&candidate));
        if (SUCCEEDED(hr) && candidate) {
            *browser = candidate.Detach();
            return S_OK;
        }

        hr = serviceProvider->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&candidate));
        if (SUCCEEDED(hr) && candidate) {
            *browser = candidate.Detach();
            return S_OK;
        }
    }

    Microsoft::WRL::ComPtr<IShellBrowser> shellBrowser;
    hr = site->QueryInterface(IID_PPV_ARGS(&shellBrowser));
    if (SUCCEEDED(hr) && shellBrowser) {
        serviceProvider = nullptr;
        hr = shellBrowser.As(&serviceProvider);
        if (SUCCEEDED(hr) && serviceProvider) {
            hr = serviceProvider->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&candidate));
            if (SUCCEEDED(hr) && candidate) {
                *browser = candidate.Detach();
                return S_OK;
            }
        }
    }

    return E_NOINTERFACE;
}

IFACEMETHODIMP CExplorerBHO::GetSite(REFIID riid, void** site) {
    return GuardExplorerCall(
        L"CExplorerBHO::GetSite",
        [&]() -> HRESULT {
            if (!site) {
                return E_POINTER;
            }
            *site = nullptr;
            if (!m_site) {
                return E_FAIL;
            }
            return m_site->QueryInterface(riid, site);
        },
        []() -> HRESULT { return E_FAIL; });
}

CExplorerBHO::TreeItemPidlResolution CExplorerBHO::ResolveTreeViewItemPidl(HWND treeView,
                                                                           const TVITEMEXW& item) const {
    TreeItemPidlResolution resolved;

    if (!item.hItem) {
        return resolved;
    }

    if (treeView && m_namespaceTreeControl) {
        RECT itemBounds{};
        HTREEITEM handle = item.hItem;
        if (TreeView_GetItemRect(treeView, handle, &itemBounds, TRUE)) {
            const LONG centerX = static_cast<LONG>(itemBounds.left + (itemBounds.right - itemBounds.left) / 2);
            const LONG centerY = static_cast<LONG>(itemBounds.top + (itemBounds.bottom - itemBounds.top) / 2);
            POINT queryPoint{centerX, centerY};

            Microsoft::WRL::ComPtr<IShellItem> shellItem;
            HRESULT hr = m_namespaceTreeControl->HitTest(&queryPoint, &shellItem);
            if (SUCCEEDED(hr) && shellItem) {
                PIDLIST_ABSOLUTE pidl = nullptr;
                hr = SHGetIDListFromObject(shellItem.Get(), &pidl);
                if (SUCCEEDED(hr) && pidl) {
                    resolved.owned.reset(reinterpret_cast<ITEMIDLIST*>(pidl));
                    resolved.raw = resolved.owned.get();
                    return resolved;
                }
            }
        }
    }

    resolved.raw = reinterpret_cast<PCIDLIST_ABSOLUTE>(item.lParam);
    return resolved;
}

HRESULT CExplorerBHO::ConnectEvents() {
    return GuardExplorerCall(
        L"CExplorerBHO::ConnectEvents",
        [&]() -> HRESULT {
            if (!m_webBrowser || m_connectionCookie != 0) {
                return S_OK;
            }

            Microsoft::WRL::ComPtr<IConnectionPointContainer> container;
            HRESULT hr = m_webBrowser.As(&container);
            if (FAILED(hr) || !container) {
                return hr;
            }

            Microsoft::WRL::ComPtr<IConnectionPoint> connectionPoint;
            hr = container->FindConnectionPoint(DIID_DWebBrowserEvents2, &connectionPoint);
            if (FAILED(hr) || !connectionPoint) {
                return hr;
            }

            DWORD cookie = 0;
            hr = connectionPoint->Advise(static_cast<IDispatch*>(this), &cookie);
            if (FAILED(hr)) {
                return hr;
            }

            m_connectionPoint = connectionPoint;
            m_connectionCookie = cookie;
            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

void CExplorerBHO::DisconnectEvents() {
    if (m_connectionPoint && m_connectionCookie != 0) {
        m_connectionPoint->Unadvise(m_connectionCookie);
    }
    m_connectionPoint.Reset();
    m_connectionCookie = 0;
}

IFACEMETHODIMP CExplorerBHO::Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*,
                                    UINT*) {
    return GuardExplorerCall(
        L"CExplorerBHO::Invoke",
        [&]() -> HRESULT {
            switch (dispIdMember) {
                case DISPID_ONVISIBLE:
                case DISPID_WINDOWSTATECHANGED:
                    if (!m_bandVisible) {
                        m_shouldRetryEnsure = true;
                        EnsureBandVisible();
                        UpdateBreadcrumbSubclass();
                    }
                    break;
                case DISPID_DOCUMENTCOMPLETE:
                case DISPID_NAVIGATECOMPLETE2:
                    UpdateBreadcrumbSubclass();
                    UpdateExplorerViewSubclass();
                    break;
                case DISPID_ONQUIT:
                    Disconnect();
                    break;
                default:
                    break;
            }

            return S_OK;
        },
        []() -> HRESULT { return E_FAIL; });
}

HWND CExplorerBHO::GetTopLevelExplorerWindow() const {
    HWND hwnd = nullptr;
    if (m_shellBrowser && SUCCEEDED(m_shellBrowser->GetWindow(&hwnd)) && hwnd) {
        // fall through to normalize the window handle below
    } else if (m_webBrowser) {
        SHANDLE_PTR raw = 0;
        if (SUCCEEDED(m_webBrowser->get_HWND(&raw)) && raw) {
            hwnd = reinterpret_cast<HWND>(raw);
        }
    }

    if (!hwnd) {
        return nullptr;
    }

    HWND ancestor = GetAncestor(hwnd, GA_ROOTOWNER);
    if (ancestor) {
        hwnd = ancestor;
    }

    ancestor = GetAncestor(hwnd, GA_ROOT);
    if (ancestor) {
        hwnd = ancestor;
    }

    // Walk up the parent chain in case GetAncestor returned a child window.
    HWND current = hwnd;
    HWND parent = nullptr;
    int safety = 0;
    while (current && safety++ < 32) {
        parent = GetParent(current);
        if (!parent) {
            break;
        }
        current = parent;
    }

    return current ? current : hwnd;
}

void CExplorerBHO::LogBreadcrumbStage(BreadcrumbDiscoveryStage stage, const wchar_t* format, ...) const {
    if (!format) {
        return;
    }
    if (m_lastBreadcrumbStage == stage) {
        return;
    }

    m_lastBreadcrumbStage = stage;

    va_list args;
    va_start(args, format);
    LogMessageV(LogLevel::Info, format, args);
    va_end(args);
}

HWND CExplorerBHO::FindBreadcrumbToolbar() const {
    auto queryBreadcrumbToolbar = [&](const Microsoft::WRL::ComPtr<IServiceProvider>& provider,
                                     const wchar_t* source) -> HWND {
        if (!provider) {
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IUnknown> breadcrumbService;
        HRESULT hr = provider->QueryService(CLSID_CBreadcrumbBar, IID_PPV_ARGS(&breadcrumbService));
        if (FAILED(hr) || !breadcrumbService) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceUnavailable,
                               L"Breadcrumb QueryService(%s) failed: 0x%08X", source ? source : L"?", hr);
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
        hr = breadcrumbService.As(&oleWindow);
        if (FAILED(hr) || !oleWindow) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceWindowMissing,
                               L"Breadcrumb service missing IOleWindow (%s): 0x%08X", source ? source : L"?", hr);
            return nullptr;
        }

        HWND bandWindow = nullptr;
        hr = oleWindow->GetWindow(&bandWindow);
        if (FAILED(hr) || !bandWindow) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceWindowMissing,
                               L"Breadcrumb service window unavailable (%s): 0x%08X", source ? source : L"?", hr);
            return nullptr;
        }

        HWND toolbar = FindWindowExW(bandWindow, nullptr, TOOLBARCLASSNAME, nullptr);
        if (!toolbar) {
            toolbar = FindDescendantWindow(bandWindow, TOOLBARCLASSNAME);
        }
        if (toolbar) {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                               L"Breadcrumb toolbar located via %s service (hwnd=%p)", source ? source : L"?",
                               toolbar);
        } else {
            LogBreadcrumbStage(BreadcrumbDiscoveryStage::ServiceToolbarMissing,
                               L"Breadcrumb service band (%s hwnd=%p) missing toolbar child", source ? source : L"?",
                               bandWindow);
        }
        return toolbar;
    };

    auto probeAdditionalProviders = [&](const Microsoft::WRL::ComPtr<IServiceProvider>& provider,
                                        const wchar_t* source) -> HWND {
        if (!provider) {
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IUnknown> frameService;
        HRESULT hr = provider->QueryService(SID_STopLevelBrowserFrame, IID_PPV_ARGS(&frameService));
        if (FAILED(hr) || !frameService) {
            return nullptr;
        }

        Microsoft::WRL::ComPtr<IOleWindow> frameWindow;
        if (SUCCEEDED(frameService.As(&frameWindow)) && frameWindow) {
            HWND frameHwnd = nullptr;
            if (SUCCEEDED(frameWindow->GetWindow(&frameHwnd)) && frameHwnd) {
                LogMessage(LogLevel::Info,
                           L"Breadcrumb ribbon frame discovered via %s (hwnd=%p)", source ? source : L"?",
                           frameHwnd);
                if (HWND fromWindow = FindBreadcrumbToolbarInWindow(frameHwnd)) {
                    return fromWindow;
                }
            }
        }

        Microsoft::WRL::ComPtr<IServiceProvider> nestedProvider;
        if (SUCCEEDED(frameService.As(&nestedProvider)) && nestedProvider) {
            if (HWND fromNested = queryBreadcrumbToolbar(nestedProvider, L"RibbonFrame")) {
                return fromNested;
            }
        }

        return nullptr;
    };

    if (m_shellBrowser) {
        Microsoft::WRL::ComPtr<IServiceProvider> provider;
        if (SUCCEEDED(m_shellBrowser.As(&provider))) {
            if (HWND fromService = queryBreadcrumbToolbar(provider, L"IShellBrowser")) {
                return fromService;
            }
            if (HWND fromFrame = probeAdditionalProviders(provider, L"IShellBrowser")) {
                return fromFrame;
            }
        }
    }

    if (m_webBrowser) {
        Microsoft::WRL::ComPtr<IServiceProvider> provider;
        if (SUCCEEDED(m_webBrowser.As(&provider))) {
            if (HWND fromService = queryBreadcrumbToolbar(provider, L"IWebBrowser2")) {
                return fromService;
            }
            if (HWND fromFrame = probeAdditionalProviders(provider, L"IWebBrowser2")) {
                return fromFrame;
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::FrameMissing,
                           L"Top-level Explorer window unavailable during breadcrumb search");
        return nullptr;
    }

    HWND travelBand = FindDescendantWindow(frame, L"TravelBand");
    HWND rebar = travelBand ? GetParent(travelBand) : nullptr;
    if (!rebar) {
        rebar = FindDescendantWindow(frame, L"ReBarWindow32");
    }
    if (!rebar) {
        DWORD threadId = GetWindowThreadProcessId(frame, nullptr);
        if (threadId != 0) {
            struct EnumData {
                HWND rebar = nullptr;
            } data{};
            EnumThreadWindows(
                threadId,
                [](HWND hwnd, LPARAM param) -> BOOL {
                    auto* data = reinterpret_cast<EnumData*>(param);
                    if (!data) {
                        return FALSE;
                    }
                    if (MatchesClass(hwnd, L"ReBarWindow32")) {
                        data->rebar = hwnd;
                        return FALSE;
                    }
                    return TRUE;
                },
                reinterpret_cast<LPARAM>(&data));

            if (data.rebar) {
                LogMessage(LogLevel::Info, L"Breadcrumb rebar located via thread scan (hwnd=%p)", data.rebar);
                rebar = data.rebar;
            }
        }
    }
    if (!rebar) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::RebarMissing,
                           L"Failed to locate Explorer rebar while searching for breadcrumbs");
        return FindBreadcrumbToolbarInWindow(frame);
    }

    HWND breadcrumbParent = FindWindowExW(rebar, nullptr, L"Breadcrumb Parent", nullptr);
    if (!breadcrumbParent) {
        breadcrumbParent = FindDescendantWindow(rebar, L"Breadcrumb Parent");
    }
    if (!breadcrumbParent) {
        breadcrumbParent = FindDescendantWindow(frame, L"Breadcrumb Parent");
    }
    if (!breadcrumbParent) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::ParentMissing,
                           L"Failed to find 'Breadcrumb Parent' window during breadcrumb search");
        return FindBreadcrumbToolbarInWindow(frame);
    }

    HWND toolbar = FindWindowExW(breadcrumbParent, nullptr, TOOLBARCLASSNAME, nullptr);
    if (!toolbar) {
        toolbar = FindDescendantWindow(breadcrumbParent, TOOLBARCLASSNAME);
    }
    if (!toolbar) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::ToolbarMissing,
                           L"'Breadcrumb Parent' hwnd=%p missing ToolbarWindow32 child", breadcrumbParent);
        return FindBreadcrumbToolbarInWindow(breadcrumbParent);
    }

    LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                       L"Breadcrumb toolbar located via window enumeration (hwnd=%p)", toolbar);
    return toolbar;
}

HWND CExplorerBHO::FindBreadcrumbToolbarInWindow(HWND root) const {
    if (!root) {
        return nullptr;
    }

    struct EnumData {
        const CExplorerBHO* self = nullptr;
        HWND toolbar = nullptr;
    } data{this, nullptr};

    EnumChildWindows(
        root,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* data = reinterpret_cast<EnumData*>(param);
            if (!data || data->toolbar) {
                return FALSE;
            }
            if (!MatchesClass(hwnd, TOOLBARCLASSNAME)) {
                return TRUE;
            }
            if (!data->self->IsBreadcrumbToolbarCandidate(hwnd)) {
                return TRUE;
            }
            data->toolbar = hwnd;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    if (data.toolbar) {
        LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                           L"Breadcrumb toolbar located via deep enumeration (hwnd=%p)", data.toolbar);
    }

    return data.toolbar;
}

HWND CExplorerBHO::FindProgressWindow() const {
    if (m_breadcrumbToolbar && IsWindow(m_breadcrumbToolbar)) {
        HWND breadcrumbParent = GetParent(m_breadcrumbToolbar);
        if (breadcrumbParent) {
            if (MatchesClass(breadcrumbParent, PROGRESS_CLASSW)) {
                return breadcrumbParent;
            }
            HWND progressParent = GetParent(breadcrumbParent);
            if (progressParent && MatchesClass(progressParent, PROGRESS_CLASSW)) {
                return progressParent;
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return nullptr;
    }

    struct EnumData {
        HWND progress = nullptr;
    } data{};

    EnumChildWindows(
        frame,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* data = reinterpret_cast<EnumData*>(param);
            if (!data || data->progress) {
                return FALSE;
            }
            if (!MatchesClass(hwnd, PROGRESS_CLASSW)) {
                return TRUE;
            }
            if (!FindDescendantWindow(hwnd, L"Breadcrumb Parent")) {
                return TRUE;
            }
            data->progress = hwnd;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    return data.progress;
}

HWND CExplorerBHO::FindAddressEditControl() const {
    auto resolveEdit = [&](HWND window) -> HWND {
        if (!window || !IsWindow(window)) {
            return nullptr;
        }
        HWND edit = nullptr;
        if (MatchesClass(window, L"ComboBoxEx32")) {
            edit = reinterpret_cast<HWND>(SendMessageW(window, CBEM_GETEDITCONTROL, 0, 0));
            if (!edit) {
                edit = FindDescendantWindow(window, L"Edit");
            }
        } else if (MatchesClass(window, L"Edit")) {
            edit = window;
        } else {
            edit = FindDescendantWindow(window, L"Edit");
        }
        if (!edit || !IsWindow(edit)) {
            return nullptr;
        }
        if (!MatchesClass(edit, L"Edit")) {
            return nullptr;
        }
        if (!IsBreadcrumbToolbarAncestor(edit) || !IsWindowOwnedByThisExplorer(edit)) {
            return nullptr;
        }
        return edit;
    };

    if (m_breadcrumbToolbar && IsWindow(m_breadcrumbToolbar)) {
        if (HWND parent = GetParent(m_breadcrumbToolbar)) {
            if (HWND edit = resolveEdit(parent)) {
                return edit;
            }
            if (HWND grandparent = GetParent(parent)) {
                if (HWND edit = resolveEdit(grandparent)) {
                    return edit;
                }
            }
        }
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return nullptr;
    }

    struct EnumData {
        const CExplorerBHO* self = nullptr;
        HWND edit = nullptr;
    } data{this, nullptr};

    EnumChildWindows(
        frame,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* data = reinterpret_cast<EnumData*>(param);
            if (!data || data->edit) {
                return FALSE;
            }
            if (!MatchesClass(hwnd, L"ComboBoxEx32") && !MatchesClass(hwnd, L"Edit")) {
                return TRUE;
            }
            HWND edit = nullptr;
            if (MatchesClass(hwnd, L"ComboBoxEx32")) {
                edit = reinterpret_cast<HWND>(SendMessageW(hwnd, CBEM_GETEDITCONTROL, 0, 0));
                if (!edit) {
                    edit = FindDescendantWindow(hwnd, L"Edit");
                }
            } else {
                edit = hwnd;
            }
            if (!edit || !IsWindow(edit) || !MatchesClass(edit, L"Edit")) {
                return TRUE;
            }
            if (!data->self->IsBreadcrumbToolbarAncestor(edit) ||
                !data->self->IsWindowOwnedByThisExplorer(edit)) {
                return TRUE;
            }
            data->edit = edit;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&data));

    return data.edit;
}

std::vector<HWND> CExplorerBHO::FindExplorerEditControls() const {
    std::vector<HWND> edits;
    std::unordered_set<HWND, HandleHasher> seen;

    if (HWND address = FindAddressEditControl()) {
        MaybeAddExplorerEdit(address, seen, edits);
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return edits;
    }

    struct EnumContext {
        const CExplorerBHO* self = nullptr;
        std::unordered_set<HWND, HandleHasher>* seen = nullptr;
        std::vector<HWND>* edits = nullptr;
    } context{this, &seen, &edits};

    EnumChildWindows(
        frame,
        [](HWND hwnd, LPARAM param) -> BOOL {
            auto* ctx = reinterpret_cast<EnumContext*>(param);
            if (!ctx || !ctx->self || !ctx->seen || !ctx->edits) {
                return TRUE;
            }
            if (!IsWindow(hwnd)) {
                return TRUE;
            }
            if (MatchesClass(hwnd, L"DirectUIHWND")) {
                ctx->self->EnumerateDirectUIEditChildren(hwnd, *ctx->seen, *ctx->edits);
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&context));

    return edits;
}

void CExplorerBHO::EnumerateDirectUIEditChildren(HWND root, std::unordered_set<HWND, HandleHasher>& seen,
                                                 std::vector<HWND>& edits) const {
    if (!root || !IsWindow(root)) {
        return;
    }

    struct EnumContext {
        const CExplorerBHO* self = nullptr;
        std::unordered_set<HWND, HandleHasher>* seen = nullptr;
        std::vector<HWND>* edits = nullptr;
    } context{this, &seen, &edits};

    EnumChildWindows(
        root,
        [](HWND child, LPARAM param) -> BOOL {
            auto* ctx = reinterpret_cast<EnumContext*>(param);
            if (!ctx || !ctx->self || !ctx->seen || !ctx->edits) {
                return TRUE;
            }
            if (!IsWindow(child)) {
                return TRUE;
            }
            if (MatchesClass(child, L"Edit")) {
                ctx->self->MaybeAddExplorerEdit(child, *ctx->seen, *ctx->edits);
                return TRUE;
            }
            if (MatchesClass(child, L"DirectUIHWND")) {
                ctx->self->EnumerateDirectUIEditChildren(child, *ctx->seen, *ctx->edits);
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&context));
}

void CExplorerBHO::MaybeAddExplorerEdit(HWND candidate, std::unordered_set<HWND, HandleHasher>& seen,
                                        std::vector<HWND>& edits) const {
    if (!candidate || !IsWindow(candidate)) {
        return;
    }
    if (!MatchesClass(candidate, L"Edit")) {
        return;
    }
    if (!IsWindowOwnedByThisExplorer(candidate)) {
        return;
    }
    if (!IsExplorerEditAncestor(candidate)) {
        return;
    }
    if (seen.insert(candidate).second) {
        edits.push_back(candidate);
    }
}

bool CExplorerBHO::IsBreadcrumbToolbarAncestor(HWND hwnd) const {
    HWND current = hwnd;
    bool sawRebar = false;
    int depth = 0;
    while (current && depth++ < 16) {
        if (MatchesClass(current, L"Breadcrumb Parent") || MatchesClass(current, L"Address Band Root") ||
            MatchesClass(current, L"AddressBandRoot") || MatchesClass(current, L"CabinetAddressBand") ||
            MatchesClass(current, L"NavigationBand")) {
            return true;
        }
        if (MatchesClass(current, L"ReBarWindow32")) {
            sawRebar = true;
        }
        if (MatchesClass(current, L"CabinetWClass")) {
            break;
        }
        current = GetParent(current);
    }
    return sawRebar;
}

bool CExplorerBHO::IsExplorerEditAncestor(HWND hwnd) const {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    if (IsBreadcrumbToolbarAncestor(hwnd)) {
        return true;
    }

    HWND current = hwnd;
    bool sawDirectUI = false;
    int depth = 0;
    while (current && depth++ < 32) {
        if (MatchesClass(current, L"DirectUIHWND")) {
            sawDirectUI = true;
        }
        if (MatchesClass(current, L"ReBarWindow32")) {
            if (sawDirectUI) {
                return true;
            }
        }
        if (MatchesClass(current, L"CabinetWClass")) {
            break;
        }
        current = GetParent(current);
    }

    return sawDirectUI;
}

bool CExplorerBHO::IsBreadcrumbToolbarCandidate(HWND hwnd) const {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }
    if (!MatchesClass(hwnd, TOOLBARCLASSNAME)) {
        return false;
    }

    if (!IsBreadcrumbToolbarAncestor(hwnd)) {
        return false;
    }

    LRESULT buttonCount = SendMessage(hwnd, TB_BUTTONCOUNT, 0, 0);
    if (buttonCount <= 0) {
        return false;
    }

    const int maxToCheck = static_cast<int>(std::min<LRESULT>(buttonCount, 5));
    std::array<wchar_t, 260> buffer{};
    TBBUTTON button{};
    for (int i = 0; i < maxToCheck; ++i) {
        if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&button))) {
            continue;
        }
        if ((button.fsStyle & TBSTYLE_SEP) != 0 || (button.fsState & TBSTATE_HIDDEN) != 0) {
            continue;
        }

        buffer.fill(L'\0');
        LRESULT copied =
            SendMessage(hwnd, TB_GETBUTTONTEXTW, button.idCommand, reinterpret_cast<LPARAM>(buffer.data()));
        if (copied > 0 && buffer[0] != L'\0') {
            return true;
        }

        if (copied == -1) {
            // LPSTR_TEXTCALLBACK indicates the button supplies text dynamically.
            return true;
        }

        if (button.iString != 0) {
            // Non-null string pointers or string-pool indices also imply textual content.
            return true;
        }
    }

    return false;
}

bool CExplorerBHO::IsWindowOwnedByThisExplorer(HWND hwnd) const {
    HWND frame = GetTopLevelExplorerWindow();
    if (!frame || !IsWindow(frame)) {
        return false;
    }

    HWND current = hwnd;
    int depth = 0;
    while (current && depth++ < 32) {
        if (current == frame) {
            return true;
        }
        current = GetParent(current);
    }

    HWND root = GetAncestor(hwnd, GA_ROOT);
    return root == frame;
}

void CExplorerBHO::DetachListView() {
    HWND listView = m_listView;
    HWND controlWindow = m_listViewControlWindow;

    if (listView) {
        if (HWND header = ListView_GetHeader(listView)) {
            if (IsWindow(header)) {
                RemoveWindowSubclass(header, &CExplorerBHO::ExplorerViewSubclassProc,
                                     reinterpret_cast<UINT_PTR>(this));
            }
            UnregisterGlowSurface(header);
        }
    }

    if (listView && m_listViewSubclassInstalled && IsWindow(listView)) {
        RemoveWindowSubclass(listView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }

    if (listView) {
        UnregisterGlowSurface(listView);
    }

    if (controlWindow) {
        UnregisterGlowSurface(controlWindow);
    }

    m_listView = nullptr;
    m_listViewSubclassInstalled = false;
    m_listViewControlWindow = nullptr;

    m_listViewControl.reset();

    if (m_nativeListView && IsWindow(m_nativeListView)) {
        EnableWindow(m_nativeListView, TRUE);
        ShowWindow(m_nativeListView, SW_SHOW);
    }
    m_nativeListView = nullptr;
}

bool CExplorerBHO::AttachListView(HWND listView) {
    if (!listView || !IsWindow(listView)) {
        DetachListView();
        return false;
    }

    if (ShellTabsListView::IsShellTabsListView(listView)) {
        return m_listView == listView && m_listViewSubclassInstalled;
    }

    if (m_nativeListView == listView && m_listView && m_listViewSubclassInstalled && IsWindow(m_listView)) {
        return true;
    }

    DetachListView();

    HWND parent = GetParent(listView);
    if (!parent || !IsWindow(parent)) {
        return false;
    }

    auto control = std::make_unique<ShellTabsListView>();

    ShellTabsListView::HighlightResolver highlightResolver =
        [this](PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) {
            return ResolveHighlightFromPidl(pidl, highlight);
        };
    ShellTabsListView::BackgroundResolver backgroundResolver =
        [this]() -> ShellTabsListView::BackgroundSource {
        ShellTabsListView::BackgroundSource source{};
        source.cacheKey = ResolveBackgroundCacheKey();
        source.bitmap = ResolveCurrentFolderBackground();
        return source;
    };
    ShellTabsListView::AccentColorResolver accentResolver =
        [this](COLORREF* accent, COLORREF* text) {
            return ResolveActiveGroupAccent(accent, text);
        };

    if (!control->Initialize(parent, m_folderView2.Get(), std::move(highlightResolver),
                             std::move(backgroundResolver), std::move(accentResolver),
                             ShouldUseListViewAccentColors())) {
        return false;
    }

    HWND controlWindow = control->GetWindow();
    HWND newListView = control->GetListView();
    if (!controlWindow || !IsWindow(controlWindow) || !newListView || !IsWindow(newListView)) {
        return false;
    }

    const int controlId = GetDlgCtrlID(listView);
    if (controlId != 0) {
        SetWindowLongPtr(controlWindow, GWLP_ID, controlId);
    }

    RECT nativeRect{};
    if (GetWindowRect(listView, &nativeRect)) {
        POINT points[2] = {{nativeRect.left, nativeRect.top}, {nativeRect.right, nativeRect.bottom}};
        MapWindowPoints(nullptr, parent, points, 2);
        const int width = points[1].x - points[0].x;
        const int height = points[1].y - points[0].y;
        SetWindowPos(controlWindow, listView, points[0].x, points[0].y, width, height,
                     SWP_NOACTIVATE | SWP_SHOWWINDOW);
    } else {
        ShowWindow(controlWindow, SW_SHOW);
    }

    EnableWindow(listView, FALSE);
    ShowWindow(listView, SW_HIDE);

    if (!SetWindowSubclass(newListView, &CExplorerBHO::ExplorerViewSubclassProc,
                           reinterpret_cast<UINT_PTR>(this), 0)) {
        EnableWindow(listView, TRUE);
        ShowWindow(listView, SW_SHOW);
        LogLastError(L"SetWindowSubclass(list view)", GetLastError());
        return false;
    }

    m_listView = newListView;
    m_listViewSubclassInstalled = true;
    m_nativeListView = listView;
    m_listViewControlWindow = controlWindow;
    m_listViewControl = std::move(control);

    RegisterGlowSurface(newListView, ExplorerSurfaceKind::ListView, true);
    if (HWND header = ListView_GetHeader(m_listView)) {
        RegisterGlowSurface(header, ExplorerSurfaceKind::Header, true);
    }

    LogMessage(LogLevel::Info,
               L"Installed ShellTabs list view control (native=%p control=%p list=%p)", listView,
               controlWindow, m_listView);

    RefreshListViewControlBackground();
    RefreshListViewAccentState();
    InvalidateRect(controlWindow, nullptr, FALSE);
    return true;
}

bool CExplorerBHO::AttachTreeView(HWND treeView) {
    if (!treeView || !IsWindow(treeView)) {
        if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
            RemoveWindowSubclass(m_treeView, &CExplorerBHO::ExplorerViewSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
        }
        m_treeView = nullptr;
        m_treeViewSubclassInstalled = false;
        m_paneHooks.SetTreeView(nullptr);
        return false;
    }

    if (treeView == m_listView || treeView == m_listViewControlWindow) {
        return false;
    }

    if (m_treeView == treeView && m_treeViewSubclassInstalled) {
        return true;
    }

    if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
        RemoveWindowSubclass(m_treeView, &CExplorerBHO::ExplorerViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }

    if (!SetWindowSubclass(treeView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        LogLastError(L"SetWindowSubclass(tree view)", GetLastError());
        m_treeView = nullptr;
        m_treeViewSubclassInstalled = false;
        m_paneHooks.SetTreeView(nullptr);
        return false;
    }

    m_treeView = treeView;
    m_treeViewSubclassInstalled = true;
    m_paneHooks.SetTreeView(
        m_treeView,
        [this](PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) {
            return ResolveHighlightFromPidl(pidl, highlight);
        },
        m_namespaceTreeControl.Get());
    LogMessage(LogLevel::Info, L"Installed explorer tree view subclass (tree=%p)", treeView);
    return true;
}

void CExplorerBHO::EnsureListViewHostSubclass(HWND hostWindow) {
    if (!hostWindow || !IsWindow(hostWindow)) {
        return;
    }

    if (hostWindow == m_listView || hostWindow == m_listViewControlWindow || hostWindow == m_shellViewWindow ||
        hostWindow == m_directUiView) {
        return;
    }

    if (m_listViewHostSubclassed.find(hostWindow) != m_listViewHostSubclassed.end()) {
        return;
    }

    if (SetWindowSubclass(hostWindow, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_listViewHostSubclassed.insert(hostWindow);
        LogMessage(LogLevel::Info, L"Installed explorer list host subclass (host=%p)", hostWindow);
    } else {
        LogLastError(L"SetWindowSubclass(list host)", GetLastError());
    }
}

void CExplorerBHO::DetachListViewHosts() {
    for (HWND hostWindow : m_listViewHostSubclassed) {
        if (hostWindow && IsWindow(hostWindow)) {
            RemoveWindowSubclass(hostWindow, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
        }
    }
    m_listViewHostSubclassed.clear();
}

bool CExplorerBHO::RegisterGlowSurface(HWND hwnd, ExplorerSurfaceKind kind, bool ensureSubclass) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }
    if (!IsWindowOwnedByThisExplorer(hwnd)) {
        return false;
    }
    if (!m_glowCoordinator.ShouldRenderSurface(kind)) {
        UnregisterGlowSurface(hwnd);
        return false;
    }

    switch (kind) {
        case ExplorerSurfaceKind::Toolbar:
            ConfigureToolbarForCustomSeparators(hwnd);
            break;
        case ExplorerSurfaceKind::Header:
            ConfigureHeaderForCustomDividers(hwnd);
            break;
        default:
            break;
    }

    auto existing = m_glowSurfaces.find(hwnd);
    const bool hadExisting = (existing != m_glowSurfaces.end());
    if (existing != m_glowSurfaces.end()) {
        if (existing->second && existing->second->Kind() == kind && existing->second->IsAttached()) {
            existing->second->RequestRepaint();
            return true;
        }
        if (existing->second) {
            existing->second->Detach();
        }
        m_glowSurfaces.erase(existing);
    }

    bool installedSubclass = false;
    if (ensureSubclass && !hadExisting) {
        if (!SetWindowSubclass(hwnd, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
            const DWORD error = GetLastError();
            std::wstring message = L"SetWindowSubclass(";
            message += DescribeSurfaceKind(kind);
            message += L")";
            LogLastError(message.c_str(), error);
            return false;
        }
        installedSubclass = true;
    }

    auto surface = CreateGlowSurfaceWrapper(kind, m_glowCoordinator);
    if (!surface) {
        if (installedSubclass) {
            RemoveWindowSubclass(hwnd, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
        }
        return false;
    }
    if (!surface->Attach(hwnd)) {
        if (installedSubclass) {
            RemoveWindowSubclass(hwnd, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
        }
        return false;
    }

    surface->RequestRepaint();
    LogMessage(LogLevel::Info, L"Registered glow surface %ls (hwnd=%p)", DescribeSurfaceKind(kind), hwnd);
    m_glowSurfaces.emplace(hwnd, std::move(surface));

    if (kind == ExplorerSurfaceKind::Scrollbar) {
        if (m_glowCoordinator.ShouldRenderSurface(kind)) {
            EnsureScrollbarTransparency(hwnd);
        }

        if (m_scrollbarGlowSubclassed.find(hwnd) == m_scrollbarGlowSubclassed.end()) {
            if (SetWindowSubclass(hwnd, &CExplorerBHO::ScrollbarGlowSubclassProc,
                                  reinterpret_cast<UINT_PTR>(this), 0)) {
                m_scrollbarGlowSubclassed.insert(hwnd);
            } else {
                const DWORD error = GetLastError();
                LogLastError(L"SetWindowSubclass(scrollbar glow)", error);
            }
        }
    }
    return true;
}

void CExplorerBHO::UnregisterGlowSurface(HWND hwnd) {
    if (!hwnd) {
        return;
    }
    auto it = m_glowSurfaces.find(hwnd);
    if (it == m_glowSurfaces.end()) {
        return;
    }

    if (HWND target = it->first; target && IsWindow(target)) {
        RemoveWindowSubclass(target, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
        RemoveWindowSubclass(target, &CExplorerBHO::ScrollbarGlowSubclassProc, reinterpret_cast<UINT_PTR>(this));
        m_scrollbarGlowSubclassed.erase(target);
        RestoreScrollbarTransparency(target);
        InvalidateRect(target, nullptr, FALSE);
    }

    if (it->second) {
        it->second->Detach();
    }

    m_glowSurfaces.erase(it);
}

void CExplorerBHO::RequestHeaderGlowRepaint() const {
    for (const auto& entry : m_glowSurfaces) {
        if (!entry.second) {
            continue;
        }
        if (entry.second->Kind() != ExplorerSurfaceKind::Header) {
            continue;
        }
        entry.second->RequestRepaint();
    }
}

ExplorerGlowSurface* CExplorerBHO::ResolveGlowSurface(HWND hwnd) {
    auto it = m_glowSurfaces.find(hwnd);
    if (it == m_glowSurfaces.end()) {
        return nullptr;
    }
    return it->second.get();
}

const ExplorerGlowSurface* CExplorerBHO::ResolveGlowSurface(HWND hwnd) const {
    auto it = m_glowSurfaces.find(hwnd);
    if (it == m_glowSurfaces.end()) {
        return nullptr;
    }
    return it->second.get();
}

bool CExplorerBHO::ShouldSuppressScrollbarDrawing(HWND hwnd) const {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }
    const ExplorerGlowSurface* surface = ResolveGlowSurface(hwnd);
    if (!surface) {
        return false;
    }
    if (surface->Kind() != ExplorerSurfaceKind::Scrollbar) {
        return false;
    }
    if (!surface->IsAttached()) {
        return false;
    }
    return m_glowCoordinator.ShouldRenderSurface(ExplorerSurfaceKind::Scrollbar);
}

bool CExplorerBHO::PaintScrollbarGlow(HWND hwnd, HDC existingDc, HRGN region) {
    ExplorerGlowSurface* surface = ResolveGlowSurface(hwnd);
    if (!surface || surface->Kind() != ExplorerSurfaceKind::Scrollbar) {
        return false;
    }
    if (!surface->IsAttached()) {
        return false;
    }

    HDC targetDc = existingDc;
    bool releaseDc = false;
    if (!targetDc) {
        UINT flags = DCX_CACHE | DCX_CLIPCHILDREN | DCX_CLIPSIBLINGS | DCX_WINDOW;
        if (region) {
            flags |= DCX_INTERSECTRGN;
        }
        targetDc = GetDCEx(hwnd, region, flags);
        if (!targetDc) {
            return false;
        }
        releaseDc = true;
    }

    RECT clip{};
    if (GetClipBox(targetDc, &clip) == ERROR || IsRectEmpty(&clip)) {
        if (!GetClientRect(hwnd, &clip)) {
            if (releaseDc) {
                ReleaseDC(hwnd, targetDc);
            }
            return false;
        }
    }

    if (clip.right > clip.left && clip.bottom > clip.top) {
        surface->PaintImmediately(targetDc, clip);
    }

    if (releaseDc) {
        ReleaseDC(hwnd, targetDc);
    }

    return true;
}

void CExplorerBHO::EnsureScrollbarTransparency(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return;
    }

    LONG_PTR styles = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
    bool updated = false;
    if (!(styles & WS_EX_TRANSPARENT)) {
        SetWindowLongPtr(hwnd, GWL_EXSTYLE, styles | static_cast<LONG_PTR>(WS_EX_TRANSPARENT));
        SetWindowPos(hwnd, nullptr, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        updated = true;
    }

    const bool inserted = m_transparentScrollbars.insert(hwnd).second;
    if (inserted || updated) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void CExplorerBHO::RestoreScrollbarTransparency(HWND hwnd) {
    const bool wasTracked = (m_transparentScrollbars.erase(hwnd) > 0);
    if (!hwnd || !IsWindow(hwnd)) {
        return;
    }

    LONG_PTR styles = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
    if (styles & WS_EX_TRANSPARENT) {
        SetWindowLongPtr(hwnd, GWL_EXSTYLE, styles & ~static_cast<LONG_PTR>(WS_EX_TRANSPARENT));
        SetWindowPos(hwnd, nullptr, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        InvalidateRect(hwnd, nullptr, FALSE);
        return;
    }

    if (wasTracked) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void CExplorerBHO::RequestScrollbarGlowRepaint(HWND hwnd) {
    ExplorerGlowSurface* surface = ResolveGlowSurface(hwnd);
    if (!surface || surface->Kind() != ExplorerSurfaceKind::Scrollbar) {
        return;
    }
    surface->RequestRepaint();
}

void CExplorerBHO::PruneGlowSurfaces(const std::unordered_set<HWND, HandleHasher>& active) {
    for (auto it = m_glowSurfaces.begin(); it != m_glowSurfaces.end();) {
        HWND target = it->first;
        const bool shouldKeep = target && IsWindow(target) && active.find(target) != active.end();
        if (!shouldKeep) {
            if (target && IsWindow(target)) {
                RemoveWindowSubclass(target, &CExplorerBHO::ExplorerViewSubclassProc,
                                     reinterpret_cast<UINT_PTR>(this));
                RemoveWindowSubclass(target, &CExplorerBHO::ScrollbarGlowSubclassProc,
                                     reinterpret_cast<UINT_PTR>(this));
                m_scrollbarGlowSubclassed.erase(target);
                RestoreScrollbarTransparency(target);
                InvalidateRect(target, nullptr, FALSE);
            }
            if (it->second) {
                it->second->Detach();
            }
            it = m_glowSurfaces.erase(it);
        } else {
            ++it;
        }
    }
}

void CExplorerBHO::ResetGlowSurfaces() {
    for (auto& entry : m_glowSurfaces) {
        HWND target = entry.first;
        if (!target) {
            continue;
        }
        if (IsWindow(target)) {
            RemoveWindowSubclass(target, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
            RemoveWindowSubclass(target, &CExplorerBHO::ScrollbarGlowSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            m_scrollbarGlowSubclassed.erase(target);
            RestoreScrollbarTransparency(target);
            InvalidateRect(target, nullptr, FALSE);
        }
        if (entry.second) {
            entry.second->Detach();
        }
    }
    m_glowSurfaces.clear();
    m_scrollbarGlowSubclassed.clear();
    m_transparentScrollbars.clear();
}

namespace {

bool IsValidStatusBarWindow(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }
    if (!IsWindowVisible(hwnd)) {
        return false;
    }
    RECT rc = {};
    if (!GetClientRect(hwnd, &rc)) {
        return false;
    }
    return rc.right > rc.left && rc.bottom > rc.top;
}

HWND FindVisibleStatusBarDescendant(HWND parent) {
    if (!parent || !IsWindow(parent)) {
        return nullptr;
    }

    for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT)) {
        if (MatchesClass(child, STATUSCLASSNAMEW) && IsValidStatusBarWindow(child)) {
            return child;
        }

        if (HWND found = FindVisibleStatusBarDescendant(child)) {
            return found;
        }
    }

    return nullptr;
}

HWND ResolveStatusBarWindow(IShellBrowser* shellBrowser, HWND frame) {
    HWND statusBar = nullptr;
    if (shellBrowser && SUCCEEDED(shellBrowser->GetControlWindow(FCW_STATUS, &statusBar)) &&
        IsValidStatusBarWindow(statusBar)) {
        return statusBar;
    }

    if (!frame || !IsWindow(frame)) {
        return nullptr;
    }

    return FindVisibleStatusBarDescendant(frame);
}

}  // namespace

void CExplorerBHO::UpdateGlowSurfaceTargets() {
    std::unordered_set<HWND, HandleHasher> active;

    auto registerScrollbarsFor = [&](HWND owner) {
        if (!owner || !IsWindow(owner) || !IsWindowOwnedByThisExplorer(owner)) {
            return;
        }

        struct EnumContext {
            CExplorerBHO* self = nullptr;
            std::unordered_set<HWND, HandleHasher>* active = nullptr;
            HWND parent = nullptr;
        } context{this, &active, owner};

        EnumChildWindows(
            owner,
            [](HWND child, LPARAM param) -> BOOL {
                auto* ctx = reinterpret_cast<EnumContext*>(param);
                if (!ctx || !ctx->self || !ctx->active) {
                    return TRUE;
                }
                if (GetParent(child) != ctx->parent) {
                    return TRUE;
                }
                if (!MatchesClass(child, L"ScrollBar")) {
                    return TRUE;
                }
                if (!ctx->self->IsWindowOwnedByThisExplorer(child)) {
                    return TRUE;
                }
                if (ctx->self->RegisterGlowSurface(child, ExplorerSurfaceKind::Scrollbar, true)) {
                    ctx->active->insert(child);
                }
                return TRUE;
            },
            reinterpret_cast<LPARAM>(&context));
    };

    if (m_listView && IsWindow(m_listView)) {
        if (RegisterGlowSurface(m_listView, ExplorerSurfaceKind::ListView, true)) {
            active.insert(m_listView);
        }

        if (HWND header = ListView_GetHeader(m_listView)) {
            if (RegisterGlowSurface(header, ExplorerSurfaceKind::Header, true)) {
                active.insert(header);
            }
        }
    }

    if (m_directUiView && IsWindow(m_directUiView)) {
        if (RegisterGlowSurface(m_directUiView, ExplorerSurfaceKind::DirectUi, true)) {
            active.insert(m_directUiView);
        }
    }

    registerScrollbarsFor(m_listView);
    registerScrollbarsFor(m_listViewControlWindow);
    registerScrollbarsFor(m_shellViewWindow);
    registerScrollbarsFor(m_directUiView);

    HWND frame = GetTopLevelExplorerWindow();
    HWND statusBarCandidate = ResolveStatusBarWindow(m_shellBrowser.Get(), frame);
    if (statusBarCandidate && !IsWindowOwnedByThisExplorer(statusBarCandidate)) {
        statusBarCandidate = nullptr;
    }

    if (statusBarCandidate != m_statusBar) {
        if (m_statusBar) {
            LogMessage(LogLevel::Info, L"Explorer status bar released (hwnd=%p)", m_statusBar);
            RemoveStatusBarSubclass(m_statusBar);
            ResetStatusBarTheme(m_statusBar);
        }
        m_statusBar = statusBarCandidate;
        m_statusBarThemeValid = false;
        m_statusBarBackgroundColor = CLR_DEFAULT;
        m_statusBarTextColor = CLR_DEFAULT;
        m_statusBarChromeSample.reset();
        if (m_statusBar) {
            LogMessage(LogLevel::Info, L"Explorer status bar discovered (hwnd=%p)", m_statusBar);
            InstallStatusBarSubclass();
        }
    }

    if (frame && IsWindow(frame)) {
        HWND rebar = FindDescendantWindow(frame, L"ReBarWindow32");
        if (rebar && IsWindow(rebar) && IsWindowOwnedByThisExplorer(rebar)) {
            if (RegisterGlowSurface(rebar, ExplorerSurfaceKind::Rebar, true)) {
                active.insert(rebar);
            }

            struct EnumContext {
                CExplorerBHO* self = nullptr;
                std::unordered_set<HWND, HandleHasher>* active = nullptr;
            } context{this, &active};

            EnumChildWindows(
                rebar,
                [](HWND child, LPARAM param) -> BOOL {
                    auto* ctx = reinterpret_cast<EnumContext*>(param);
                    if (!ctx || !ctx->self || !ctx->active) {
                        return TRUE;
                    }
                    if (MatchesClass(child, TOOLBARCLASSNAMEW) && ctx->self->IsWindowOwnedByThisExplorer(child)) {
                        if (HWND parent = GetParent(child); parent && MatchesClass(parent, L"ShellTabsBandWindow")) {
                            return TRUE;
                        }
                        if (ctx->self->RegisterGlowSurface(child, ExplorerSurfaceKind::Toolbar, true)) {
                            ctx->active->insert(child);
                        }
                    }
                    return TRUE;
                },
                reinterpret_cast<LPARAM>(&context));
        }

        for (HWND edit : FindExplorerEditControls()) {
            if (RegisterGlowSurface(edit, ExplorerSurfaceKind::Edit, true)) {
                active.insert(edit);
            }
        }
    }

    if (m_statusBar) {
        UpdateStatusBarTheme();
    }

    PruneGlowSurfaces(active);
}

void CExplorerBHO::ResetStatusBarTheme(HWND statusBar) {
    HWND target = statusBar ? statusBar : m_statusBar;
    if (target && IsWindow(target)) {
        const COLORREF previous =
            static_cast<COLORREF>(SendMessageW(target, SB_SETBKCOLOR, 0, CLR_DEFAULT));
        LogMessage(LogLevel::Info, L"Status bar background reset (hwnd=%p previous=0x%08X)", target, previous);
        InvalidateRect(target, nullptr, TRUE);
    }

    m_statusBarThemeValid = false;
    m_statusBarBackgroundColor = CLR_DEFAULT;
    m_statusBarTextColor = CLR_DEFAULT;
    m_statusBarChromeSample.reset();
}

void CExplorerBHO::InstallStatusBarSubclass() {
    if (!m_statusBar || m_statusBarSubclassInstalled || !IsWindow(m_statusBar)) {
        return;
    }

    if (!SetWindowSubclass(m_statusBar, &CExplorerBHO::StatusBarSubclassProc, reinterpret_cast<UINT_PTR>(this),
            reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(status bar)", GetLastError());
        return;
    }

    m_statusBarSubclassInstalled = true;
}

void CExplorerBHO::RemoveStatusBarSubclass(HWND statusBar) {
    if (!m_statusBarSubclassInstalled) {
        return;
    }

    HWND target = statusBar ? statusBar : m_statusBar;
    if (!target || !IsWindow(target)) {
        m_statusBarSubclassInstalled = false;
        return;
    }

    if (!RemoveWindowSubclass(target, &CExplorerBHO::StatusBarSubclassProc, reinterpret_cast<UINT_PTR>(this))) {
        const DWORD error = GetLastError();
        if (error != ERROR_INVALID_PARAMETER && error != ERROR_SUCCESS) {
            LogLastError(L"RemoveWindowSubclass(status bar)", error);
        }
    }

    m_statusBarSubclassInstalled = false;
}

LRESULT CExplorerBHO::HandleStatusBarMessage(
    HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, bool* handled) {
    if (handled) {
        *handled = false;
    }

    if (msg == WM_NCDESTROY) {
        const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
        RemoveStatusBarSubclass(hwnd);
        if (hwnd == m_statusBar) {
            m_statusBar = nullptr;
            m_statusBarThemeValid = false;
            m_statusBarBackgroundColor = CLR_DEFAULT;
            m_statusBarTextColor = CLR_DEFAULT;
            m_statusBarChromeSample.reset();
        }
        if (handled) {
            *handled = true;
        }
        return result;
    }

    if (!m_statusBarThemeValid || IsSystemHighContrastActive()) {
        return 0;
    }

    auto paintBackground = [&](HDC dc, const RECT& paintRect) {
        if (!dc) {
            return;
        }

        auto fillSolid = [&](HDC targetDc, const RECT& rect, COLORREF color) {
            if (color == CLR_DEFAULT) {
                FillRect(targetDc, &rect, GetSysColorBrush(COLOR_3DFACE));
                return;
            }
            HBRUSH brush = CreateSolidBrush(color);
            if (!brush) {
                FillRect(targetDc, &rect, GetSysColorBrush(COLOR_3DFACE));
                return;
            }
            FillRect(targetDc, &rect, brush);
            DeleteObject(brush);
        };

        COLORREF fallback = m_statusBarBackgroundColor;
        if (fallback == CLR_DEFAULT) {
            fallback = GetSysColor(COLOR_3DFACE);
        }

        COLORREF top = fallback;
        COLORREF bottom = fallback;
        if (m_statusBarChromeSample) {
            top = m_statusBarChromeSample->topColor;
            bottom = m_statusBarChromeSample->bottomColor;
        }
        if (top == CLR_DEFAULT) {
            top = fallback;
        }
        if (bottom == CLR_DEFAULT) {
            bottom = fallback;
        }

        if (top == bottom) {
            fillSolid(dc, paintRect, top);
            return;
        }

        TRIVERTEX vertices[2]{};
        vertices[0].x = paintRect.left;
        vertices[0].y = paintRect.top;
        vertices[0].Red = static_cast<COLOR16>(GetRValue(top) * 0x101);
        vertices[0].Green = static_cast<COLOR16>(GetGValue(top) * 0x101);
        vertices[0].Blue = static_cast<COLOR16>(GetBValue(top) * 0x101);
        vertices[0].Alpha = 0xFFFF;
        vertices[1].x = paintRect.right;
        vertices[1].y = paintRect.bottom;
        vertices[1].Red = static_cast<COLOR16>(GetRValue(bottom) * 0x101);
        vertices[1].Green = static_cast<COLOR16>(GetGValue(bottom) * 0x101);
        vertices[1].Blue = static_cast<COLOR16>(GetBValue(bottom) * 0x101);
        vertices[1].Alpha = 0xFFFF;
        GRADIENT_RECT gradientRect{0, 1};

        if (!GradientFill(dc, vertices, 2, &gradientRect, 1, GRADIENT_FILL_RECT_V)) {
            fillSolid(dc, paintRect, top);
        }
    };

    switch (msg) {
        case WM_ERASEBKGND: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (!dc) {
                break;
            }
            RECT rect{};
            if (!GetClientRect(hwnd, &rect)) {
                break;
            }
            paintBackground(dc, rect);
            if (handled) {
                *handled = true;
            }
            return TRUE;
        }
        case WM_PRINTCLIENT: {
            HDC dc = reinterpret_cast<HDC>(wParam);
            if (dc) {
                RECT rect{};
                if (lParam) {
                    rect = *reinterpret_cast<RECT*>(lParam);
                } else {
                    GetClientRect(hwnd, &rect);
                }
                paintBackground(dc, rect);
            }
            const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            if (handled) {
                *handled = true;
            }
            return result;
        }
        case WM_PAINT: {
            if (wParam) {
                HDC dc = reinterpret_cast<HDC>(wParam);
                RECT rect{};
                if (lParam) {
                    rect = *reinterpret_cast<RECT*>(lParam);
                } else {
                    GetClientRect(hwnd, &rect);
                }
                paintBackground(dc, rect);
                const LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
                if (handled) {
                    *handled = true;
                }
                return result;
            }
            break;
        }
        default:
            break;
    }

    return 0;
}

void CExplorerBHO::UpdateStatusBarTheme() {
    if (!m_statusBar || !IsWindow(m_statusBar)) {
        return;
    }

    if (!IsWindowOwnedByThisExplorer(m_statusBar)) {
        LogMessage(LogLevel::Warning, L"Status bar theme update aborted: handle no longer owned (hwnd=%p)", m_statusBar);
        RemoveStatusBarSubclass(m_statusBar);
        ResetStatusBarTheme(m_statusBar);
        m_statusBar = nullptr;
        return;
    }

    InstallStatusBarSubclass();

    if (IsSystemHighContrastActive()) {
        if (m_statusBarThemeValid) {
            LogMessage(LogLevel::Info, L"Status bar theme disabled for high contrast (hwnd=%p)", m_statusBar);
        }
        ResetStatusBarTheme(m_statusBar);
        return;
    }

    HWND frame = GetTopLevelExplorerWindow();
    HWND rebar = nullptr;
    if (frame && IsWindow(frame)) {
        rebar = FindDescendantWindow(frame, L"ReBarWindow32");
        if (rebar && !IsWindowOwnedByThisExplorer(rebar)) {
            rebar = nullptr;
        }
    }

    std::optional<ToolbarChromeSample> chrome;
    if (rebar && IsWindow(rebar)) {
        chrome = SampleToolbarChrome(rebar);
    }
    if (!chrome && frame && IsWindow(frame)) {
        chrome = SampleToolbarChrome(frame);
    }

    if (!chrome) {
        HWND parent = GetParent(m_statusBar);
        if (parent && IsWindow(parent) && IsWindowOwnedByThisExplorer(parent)) {
            chrome = SampleToolbarChrome(parent);
        }
    }

    std::optional<COLORREF> backgroundCandidate;
    COLORREF gradientTop = CLR_DEFAULT;
    COLORREF gradientBottom = CLR_DEFAULT;

    if (chrome) {
        auto averageColor = [](COLORREF first, COLORREF second) -> COLORREF {
            const int red = (static_cast<int>(GetRValue(first)) + static_cast<int>(GetRValue(second))) / 2;
            const int green = (static_cast<int>(GetGValue(first)) + static_cast<int>(GetGValue(second))) / 2;
            const int blue = (static_cast<int>(GetBValue(first)) + static_cast<int>(GetBValue(second))) / 2;
            return RGB(red, green, blue);
        };

        const COLORREF background = averageColor(chrome->topColor, chrome->bottomColor);
        backgroundCandidate = background;
        gradientTop = chrome->topColor;
        gradientBottom = chrome->bottomColor;
    } else if (IsAppDarkModePreferred()) {
        LogMessage(LogLevel::Info, L"Status bar theme fallback to dark preference (hwnd=%p)", m_statusBar);
        backgroundCandidate = RGB(32, 32, 32);
        gradientTop = backgroundCandidate.value();
        gradientBottom = backgroundCandidate.value();
    }

    if (!backgroundCandidate.has_value()) {
        if (m_statusBarThemeValid) {
            LogMessage(LogLevel::Warning, L"Status bar theme reset: failed to sample toolbar chrome (hwnd=%p)", m_statusBar);
        }
        ResetStatusBarTheme(m_statusBar);
        return;
    }

    const COLORREF background = backgroundCandidate.value();

    auto resolveGradientColor = [&](COLORREF color) -> COLORREF {
        return color == CLR_DEFAULT ? background : color;
    };

    const COLORREF resolvedTop = resolveGradientColor(gradientTop);
    const COLORREF resolvedBottom = resolveGradientColor(gradientBottom);
    const COLORREF text = ChooseStatusBarTextColor(resolvedTop, resolvedBottom);

    const bool backgroundChanged = !m_statusBarThemeValid || background != m_statusBarBackgroundColor;
    const bool textChanged = !m_statusBarThemeValid || text != m_statusBarTextColor;

    bool chromeChanged = false;
    ToolbarChromeSample chromeForStorage{background, background};
    if (chrome) {
        chromeForStorage = *chrome;
    }
    if (!m_statusBarChromeSample.has_value()) {
        chromeChanged = true;
    } else {
        chromeChanged = m_statusBarChromeSample->topColor != chromeForStorage.topColor ||
                        m_statusBarChromeSample->bottomColor != chromeForStorage.bottomColor;
    }

    if (!backgroundChanged && !textChanged && !chromeChanged) {
        return;
    }

    if (backgroundChanged) {
        LogMessage(LogLevel::Info, L"Status bar theme background updated (hwnd=%p new=0x%08X)", m_statusBar, background);
    }

    if (textChanged) {
        LogMessage(LogLevel::Info, L"Status bar theme text color updated (hwnd=%p new=0x%08X)", m_statusBar, text);
    }

    m_statusBarThemeValid = true;
    m_statusBarBackgroundColor = background;
    m_statusBarTextColor = text;
    m_statusBarChromeSample = chromeForStorage;

    InvalidateRect(m_statusBar, nullptr, TRUE);
}

void CExplorerBHO::HandleExplorerPostPaint(HWND hwnd, UINT msg, WPARAM wParam) {
    auto it = m_glowSurfaces.find(hwnd);
    if (it == m_glowSurfaces.end() || !it->second) {
        return;
    }

    ExplorerGlowSurface* surface = it->second.get();
    if (!surface->SupportsImmediatePainting()) {
        return;
    }

    HDC targetDc = nullptr;
    bool releaseDc = false;
    if (msg == WM_PAINT) {
        if (wParam) {
            targetDc = reinterpret_cast<HDC>(wParam);
        } else {
            targetDc = GetDC(hwnd);
            releaseDc = (targetDc != nullptr);
        }
    } else if (msg == WM_PRINTCLIENT) {
        targetDc = reinterpret_cast<HDC>(wParam);
    }

    if (!targetDc) {
        if (releaseDc) {
            ReleaseDC(hwnd, targetDc);
        }
        return;
    }

    RECT clipRect{0, 0, 0, 0};
    bool hasClip = false;

    if (GetClipBox(targetDc, &clipRect) != ERROR && !IsRectEmpty(&clipRect)) {
        hasClip = true;
    }

    if (!hasClip && msg == WM_PAINT && wParam == 0) {
        RECT update{};
        if (GetUpdateRect(hwnd, &update, FALSE) && !IsRectEmpty(&update)) {
            clipRect = update;
            hasClip = true;
        }
    }

    if (!hasClip) {
        if (!GetClientRect(hwnd, &clipRect) || IsRectEmpty(&clipRect)) {
            if (releaseDc) {
                ReleaseDC(hwnd, targetDc);
            }
            return;
        }
    }

    surface->PaintImmediately(targetDc, clipRect);

    if (releaseDc) {
        ReleaseDC(hwnd, targetDc);
    }
}

bool CExplorerBHO::TryAttachListViewFromFolderView() {
    HWND listView = ResolveListViewFromFolderView();
    if (!listView) {
        return false;
    }

    if (!AttachListView(listView)) {
        return false;
    }

    if (HWND parent = GetParent(listView)) {
        EnsureListViewHostSubclass(parent);
    }

    RefreshListViewAccentState();
    return true;
}

HWND CExplorerBHO::ResolveListViewFromFolderView() {
    if (!m_folderView2 && m_shellView) {
        Microsoft::WRL::ComPtr<IFolderView2> folderView;
        const HRESULT hr = m_shellView->QueryInterface(IID_PPV_ARGS(&folderView));
        if (SUCCEEDED(hr) && folderView) {
            m_folderView2 = std::move(folderView);
        }
    }

    if (!m_folderView2) {
        return nullptr;
    }

    Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
    HRESULT hr = m_folderView2.As(&oleWindow);
    if (FAILED(hr) || !oleWindow) {
        m_folderView2.Reset();
        return nullptr;
    }

    HWND listView = nullptr;
    hr = oleWindow->GetWindow(&listView);
    if (FAILED(hr) || !listView || !IsWindow(listView) || !IsWindowOwnedByThisExplorer(listView)) {
        m_folderView2.Reset();
        return nullptr;
    }

    return listView;
}

void CExplorerBHO::EnsureListViewSubclass() {
    if (m_listView && m_listViewSubclassInstalled && IsWindow(m_listView)) {
        return;
    }

    if (m_listView && !IsWindow(m_listView)) {
        DetachListView();
    }

    if (TryAttachListViewFromFolderView()) {
        return;
    }

    const HWND baseScopes[] = {m_directUiView, m_shellViewWindow, m_frameWindow};
    std::vector<HWND> hostCandidates;
    std::unordered_set<HWND, HandleHasher> visited;

    auto addCandidate = [&](HWND hwnd) {
        if (!hwnd || !IsWindow(hwnd)) {
            return;
        }
        if (visited.insert(hwnd).second) {
            hostCandidates.push_back(hwnd);
        }
    };

    for (HWND scope : baseScopes) {
        addCandidate(scope);
    }

    constexpr const wchar_t* kHostClasses[] = {
        L"UIItemsView",
        L"ItemsViewWnd",
        L"DirectUIHWND",
        L"DUIViewWndClassName",
        L"ShellTabWindowClass",
    };

    for (HWND scope : baseScopes) {
        if (!scope || !IsWindow(scope)) {
            continue;
        }

        for (const wchar_t* className : kHostClasses) {
            for (HWND ancestor = scope; ancestor && IsWindow(ancestor); ancestor = GetParent(ancestor)) {
                if (MatchesClass(ancestor, className)) {
                    addCandidate(ancestor);
                }
            }

            HWND descendant = FindDescendantWindow(scope, className);
            if (descendant) {
                addCandidate(descendant);
            }
        }
    }

    for (HWND candidate : hostCandidates) {
        if (!candidate || !IsWindow(candidate)) {
            continue;
        }

        EnsureListViewHostSubclass(candidate);

        HWND listView = nullptr;
        if (MatchesClass(candidate, L"SysListView32")) {
            listView = candidate;
        } else {
            listView = FindDescendantWindow(candidate, L"SysListView32");
        }

        if (listView && AttachListView(listView)) {
            RefreshListViewAccentState();
            return;
        }
    }
}

void CExplorerBHO::UpdateExplorerViewSubclass() {
    RemoveExplorerViewSubclass();

    if (!m_shellBrowser) {
        return;
    }

    Microsoft::WRL::ComPtr<IShellView> shellView;
    HRESULT hr = m_shellBrowser->QueryActiveShellView(&shellView);
    if (FAILED(hr) || !shellView) {
        return;
    }

    HWND viewWindow = nullptr;
    hr = shellView->GetWindow(&viewWindow);
    if (FAILED(hr) || !viewWindow) {
        return;
    }

    if (!InstallExplorerViewSubclass(viewWindow)) {
        LogMessage(LogLevel::Warning, L"Explorer view subclass installation failed (view=%p)", viewWindow);
        return;
    }

    m_shellView = shellView;
    m_folderView2.Reset();
    if (shellView) {
        Microsoft::WRL::ComPtr<IFolderView2> folderView;
        if (SUCCEEDED(shellView->QueryInterface(IID_PPV_ARGS(&folderView))) && folderView) {
            m_folderView2 = std::move(folderView);
        }
    }
    m_shellViewWindow = viewWindow;
    UpdateCurrentFolderBackground();

    if (!TryResolveExplorerPanes()) {
        ScheduleExplorerPaneRetry();
    }
}

bool CExplorerBHO::InstallExplorerViewSubclass(HWND viewWindow) {
    bool installed = false;

    if (viewWindow && IsWindow(viewWindow)) {
        if (SetWindowSubclass(viewWindow, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this),
                              0)) {
            m_shellViewWindowSubclassInstalled = true;
            installed = true;
            LogMessage(LogLevel::Info, L"Installed shell view window subclass (view=%p)", viewWindow);
        } else {
            LogLastError(L"SetWindowSubclass(shell view window)", GetLastError());
            m_shellViewWindowSubclassInstalled = false;
        }
    } else {
        m_shellViewWindowSubclassInstalled = false;
    }

    HWND frameWindow = GetTopLevelExplorerWindow();
    if (frameWindow && frameWindow != viewWindow && IsWindow(frameWindow)) {
        if (SetWindowSubclass(frameWindow, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this),
                              0)) {
            m_frameWindow = frameWindow;
            m_frameSubclassInstalled = true;
            installed = true;
            LogMessage(LogLevel::Info, L"Installed explorer frame subclass (frame=%p)", frameWindow);
        } else {
            LogLastError(L"SetWindowSubclass(explorer frame)", GetLastError());
            m_frameSubclassInstalled = false;
            m_frameWindow = nullptr;
        }
    } else {
        m_frameSubclassInstalled = false;
        m_frameWindow = nullptr;
    }

    if (installed) {
        ClearPendingOpenInNewTabState();
        LogMessage(LogLevel::Info, L"Explorer view base subclass ready (view=%p frame=%p)", viewWindow, m_frameWindow);
    } else {
        LogMessage(LogLevel::Warning,
                   L"Explorer view subclass installation skipped: no valid targets (view=%p frame=%p)", viewWindow,
                   frameWindow);
    }

    return installed;
}

bool CExplorerBHO::TryResolveExplorerPanes() {
    if (!m_shellViewWindow || !IsWindow(m_shellViewWindow)) {
        return false;
    }

    if (m_directUiView && (!IsWindow(m_directUiView) || !m_directUiSubclassInstalled)) {
        if (m_directUiView && m_directUiSubclassInstalled && IsWindow(m_directUiView)) {
            RemoveWindowSubclass(m_directUiView, &CExplorerBHO::ExplorerViewSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
        }
        UnregisterGlowSurface(m_directUiView);
        m_directUiView = nullptr;
        m_directUiSubclassInstalled = false;
    }

    if (m_listView && (!IsWindow(m_listView) || !m_listViewSubclassInstalled)) {
        DetachListView();
    }

    if (m_treeView && (!IsWindow(m_treeView) || !m_treeViewSubclassInstalled)) {
        if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
            RemoveWindowSubclass(m_treeView, &CExplorerBHO::ExplorerViewSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
        }
        m_treeView = nullptr;
        m_treeViewSubclassInstalled = false;
    }

    bool listViewResolved = (m_listView && m_listViewSubclassInstalled && IsWindow(m_listView));
    bool treeViewResolved = (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView));

    if (!m_directUiSubclassInstalled) {
        HWND directUiHost = FindDescendantWindow(m_shellViewWindow, L"UIItemsView");
        if (!directUiHost) {
            directUiHost = FindDescendantWindow(m_shellViewWindow, L"ItemsViewWnd");
        }
        if (!directUiHost) {
            directUiHost = FindDescendantWindow(m_shellViewWindow, L"DirectUIHWND");
        }

        if (directUiHost && directUiHost != m_shellViewWindow && directUiHost != m_listView &&
            directUiHost != m_listViewControlWindow &&
            IsWindow(directUiHost)) {
            if (SetWindowSubclass(directUiHost, &CExplorerBHO::ExplorerViewSubclassProc,
                                  reinterpret_cast<UINT_PTR>(this), 0)) {
                m_directUiView = directUiHost;
                m_directUiSubclassInstalled = true;
                RegisterGlowSurface(directUiHost, ExplorerSurfaceKind::DirectUi, false);
                LogMessage(LogLevel::Info, L"Installed explorer DirectUI host subclass (direct=%p)", directUiHost);
            } else {
                LogLastError(L"SetWindowSubclass(DirectUI host)", GetLastError());
                m_directUiView = nullptr;
                m_directUiSubclassInstalled = false;
            }
        } else if (!directUiHost) {
            m_directUiView = nullptr;
        }
    }

    if (!listViewResolved && TryAttachListViewFromFolderView()) {
        listViewResolved = true;
    }

    if (!listViewResolved) {
        const HWND candidates[] = {m_directUiView, m_shellViewWindow};
        for (HWND candidate : candidates) {
            if (!candidate || !IsWindow(candidate)) {
                continue;
            }
            HWND listView = FindDescendantWindow(candidate, L"SysListView32");
            if (listView && AttachListView(listView)) {
                listViewResolved = true;
                RefreshListViewAccentState();
                break;
            }
        }
    }

    if (!treeViewResolved) {
        HWND treeView = nullptr;
        if (m_shellBrowser) {
            HWND browserTree = nullptr;
            if (SUCCEEDED(m_shellBrowser->GetControlWindow(FCW_TREE, &browserTree)) && browserTree &&
                browserTree != m_listView && browserTree != m_listViewControlWindow && IsWindow(browserTree)) {
                treeView = browserTree;
            }
        }
        if (!treeView) {
            treeView = FindDescendantWindow(m_shellViewWindow, L"SysTreeView32");
        }
        if (treeView && AttachTreeView(treeView)) {
            treeViewResolved = true;
        }
    }

    if (m_treeViewSubclassInstalled && m_treeView) {
        m_paneHooks.SetTreeView(
            m_treeView,
            [this](PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) {
                return ResolveHighlightFromPidl(pidl, highlight);
            },
            m_namespaceTreeControl.Get());
    } else {
        m_paneHooks.SetTreeView(nullptr);
    }

    UpdateGlowSurfaceTargets();

    UpdateExplorerPaneCreationWatch(!listViewResolved, !treeViewResolved);

    if (listViewResolved && treeViewResolved) {
        CancelExplorerPaneRetry();
        CancelExplorerPaneFallback();
        if (!m_loggedExplorerPanesReady) {
            LogMessage(LogLevel::Info, L"Explorer panes resolved (view=%p list=%p tree=%p direct=%p)", m_shellViewWindow,
                       m_listView, m_treeView, m_directUiView);
            UpdateCurrentFolderBackground();
            m_loggedExplorerPanesReady = true;
        }
        m_loggedListViewMissing = false;
        m_loggedTreeViewMissing = false;
        return true;
    }

    m_loggedExplorerPanesReady = false;

    if (m_watchListViewCreation || m_watchTreeViewCreation) {
        ScheduleExplorerPaneFallback();
    } else {
        CancelExplorerPaneFallback();
    }

    if (!listViewResolved) {
        if (!m_loggedListViewMissing) {
            LogMessage(LogLevel::Info, L"Explorer panes not ready: list view missing (view=%p)", m_shellViewWindow);
            m_loggedListViewMissing = true;
        }
    } else {
        m_loggedListViewMissing = false;
    }

    if (!treeViewResolved) {
        if (!m_loggedTreeViewMissing) {
            LogMessage(LogLevel::Info, L"Explorer panes not ready: tree view missing (view=%p)", m_shellViewWindow);
            m_loggedTreeViewMissing = true;
        }
    } else {
        m_loggedTreeViewMissing = false;
    }

    ScheduleExplorerPaneRetry();
    return false;
}

void CExplorerBHO::HandleExplorerPaneCandidate(HWND candidate) {
    if (!candidate || !IsWindow(candidate)) {
        return;
    }

    if (!m_watchListViewCreation && !m_watchTreeViewCreation) {
        return;
    }

    wchar_t className[64] = {};
    if (GetClassNameW(candidate, className, static_cast<int>(_countof(className))) == 0) {
        return;
    }

    if (m_watchListViewCreation && _wcsicmp(className, L"SysListView32") == 0) {
        LogMessage(LogLevel::Info,
                   L"Explorer pane creation event detected: list view (child=%p parent=%p)", candidate, m_shellViewWindow);
        if (AttachListView(candidate)) {
            RefreshListViewAccentState();
        }
    } else if (m_watchTreeViewCreation && _wcsicmp(className, L"SysTreeView32") == 0) {
        LogMessage(LogLevel::Info,
                   L"Explorer pane creation event detected: tree view (child=%p parent=%p)", candidate, m_shellViewWindow);
        AttachTreeView(candidate);
    }
}

void CExplorerBHO::UpdateExplorerPaneCreationWatch(bool watchListView, bool watchTreeView) {
    const bool previousListWatch = m_watchListViewCreation;
    const bool previousTreeWatch = m_watchTreeViewCreation;

    m_watchListViewCreation = watchListView;
    m_watchTreeViewCreation = watchTreeView;

    if (previousListWatch != watchListView || previousTreeWatch != watchTreeView) {
        if (watchListView || watchTreeView) {
            LogMessage(LogLevel::Info,
                       L"Explorer pane creation watch armed (view=%p list=%d tree=%d)", m_shellViewWindow, watchListView,
                       watchTreeView);
        } else {
            LogMessage(LogLevel::Info, L"Explorer pane creation watch cleared (view=%p)", m_shellViewWindow);
            m_explorerPaneFallbackUsed = false;
        }
    }
}

void CExplorerBHO::ScheduleExplorerPaneRetry() {
    if (m_explorerPaneRetryPending) {
        return;
    }
    if (!m_shellViewWindow || !IsWindow(m_shellViewWindow)) {
        return;
    }

    DWORD nextDelay = (m_explorerPaneRetryDelayMs == 0) ? kEnsureRetryInitialDelayMs : m_explorerPaneRetryDelayMs * 2;
    if (nextDelay > kEnsureRetryMaxDelayMs) {
        nextDelay = kEnsureRetryMaxDelayMs;
    }

    UINT_PTR timerId = SetTimer(m_shellViewWindow, 0, nextDelay, nullptr);
    if (timerId == 0) {
        LogLastError(L"SetTimer(explorer pane retry)", GetLastError());
        return;
    }

    m_explorerPaneRetryPending = true;
    m_explorerPaneRetryTimerId = timerId;
    m_explorerPaneRetryDelayMs = nextDelay;
    ++m_explorerPaneRetryAttempts;

    LogMessage(LogLevel::Info,
               L"Explorer pane retry timer armed (view=%p delay=%u attempts=%zu)",
               m_shellViewWindow, nextDelay, m_explorerPaneRetryAttempts);
}

void CExplorerBHO::CancelExplorerPaneRetry(bool resetAttemptState) {
    if (m_explorerPaneRetryPending && m_shellViewWindow && IsWindow(m_shellViewWindow) &&
        m_explorerPaneRetryTimerId != 0) {
        if (!KillTimer(m_shellViewWindow, m_explorerPaneRetryTimerId)) {
            const DWORD error = GetLastError();
            if (error != 0) {
                LogLastError(L"KillTimer(explorer pane retry)", error);
            }
        }
    }

    m_explorerPaneRetryPending = false;
    m_explorerPaneRetryTimerId = 0;
    if (resetAttemptState) {
        m_explorerPaneRetryDelayMs = 0;
        m_explorerPaneRetryAttempts = 0;
    }
}

void CExplorerBHO::ScheduleExplorerPaneFallback() {
    if (m_explorerPaneFallbackPending || m_explorerPaneFallbackUsed) {
        return;
    }
    if (!m_shellViewWindow || !IsWindow(m_shellViewWindow)) {
        return;
    }

    UINT_PTR timerId = SetTimer(m_shellViewWindow, 0, kEnsureRetryInitialDelayMs, nullptr);
    if (timerId != 0) {
        m_explorerPaneFallbackPending = true;
        m_explorerPaneFallbackTimerId = timerId;
        m_explorerPaneFallbackUsed = true;
        LogMessage(LogLevel::Info, L"Explorer pane fallback timer armed (view=%p delay=%u)", m_shellViewWindow,
                   kEnsureRetryInitialDelayMs);
    } else {
        LogLastError(L"SetTimer(explorer pane fallback)", GetLastError());
    }
}

void CExplorerBHO::CancelExplorerPaneFallback() {
    if (m_explorerPaneFallbackPending && m_shellViewWindow && IsWindow(m_shellViewWindow) &&
        m_explorerPaneFallbackTimerId != 0) {
        if (!KillTimer(m_shellViewWindow, m_explorerPaneFallbackTimerId)) {
            const DWORD error = GetLastError();
            if (error != 0) {
                LogLastError(L"KillTimer(explorer pane fallback)", error);
            }
        }
    }
    m_explorerPaneFallbackPending = false;
    m_explorerPaneFallbackTimerId = 0;
}

void CExplorerBHO::RemoveExplorerViewSubclass() {
    CancelExplorerPaneFallback();
    CancelExplorerPaneRetry();
    UpdateExplorerPaneCreationWatch(false, false);
    ResetNamespaceTreeControl();

    if (m_shellViewWindow && m_shellViewWindowSubclassInstalled && IsWindow(m_shellViewWindow)) {
        RemoveWindowSubclass(m_shellViewWindow, &CExplorerBHO::ExplorerViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }
    if (m_frameWindow && m_frameSubclassInstalled && IsWindow(m_frameWindow)) {
        RemoveWindowSubclass(m_frameWindow, &CExplorerBHO::ExplorerViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }
    DetachListView();
    DetachListViewHosts();
    if (m_directUiView && m_directUiSubclassInstalled && IsWindow(m_directUiView)) {
        RemoveWindowSubclass(m_directUiView, &CExplorerBHO::ExplorerViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }
    UnregisterGlowSurface(m_directUiView);
    if (m_treeView && m_treeViewSubclassInstalled && IsWindow(m_treeView)) {
        RemoveWindowSubclass(m_treeView, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(this));
    }

    ResetGlowSurfaces();

    if (m_statusBar) {
        RemoveStatusBarSubclass();
        ResetStatusBarTheme();
        m_statusBar = nullptr;
    }

    m_shellViewWindowSubclassInstalled = false;
    m_frameWindow = nullptr;
    m_frameSubclassInstalled = false;
    m_directUiView = nullptr;
    m_directUiSubclassInstalled = false;
    m_treeView = nullptr;
    m_treeViewSubclassInstalled = false;
    m_loggedExplorerPanesReady = false;
    m_loggedListViewMissing = false;
    m_loggedTreeViewMissing = false;
    m_paneHooks.Reset();
    m_shellViewWindow = nullptr;
    m_folderView2.Reset();
    m_shellView.Reset();
    ClearPendingOpenInNewTabState();
}

void CExplorerBHO::TryAttachNamespaceTreeControl(IShellView* shellView) {
    if (!shellView) {
        return;
    }

    ResetNamespaceTreeControl();

    Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
    HRESULT hr = shellView->QueryInterface(IID_PPV_ARGS(&serviceProvider));
    if (FAILED(hr) || !serviceProvider) {
        return;
    }

    Microsoft::WRL::ComPtr<INameSpaceTreeControl> treeControl;
    hr = serviceProvider->QueryService(SID_NamespaceTreeControl, IID_PPV_ARGS(&treeControl));
    if (FAILED(hr) || !treeControl) {
        return;
    }

    m_namespaceTreeControl = treeControl;

    auto resolver = [this](PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) {
        return ResolveHighlightFromPidl(pidl, highlight);
    };

    auto host = std::make_unique<NamespaceTreeHost>(treeControl, resolver);
    if (!host || !host->Initialize()) {
        LogMessage(LogLevel::Warning, L"Namespace tree host initialization failed");
        m_namespaceTreeHost.reset();
        return;
    }

    m_namespaceTreeHost = std::move(host);

    if (m_treeView && IsWindow(m_treeView)) {
        m_paneHooks.SetTreeView(
            m_treeView,
            [this](PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) {
                return ResolveHighlightFromPidl(pidl, highlight);
            },
            m_namespaceTreeControl.Get());
    }

    InvalidateNamespaceTreeControl();
}

void CExplorerBHO::ResetNamespaceTreeControl() {
    m_namespaceTreeHost.reset();
    m_namespaceTreeControl.Reset();

    if (m_treeView && IsWindow(m_treeView)) {
        m_paneHooks.SetTreeView(
            m_treeView,
            [this](PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) {
                return ResolveHighlightFromPidl(pidl, highlight);
            },
            nullptr);
    }
}

void CExplorerBHO::InvalidateNamespaceTreeControl() const {
    if (m_namespaceTreeHost) {
        m_namespaceTreeHost->InvalidateAll();
        return;
    }

    if (!m_namespaceTreeControl) {
        return;
    }

    Microsoft::WRL::ComPtr<INameSpaceTreeControl> control(m_namespaceTreeControl);
    Microsoft::WRL::ComPtr<IOleWindow> oleWindow;
    if (FAILED(control.As(&oleWindow)) || !oleWindow) {
        return;
    }

    HWND hwnd = nullptr;
    if (SUCCEEDED(oleWindow->GetWindow(&hwnd)) && hwnd) {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

void CExplorerBHO::ClearFolderBackgrounds() {
    m_folderBackgroundEntries.clear();
    m_folderBackgroundBitmaps.clear();
    m_universalBackgroundImagePath.clear();
    m_universalBackgroundBitmap.reset();
    m_failedBackgroundKeys.clear();
    m_folderBackgroundsEnabled = false;
    RefreshListViewControlBackground();
}

std::wstring CExplorerBHO::NormalizeBackgroundKey(const std::wstring& path) const {
    std::wstring normalized = NormalizeFileSystemPath(path);
    if (normalized.empty()) {
        return {};
    }

    std::transform(normalized.begin(), normalized.end(), normalized.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });
    return normalized;
}

void CExplorerBHO::ReloadFolderBackgrounds(const ShellTabsOptions& options) {
    ClearFolderBackgrounds();

    if (!m_gdiplusInitialized) {
        return;
    }

    if (!options.enableFolderBackgrounds) {
        InvalidateFolderBackgroundTargets();
        return;
    }

    m_folderBackgroundsEnabled = true;

    if (!options.universalFolderBackgroundImage.cachedImagePath.empty()) {
        m_universalBackgroundImagePath = options.universalFolderBackgroundImage.cachedImagePath;
        m_universalBackgroundBitmap.reset();
    }

    for (const auto& entry : options.folderBackgroundEntries) {
        if (entry.folderPath.empty() || entry.image.cachedImagePath.empty()) {
            continue;
        }

        std::wstring key = NormalizeBackgroundKey(entry.folderPath);
        if (key.empty()) {
            continue;
        }

        FolderBackgroundEntryData data{};
        data.imagePath = entry.image.cachedImagePath;
        data.folderDisplayPath = entry.folderPath;

        auto [it, inserted] = m_folderBackgroundEntries.try_emplace(std::move(key), std::move(data));
        if (!inserted) {
            it->second.imagePath = entry.image.cachedImagePath;
            it->second.folderDisplayPath = entry.folderPath;
        }
    }

    InvalidateFolderBackgroundTargets();
    RefreshListViewControlBackground();
}

bool CExplorerBHO::EnsureFolderBackgroundBitmap(const std::wstring& key) const {
    if (key.empty() || !m_gdiplusInitialized || !m_folderBackgroundsEnabled) {
        return false;
    }

    auto existing = m_folderBackgroundBitmaps.find(key);
    if (existing != m_folderBackgroundBitmaps.end()) {
        return existing->second != nullptr;
    }

    if (m_failedBackgroundKeys.find(key) != m_failedBackgroundKeys.end()) {
        return false;
    }

    auto entryIt = m_folderBackgroundEntries.find(key);
    if (entryIt == m_folderBackgroundEntries.end() || entryIt->second.imagePath.empty()) {
        return false;
    }

    auto bitmap = LoadBackgroundBitmap(entryIt->second.imagePath);
    if (!bitmap) {
        m_failedBackgroundKeys.insert(key);
        LogMessage(LogLevel::Warning, L"Failed to load background for %ls from %ls",
                   entryIt->second.folderDisplayPath.c_str(), entryIt->second.imagePath.c_str());
        return false;
    }

    m_folderBackgroundBitmaps.emplace(key, std::move(bitmap));
    return true;
}

bool CExplorerBHO::EnsureUniversalBackgroundBitmap() const {
    if (!m_folderBackgroundsEnabled || !m_gdiplusInitialized) {
        return false;
    }

    if (m_universalBackgroundBitmap) {
        return true;
    }

    if (m_universalBackgroundImagePath.empty()) {
        return false;
    }

    auto bitmap = LoadBackgroundBitmap(m_universalBackgroundImagePath);
    if (!bitmap) {
        LogMessage(LogLevel::Warning, L"Failed to load universal folder background from %ls",
                   m_universalBackgroundImagePath.c_str());
        m_universalBackgroundImagePath.clear();
        return false;
    }

    m_universalBackgroundBitmap = std::move(bitmap);
    return true;
}

Gdiplus::Bitmap* CExplorerBHO::ResolveCurrentFolderBackground() const {
    if (!m_folderBackgroundsEnabled || !m_gdiplusInitialized) {
        return nullptr;
    }

    if (EnsureFolderBackgroundBitmap(m_currentFolderKey)) {
        auto it = m_folderBackgroundBitmaps.find(m_currentFolderKey);
        if (it != m_folderBackgroundBitmaps.end()) {
            return it->second.get();
        }
    }

    if (EnsureUniversalBackgroundBitmap()) {
        return m_universalBackgroundBitmap.get();
    }

    return nullptr;
}

std::wstring CExplorerBHO::ResolveBackgroundCacheKey() const {
    if (!m_currentFolderKey.empty()) {
        return m_currentFolderKey;
    }

    if (m_universalBackgroundBitmap) {
        return std::wstring(kUniversalBackgroundCacheKey);
    }

    return {};
}

void CExplorerBHO::RefreshListViewControlBackground() {
    if (!m_listViewControl) {
        return;
    }

    auto resolver = [this]() -> ShellTabsListView::BackgroundSource {
        ShellTabsListView::BackgroundSource source{};
        source.cacheKey = ResolveBackgroundCacheKey();
        source.bitmap = ResolveCurrentFolderBackground();
        return source;
    };
    m_listViewControl->SetBackgroundResolver(resolver);
}

void CExplorerBHO::UpdateCurrentFolderBackground() {
    if (!m_folderBackgroundsEnabled) {
        if (!m_currentFolderKey.empty()) {
            m_currentFolderKey.clear();
            InvalidateFolderBackgroundTargets();
        }
        return;
    }

    bool resolvedKey = false;
    std::wstring newKey;

    if (m_shellBrowser) {
        UniquePidl current = GetCurrentFolderPidL(m_shellBrowser, m_webBrowser);
        if (current) {
            resolvedKey = true;
            PWSTR path = nullptr;
            if (SUCCEEDED(SHGetNameFromIDList(current.get(), SIGDN_FILESYSPATH, &path)) && path && path[0] != L'\0') {
                newKey = NormalizeBackgroundKey(path);
            }
            if (path) {
                CoTaskMemFree(path);
            }
        }
    }

    if (!resolvedKey) {
        return;
    }

    if (newKey == m_currentFolderKey) {
        if (!EnsureFolderBackgroundBitmap(m_currentFolderKey)) {
            EnsureUniversalBackgroundBitmap();
        }
        return;
    }

    m_currentFolderKey = std::move(newKey);

    if (!EnsureFolderBackgroundBitmap(m_currentFolderKey)) {
        EnsureUniversalBackgroundBitmap();
    }

    InvalidateFolderBackgroundTargets();
    RefreshListViewControlBackground();
}

void CExplorerBHO::InvalidateFolderBackgroundTargets() const {
    auto requestRedraw = [](HWND hwnd) {
        if (!hwnd || !IsWindow(hwnd)) {
            return;
        }
        RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE | RDW_INTERNALPAINT);
    };

    requestRedraw(m_listView);
    requestRedraw(m_listViewControlWindow);
    requestRedraw(m_directUiView);
    for (HWND host : m_listViewHostSubclassed) {
        requestRedraw(host);
    }
    requestRedraw(m_shellViewWindow);
    requestRedraw(m_frameWindow);
}

bool CExplorerBHO::ShouldUseListViewAccentColors() const {
    if (!m_useExplorerAccentColors) {
        return false;
    }
    if (!m_listView || !IsWindow(m_listView)) {
        return false;
    }
    if (IsSystemHighContrastActive()) {
        return false;
    }
    return true;
}

bool CExplorerBHO::ResolveActiveGroupAccent(COLORREF* accent, COLORREF* text) const {
    if (!accent || !text) {
        return false;
    }
    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        return false;
    }

    TabManager::ExplorerWindowId id;
    id.hwnd = frame;
    if (m_webBrowser) {
        id.frameCookie = reinterpret_cast<uintptr_t>(m_webBrowser.Get());
    }

    TabManager* manager = TabManager::Find(id);
    if (!manager) {
        return false;
    }

    TabLocation selected = manager->SelectedLocation();
    if (!selected.IsValid()) {
        return false;
    }

    const TabGroup* group = manager->GetGroup(selected.groupIndex);
    if (!group) {
        return false;
    }

    const COLORREF resolved = group->hasCustomOutline ? group->outlineColor : GetSysColor(COLOR_HOTLIGHT);
    *accent = resolved;
    *text = ChooseAccentTextColor(resolved);
    return true;
}

void CExplorerBHO::RefreshListViewAccentState() {
    if (m_listViewControl) {
        auto resolver = [this](COLORREF* accent, COLORREF* text) {
            return ResolveActiveGroupAccent(accent, text);
        };
        m_listViewControl->SetAccentColorResolver(resolver);
        m_listViewControl->SetUseAccentColors(ShouldUseListViewAccentColors());
    } else if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
    InvalidateNamespaceTreeControl();
}

bool CExplorerBHO::HandleExplorerViewMessage(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                             LRESULT* result) {
    if (!result) {
        return false;
    }

    const bool isListView = (hwnd == m_listView);
    const bool isDirectUiHost = (hwnd == m_directUiView);
    const bool isListViewHost = (m_listViewHostSubclassed.find(hwnd) != m_listViewHostSubclassed.end());
    const bool isShellViewWindow = (hwnd == m_shellViewWindow);
    const bool isGlowSurface = (m_glowSurfaces.find(hwnd) != m_glowSurfaces.end());

    if (msg == WM_TIMER) {
        if (m_explorerPaneRetryPending && wParam == m_explorerPaneRetryTimerId) {
            const size_t attempts = m_explorerPaneRetryAttempts;
            CancelExplorerPaneRetry(false);
            if (!isShellViewWindow || !IsWindow(hwnd)) {
                return false;
            }
            LogMessage(LogLevel::Info,
                       L"Explorer pane retry timer fired (view=%p attempts=%zu)", hwnd,
                       attempts);
            TryResolveExplorerPanes();
            *result = 0;
            return true;
        }

        if (m_explorerPaneFallbackPending && wParam == m_explorerPaneFallbackTimerId) {
            CancelExplorerPaneFallback();
            if (!isShellViewWindow || !IsWindow(hwnd)) {
                return false;
            }
            LogMessage(LogLevel::Info, L"Explorer pane fallback timer fired (view=%p)", hwnd);
            TryResolveExplorerPanes();
            *result = 0;
            return true;
        }
    }

    const UINT optionsChangedMessage = GetOptionsChangedMessage();
    if (optionsChangedMessage != 0 && msg == optionsChangedMessage) {
        UpdateBreadcrumbSubclass();
        if (m_breadcrumbToolbar && m_breadcrumbSubclassInstalled && IsWindow(m_breadcrumbToolbar)) {
            InvalidateRect(m_breadcrumbToolbar, nullptr, TRUE);
        }
        UpdateCurrentFolderBackground();
        InvalidateFolderBackgroundTargets();
        RefreshListViewAccentState();
        for (auto& entry : m_glowSurfaces) {
            if (entry.second) {
                entry.second->RequestRepaint();
            }
        }
        UpdateGlowSurfaceTargets();
        *result = 0;
        return true;
    }

    if ((isShellViewWindow || isDirectUiHost || isListViewHost) &&
        (msg == WM_WINDOWPOSCHANGED || msg == WM_SHOWWINDOW || msg == WM_SIZE || msg == WM_PAINT)) {
        EnsureListViewSubclass();
        UpdateGlowSurfaceTargets();
    }

    switch (msg) {
        case WM_PAINT:
        case WM_ERASEBKGND:
        case WM_PRINTCLIENT: {
            break;
        }
        case WM_PARENTNOTIFY: {
            if (LOWORD(wParam) == WM_DESTROY) {
                HWND child = reinterpret_cast<HWND>(lParam);
                if (child && child == m_statusBar) {
                    LogMessage(LogLevel::Info, L"Explorer status bar WM_DESTROY observed (hwnd=%p)", child);
                    RemoveStatusBarSubclass(child);
                    ResetStatusBarTheme(child);
                    m_statusBar = nullptr;
                }
            }
            if ((isShellViewWindow || isDirectUiHost) && (LOWORD(wParam) == WM_CREATE || LOWORD(wParam) == WM_DESTROY)) {
                EnsureListViewSubclass();
                UpdateGlowSurfaceTargets();
                if (LOWORD(wParam) == WM_CREATE) {
                    HandleExplorerPaneCandidate(reinterpret_cast<HWND>(lParam));
                }
                TryResolveExplorerPanes();
            }
            break;
        }
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
        case WM_DWMCOLORIZATIONCOLORCHANGED: {
            if (isListView) {
                RefreshListViewAccentState();
                if (msg != WM_DWMCOLORIZATIONCOLORCHANGED) {
                    RefreshListViewControlBackground();
                }
            }

            bool paletteUpdated = false;
            if (msg == WM_THEMECHANGED) {
                paletteUpdated = m_glowCoordinator.HandleThemeChanged();
            } else {
                paletteUpdated = m_glowCoordinator.HandleSettingChanged();
            }

            if (paletteUpdated) {
                for (auto& entry : m_glowSurfaces) {
                    if (entry.second) {
                        entry.second->RequestRepaint();
                    }
                }
            } else if (isGlowSurface) {
                auto glow = m_glowSurfaces.find(hwnd);
                if (glow != m_glowSurfaces.end() && glow->second) {
                    glow->second->RequestRepaint();
                }
            }

            UpdateStatusBarTheme();
            break;
        }
        case WM_DPICHANGED: {
            if (isGlowSurface) {
                auto glow = m_glowSurfaces.find(hwnd);
                if (glow != m_glowSurfaces.end() && glow->second) {
                    glow->second->RequestRepaint();
                }
            }
            break;
        }
        case WM_SIZE: {
            if (isListView) {
                RefreshListViewControlBackground();
            }
            break;
        }
        case WM_NOTIFY: {
            const NMHDR* header = reinterpret_cast<const NMHDR*>(lParam);
            if (!header) {
                break;
            }
            bool handled = false;
            if (m_paneHooks.HandleNotify(header, result)) {
                handled = true;
            }
            if (m_namespaceTreeHost && header->hwndFrom == m_namespaceTreeHost->GetWindow()) {
                if (m_namespaceTreeHost->HandleNotify(header, result)) {
                    handled = true;
                }
            }
            auto glowSurface = m_glowSurfaces.find(header->hwndFrom);
            if (glowSurface != m_glowSurfaces.end() && glowSurface->second) {
                LRESULT glowResult = 0;
                if (glowSurface->second->HandleNotify(*header, &glowResult)) {
                    *result = glowResult;
                    return true;
                }
            }
            if (m_statusBar && header->hwndFrom == m_statusBar && header->code == NM_CUSTOMDRAW) {
                handled = true;
                auto* customDraw = reinterpret_cast<NMCUSTOMDRAW*>(const_cast<NMHDR*>(header));
                if (!customDraw) {
                    *result = CDRF_DODEFAULT;
                } else if ((customDraw->dwDrawStage & CDDS_PREPAINT) == CDDS_PREPAINT) {
                    *result = CDRF_NOTIFYITEMDRAW;
                } else if ((customDraw->dwDrawStage & CDDS_ITEMPREPAINT) == CDDS_ITEMPREPAINT) {
                    if (m_statusBarThemeValid) {
                        if (m_statusBarTextColor != CLR_DEFAULT) {
                            ::SetTextColor(customDraw->hdc, m_statusBarTextColor);
                        }
                        ::SetBkMode(customDraw->hdc, TRANSPARENT);
                        if (m_statusBarBackgroundColor != CLR_DEFAULT) {
                            ::SetBkColor(customDraw->hdc, m_statusBarBackgroundColor);
                        }
                    }
                    *result = CDRF_DODEFAULT;
                } else {
                    *result = CDRF_DODEFAULT;
                }
            }
            if (handled) {
                return true;
            }
            break;
        }
        case WM_INITMENUPOPUP: {
            if (HIWORD(lParam) == 0) {
                HandleExplorerContextMenuInit(hwnd, reinterpret_cast<HMENU>(wParam));
            }
            break;
        }
        case WM_CONTEXTMENU: {
            POINT screenPoint{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            PrepareContextMenuSelection(reinterpret_cast<HWND>(wParam), screenPoint);
            break;
        }
        case WM_COMMAND: {
            const UINT commandId = LOWORD(wParam);
            if (commandId == kOpenInNewTabCommandId) {
                HandleExplorerCommand(commandId);
                *result = 0;
                return true;
            }
            break;
        }
        case WM_MENUCOMMAND: {
            HMENU menu = reinterpret_cast<HMENU>(lParam);
            const UINT position = static_cast<UINT>(wParam);
            if (menu && GetMenuItemID(menu, position) == kOpenInNewTabCommandId) {
                HandleExplorerCommand(kOpenInNewTabCommandId);
                *result = 0;
                return true;
            }
            break;
        }
        case WM_UNINITMENUPOPUP: {
            HandleExplorerMenuDismiss(reinterpret_cast<HMENU>(wParam));
            break;
        }
        case WM_CANCELMODE: {
            HandleExplorerMenuDismiss(m_trackedContextMenu);
            break;
        }
        default:
            break;
    }

    return false;
}

void CExplorerBHO::HandleExplorerContextMenuInit(HWND source, HMENU menu) {
    LogMessage(LogLevel::Info, L"Explorer context menu init (menu=%p source=%p inserted=%d tracking=%p)", menu, source,
               m_contextMenuInserted ? 1 : 0, m_trackedContextMenu);

    if (!menu) {
        LogMessage(LogLevel::Warning, L"Context menu init aborted: menu handle missing");
        return;
    }

    if (m_contextMenuInserted) {
        LogMessage(LogLevel::Info, L"Context menu init skipped: already inserted for this cycle");
        return;
    }

    if (m_trackedContextMenu && menu != m_trackedContextMenu) {
        LogMessage(LogLevel::Info, L"Context menu init skipped: still tracking menu %p", m_trackedContextMenu);
        return;
    }

    ClearPendingOpenInNewTabState();

    UINT position = 0;
    const bool anchorFound = FindOpenInNewWindowMenuItem(menu, &position, nullptr);

    ContextMenuSelectionSnapshot selection;
    CollectContextMenuSelection(selection);
    m_contextMenuSelection = selection;

    std::vector<std::wstring> folderPaths;
    folderPaths.reserve(selection.items.size());
    for (const auto& item : selection.items) {
        if (item.isFolder && !item.path.empty()) {
            if (std::find_if(folderPaths.begin(), folderPaths.end(), [&](const std::wstring& value) {
                    return _wcsicmp(value.c_str(), item.path.c_str()) == 0;
                }) == folderPaths.end()) {
                folderPaths.push_back(item.path);
            }
        }
    }

    bool insertedAny = false;
    UINT customAnchorPosition = position;
    bool customAnchorFound = anchorFound;

    if (!folderPaths.empty() && GetMenuState(menu, kOpenInNewTabCommandId, MF_BYCOMMAND) == static_cast<UINT>(-1)) {
        MENUITEMINFOW item{};
        item.cbSize = sizeof(item);
        item.fMask = MIIM_ID | MIIM_STRING | MIIM_FTYPE | MIIM_STATE;
        item.fType = MFT_STRING;
        item.fState = MFS_ENABLED;
        item.wID = kOpenInNewTabCommandId;
        item.dwTypeData = const_cast<wchar_t*>(kOpenInNewTabLabel);

        UINT insertPosition = 0;
        if (anchorFound) {
            insertPosition = position + 1;
        } else {
            LogMessage(LogLevel::Info, L"Context menu init continuing without explicit anchor");
            insertPosition = GetMenuItemCount(menu) > 0 ? static_cast<UINT>(GetMenuItemCount(menu)) : 0;
        }

        if (InsertMenuItemW(menu, insertPosition, TRUE, &item)) {
            m_pendingOpenInNewTabPaths = folderPaths;
            insertedAny = true;
            customAnchorFound = true;
            customAnchorPosition = insertPosition;
            LogMessage(LogLevel::Info, L"Open In New Tab inserted at position %u for %zu paths", insertPosition + 1,
                       m_pendingOpenInNewTabPaths.size());
        } else {
            LogLastError(L"InsertMenuItem(Open In New Tab)", GetLastError());
        }
    } else if (folderPaths.empty()) {
        LogMessage(LogLevel::Info, L"Open In New Tab not inserted: selection contains no folders");
    } else {
        LogMessage(LogLevel::Info, L"Context menu already contains Open In New Tab entry");
    }

    if (PopulateCustomContextMenus(menu, m_contextMenuSelection, customAnchorFound, customAnchorPosition)) {
        insertedAny = true;
    }

    if (insertedAny) {
        m_contextMenuInserted = true;
        m_trackedContextMenu = menu;
    } else {
        m_contextMenuSelection.Clear();
        m_pendingOpenInNewTabPaths.clear();
        LogMessage(LogLevel::Info, L"Context menu init completed without inserting custom entries");
    }
}

void CExplorerBHO::PrepareContextMenuSelection(HWND sourceWindow, POINT screenPoint) {
    HWND target = sourceWindow;
    if (!target || !IsWindow(target)) {
        target = GetFocus();
    }

    if (!target || (!IsWindow(target))) {
        return;
    }

    if (target == m_listView) {
        if (!m_listViewControl) {
            return;
        }

        if (screenPoint.x == -1 && screenPoint.y == -1) {
            return;
        }

        POINT clientPoint = screenPoint;
        if (!ScreenToClient(target, &clientPoint)) {
            return;
        }

        ShellTabsListView::HitTestResult hit{};
        if (!m_listViewControl->HitTest(clientPoint, &hit) || hit.index < 0 || (hit.flags & LVHT_ONITEM) == 0) {
            return;
        }

        if ((m_listViewControl->GetItemState(hit.index, LVIS_SELECTED) & LVIS_SELECTED) != 0) {
            return;
        }

        if (m_listViewControl->SelectExclusive(hit.index)) {
            LogMessage(LogLevel::Info, L"Context menu selection synchronized to list view item %d", hit.index);
        }
        return;
    }

    if (target == m_treeView) {
        if (screenPoint.x == -1 && screenPoint.y == -1) {
            return;
        }

        POINT clientPoint = screenPoint;
        if (!ScreenToClient(target, &clientPoint)) {
            return;
        }

        TVHITTESTINFO hit{};
        hit.pt = clientPoint;
        HTREEITEM item = TreeView_HitTest(m_treeView, &hit);
        if (!item || (hit.flags & (TVHT_ONITEM | TVHT_ONITEMBUTTON | TVHT_ONITEMINDENT)) == 0) {
            return;
        }

        HTREEITEM current = TreeView_GetSelection(m_treeView);
        if (current == item) {
            return;
        }

        TreeView_SelectItem(m_treeView, item);
        LogMessage(LogLevel::Info, L"Context menu selection synchronized to tree view item %p", item);
    }
}

void CExplorerBHO::HandleExplorerCommand(UINT commandId) {
    if (commandId != kOpenInNewTabCommandId) {
        auto it = m_contextMenuCommandMap.find(commandId);
        if (it != m_contextMenuCommandMap.end() && it->second) {
            ExecuteContextMenuCommand(*it->second);
        }
        return;
    }

    std::vector<std::wstring> paths = m_pendingOpenInNewTabPaths;
    if (paths.empty()) {
        if (!CollectSelectedFolderPaths(paths)) {
            LogMessage(LogLevel::Warning, L"Open In New Tab command aborted: unable to resolve folder selection");
            ClearPendingOpenInNewTabState();
            return;
        }
    }

    LogMessage(LogLevel::Info, L"Open In New Tab command executing for %zu paths", paths.size());
    DispatchOpenInNewTab(paths);
    ClearPendingOpenInNewTabState();
}

void CExplorerBHO::HandleExplorerMenuDismiss(HMENU menu) {
    if (!m_trackedContextMenu) {
        return;
    }

    if (!menu || menu == m_trackedContextMenu) {
        LogMessage(LogLevel::Info, L"Explorer context menu dismissed (menu=%p)", menu);
        ClearPendingOpenInNewTabState();
    }
}

bool CExplorerBHO::CollectSelectedFolderPaths(std::vector<std::wstring>& paths) const {
    paths.clear();

    ContextMenuSelectionSnapshot selection;
    if (!CollectContextMenuSelection(selection) || selection.items.empty()) {
        LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths found no eligible folders");
        return false;
    }

    for (const auto& item : selection.items) {
        if (!item.isFolder || item.path.empty()) {
            continue;
        }

        const auto existing = std::find_if(paths.begin(), paths.end(), [&](const std::wstring& value) {
            return _wcsicmp(value.c_str(), item.path.c_str()) == 0;
        });
        if (existing == paths.end()) {
            paths.push_back(item.path);
        }
    }

    if (paths.empty()) {
        LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths found no eligible folders");
        return false;
    }

    LogMessage(LogLevel::Info, L"CollectSelectedFolderPaths captured %zu path(s)", paths.size());
    return true;
}

bool CExplorerBHO::CollectContextMenuSelection(ContextMenuSelectionSnapshot& selection) const {
    selection.Clear();

    if (CollectContextSelectionFromShellView(selection) && !selection.items.empty()) {
        LogMessage(LogLevel::Info, L"CollectContextMenuSelection resolved %zu item(s) from shell view",
                   selection.items.size());
        return true;
    }

    selection.Clear();
    if (CollectContextSelectionFromFolderView(selection) && !selection.items.empty()) {
        LogMessage(LogLevel::Info, L"CollectContextMenuSelection resolved %zu item(s) from folder view",
                   selection.items.size());
        return true;
    }

    selection.Clear();
    if (CollectContextSelectionFromListView(selection) && !selection.items.empty()) {
        LogMessage(LogLevel::Info, L"CollectContextMenuSelection resolved %zu item(s) from list view",
                   selection.items.size());
        return true;
    }

    selection.Clear();
    if (CollectContextSelectionFromTreeView(selection) && !selection.items.empty()) {
        LogMessage(LogLevel::Info, L"CollectContextMenuSelection resolved %zu item(s) from tree view",
                   selection.items.size());
        return true;
    }

    LogMessage(LogLevel::Info, L"CollectContextMenuSelection found no eligible selection");
    selection.Clear();
    return false;
}

bool CExplorerBHO::CollectContextSelectionFromShellView(ContextMenuSelectionSnapshot& selection) const {
    if (!m_shellView) {
        LogMessage(LogLevel::Warning, L"CollectContextSelectionFromShellView failed: shell view unavailable");
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItemArray> items;
    HRESULT hr = m_shellView->GetItemObject(SVGIO_SELECTION, IID_PPV_ARGS(&items));
    if (FAILED(hr) || !items) {
        LogMessage(LogLevel::Info,
                   L"CollectContextSelectionFromShellView skipped: selection unavailable (hr=0x%08lX)", hr);
        return false;
    }

    return CollectContextSelectionFromItemArray(items.Get(), selection);
}

bool CExplorerBHO::CollectContextSelectionFromFolderView(ContextMenuSelectionSnapshot& selection) const {
    if (!m_shellView) {
        return false;
    }

    Microsoft::WRL::ComPtr<IFolderView2> folderView;
    HRESULT hr = m_shellView.As(&folderView);
    if (FAILED(hr) || !folderView) {
        return false;
    }

    Microsoft::WRL::ComPtr<IShellItemArray> items;
    hr = folderView->GetSelection(TRUE, &items);
    if (FAILED(hr) || !items) {
        LogMessage(LogLevel::Info,
                   L"CollectContextSelectionFromFolderView skipped: unable to resolve selection (hr=0x%08lX)", hr);
        return false;
    }

    return CollectContextSelectionFromItemArray(items.Get(), selection);
}

bool CExplorerBHO::CollectContextSelectionFromItemArray(IShellItemArray* items,
                                                        ContextMenuSelectionSnapshot& selection) const {
    if (!items) {
        return false;
    }

    DWORD count = 0;
    HRESULT hr = items->GetCount(&count);
    if (FAILED(hr) || count == 0) {
        LogMessage(LogLevel::Info, L"CollectContextSelectionFromItemArray skipped: count=%lu hr=0x%08lX", count, hr);
        return false;
    }

    bool appended = false;
    for (DWORD index = 0; index < count; ++index) {
        Microsoft::WRL::ComPtr<IShellItem> item;
        if (FAILED(items->GetItemAt(index, &item)) || !item) {
            LogMessage(LogLevel::Warning, L"CollectContextSelectionFromItemArray failed: unable to access item %lu",
                       index);
            continue;
        }

        if (AppendSelectionItemFromShellItem(item.Get(), selection)) {
            appended = true;
        }
    }

    return appended;
}

bool CExplorerBHO::CollectContextSelectionFromListView(ContextMenuSelectionSnapshot& selection) const {
    if (m_listViewControl) {
        bool appended = false;
        for (const auto& item : m_listViewControl->GetSelectionSnapshot()) {
            if (!item.pidl) {
                continue;
            }
            if (AppendSelectionItemFromPidl(item.pidl.get(), selection)) {
                appended = true;
            }
        }

        if (!appended) {
            LogMessage(LogLevel::Info, L"CollectContextSelectionFromListView found no selection");
        }

        return appended;
    }

    if (!m_listView || !IsWindow(m_listView)) {
        return false;
    }

    int index = -1;
    bool appended = false;
    while ((index = ListView_GetNextItem(m_listView, index, LVNI_SELECTED)) != -1) {
        LVITEMW item{};
        item.mask = LVIF_PARAM;
        item.iItem = index;
        if (!ListView_GetItemW(m_listView, &item)) {
            LogLastError(L"ListView_GetItem(selection)", GetLastError());
            continue;
        }

        if (AppendSelectionItemFromPidl(reinterpret_cast<PCIDLIST_ABSOLUTE>(item.lParam), selection)) {
            appended = true;
        }
    }

    if (!appended) {
        LogMessage(LogLevel::Info, L"CollectContextSelectionFromListView found no selection");
    }

    return appended;
}

bool CExplorerBHO::CollectContextSelectionFromTreeView(ContextMenuSelectionSnapshot& selection) const {
    bool appended = false;

    if (m_namespaceTreeControl) {
        Microsoft::WRL::ComPtr<IShellItemArray> items;
        const HRESULT hr = m_namespaceTreeControl->GetSelectedItems(&items);
        if (SUCCEEDED(hr) && items) {
            if (CollectContextSelectionFromItemArray(items.Get(), selection)) {
                appended = true;
            } else {
                LogMessage(LogLevel::Info,
                           L"CollectContextSelectionFromTreeView skipped: namespace tree selection produced no items");
            }
        } else {
            LogMessage(LogLevel::Info,
                       L"CollectContextSelectionFromTreeView skipped: unable to query namespace tree selection (hr=0x%08lX)",
                       hr);
        }
    }

    if (appended) {
        return true;
    }

    if (!m_treeView || !IsWindow(m_treeView)) {
        return false;
    }

    HTREEITEM selectionItem = TreeView_GetSelection(m_treeView);
    if (!selectionItem) {
        LogMessage(LogLevel::Info, L"CollectContextSelectionFromTreeView skipped: no selection");
        return false;
    }

    TVITEMEXW tvItem{};
    tvItem.mask = TVIF_PARAM;
    tvItem.hItem = selectionItem;
    if (!TreeView_GetItemW(m_treeView, &tvItem)) {
        LogLastError(L"TreeView_GetItem(selection)", GetLastError());
        return false;
    }

    TreeItemPidlResolution resolved = ResolveTreeViewItemPidl(m_treeView, tvItem);
    if (resolved.empty()) {
        LogMessage(LogLevel::Info, L"CollectContextSelectionFromTreeView skipped: selection PIDL unresolved");
        return false;
    }

    if (AppendSelectionItemFromPidl(resolved.raw, selection)) {
        return true;
    }

    LogMessage(LogLevel::Info, L"CollectContextSelectionFromTreeView skipped: selection not eligible");
    return false;
}

bool CExplorerBHO::AppendSelectionItemFromShellItem(IShellItem* item,
                                                    ContextMenuSelectionSnapshot& selection) const {
    if (!item) {
        return false;
    }

    PIDLIST_ABSOLUTE pidl = nullptr;
    HRESULT hr = SHGetIDListFromObject(item, &pidl);
    if (SUCCEEDED(hr) && pidl) {
        const bool appended = AppendSelectionItemFromPidl(pidl, selection);
        CoTaskMemFree(pidl);
        if (appended) {
            return true;
        }
    }

    PWSTR parsingName = nullptr;
    hr = item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &parsingName);
    if (FAILED(hr) || !parsingName || parsingName[0] == L'\0') {
        if (parsingName) {
            CoTaskMemFree(parsingName);
        }
        return false;
    }

    UniquePidl parsed = ParseDisplayName(parsingName);
    CoTaskMemFree(parsingName);
    if (!parsed) {
        return false;
    }

    return AppendSelectionItemFromPidl(parsed.get(), selection);
}

bool CExplorerBHO::AppendSelectionItemFromPidl(PCIDLIST_ABSOLUTE pidl,
                                               ContextMenuSelectionSnapshot& selection) const {
    if (!pidl) {
        return false;
    }

    for (const auto& existing : selection.items) {
        if (ArePidlsEqual(existing.raw, pidl)) {
            return false;
        }
    }

    ContextMenuSelectionItem entry;
    entry.pidl = ClonePidl(pidl);
    if (!entry.pidl) {
        return false;
    }
    entry.raw = entry.pidl.get();

    SHFILEINFOW info{};
    if (SHGetFileInfoW(reinterpret_cast<LPCWSTR>(pidl), 0, &info, sizeof(info), SHGFI_PIDL | SHGFI_ATTRIBUTES)) {
        entry.attributes = info.dwAttributes;
    } else {
        entry.attributes = 0;
    }

    entry.isFolder = (entry.attributes & SFGAO_FOLDER) != 0 && (entry.attributes & SFGAO_STREAM) == 0;
    entry.isFileSystem = (entry.attributes & SFGAO_FILESYSTEM) != 0;

    entry.path = GetCanonicalParsingName(pidl);
    if (entry.path.empty()) {
        entry.path = GetParsingName(pidl);
    }

    if (entry.path.empty() && entry.isFileSystem) {
        PWSTR fileSystemPath = nullptr;
        if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_FILESYSPATH, &fileSystemPath)) && fileSystemPath) {
            entry.path.assign(fileSystemPath);
        }
        if (fileSystemPath) {
            CoTaskMemFree(fileSystemPath);
        }
    }

    if (!entry.path.empty()) {
        entry.extension = ExtractLowercaseExtension(entry.path);
        entry.parentPath = ExtractParentDirectory(entry.path);

        for (const auto& existing : selection.items) {
            if (!existing.path.empty() && _wcsicmp(existing.path.c_str(), entry.path.c_str()) == 0) {
                return false;
            }
        }
    }

    selection.items.push_back(std::move(entry));
    ContextMenuSelectionItem& stored = selection.items.back();
    if (stored.isFolder) {
        ++selection.folderCount;
    } else {
        ++selection.fileCount;
    }

    return true;
}

bool CExplorerBHO::IsSelectionCountAllowed(const ContextMenuSelectionRule& rule,
                                           size_t selectionCount) const {
    const size_t minimum = rule.minimumSelection > 0 ? static_cast<size_t>(rule.minimumSelection) : 0;
    if (selectionCount < minimum) {
        return false;
    }

    if (rule.maximumSelection > 0 && selectionCount > static_cast<size_t>(rule.maximumSelection)) {
        return false;
    }

    return true;
}

bool CExplorerBHO::DoesSelectionMatchScope(const ContextMenuItemScope& scope,
                                           const ContextMenuSelectionSnapshot& selection) const {
    if (selection.items.empty()) {
        return false;
    }

    for (const auto& item : selection.items) {
        if (item.isFolder) {
            if (!scope.includeAllFolders) {
                return false;
            }
            continue;
        }

        bool fileAllowed = scope.includeAllFiles;
        if (!fileAllowed && !item.extension.empty()) {
            fileAllowed = std::find(scope.extensions.begin(), scope.extensions.end(), item.extension) !=
                          scope.extensions.end();
        }

        if (!fileAllowed) {
            return false;
        }
    }

    return true;
}

bool CExplorerBHO::ShouldDisplayMenuItem(const ContextMenuItem& item,
                                         const ContextMenuSelectionSnapshot& selection) const {
    const size_t selectionCount = selection.items.size();
    if (!IsSelectionCountAllowed(item.selection, selectionCount)) {
        return false;
    }

    if (!DoesSelectionMatchScope(item.scope, selection)) {
        return false;
    }

    if (item.type == ContextMenuItemType::kSeparator) {
        return true;
    }

    if (item.type == ContextMenuItemType::kCommand) {
        return !item.label.empty() || !item.commandTemplate.empty();
    }

    return true;
}

UINT CExplorerBHO::AllocateContextMenuCommandId(HMENU menu) {
    if (m_nextContextCommandId < kCustomCommandIdBase) {
        m_nextContextCommandId = kCustomCommandIdBase;
    }

    UINT candidate = m_nextContextCommandId;
    for (;;) {
        if (candidate == kOpenInNewTabCommandId) {
            ++candidate;
            continue;
        }

        if (menu && GetMenuState(menu, candidate, MF_BYCOMMAND) != static_cast<UINT>(-1)) {
            ++candidate;
            continue;
        }

        if (m_contextMenuCommandMap.find(candidate) != m_contextMenuCommandMap.end()) {
            ++candidate;
            continue;
        }

        break;
    }

    m_nextContextCommandId = candidate + 1;
    return candidate;
}

void CExplorerBHO::TrackContextCommand(UINT commandId, const ContextMenuItem* item) {
    if (!commandId || !item) {
        return;
    }

    m_contextMenuCommandMap[commandId] = item;
}

HBITMAP CExplorerBHO::CreateBitmapFromIcon(HICON icon, SIZE desiredSize) const {
    if (!icon || desiredSize.cx <= 0 || desiredSize.cy <= 0) {
        return nullptr;
    }

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(info.bmiHeader);
    info.bmiHeader.biWidth = desiredSize.cx;
    info.bmiHeader.biHeight = -desiredSize.cy;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HDC screen = GetDC(nullptr);
    if (!screen) {
        return nullptr;
    }

    HBITMAP bitmap = CreateDIBSection(screen, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap) {
        ReleaseDC(nullptr, screen);
        return nullptr;
    }

    if (bits) {
        std::memset(bits, 0, static_cast<size_t>(desiredSize.cx) * static_cast<size_t>(desiredSize.cy) * 4);
    }

    HDC memory = CreateCompatibleDC(screen);
    if (!memory) {
        DeleteObject(bitmap);
        ReleaseDC(nullptr, screen);
        return nullptr;
    }

    HGDIOBJ old = SelectObject(memory, bitmap);
    DrawIconEx(memory, 0, 0, icon, desiredSize.cx, desiredSize.cy, 0, nullptr, DI_NORMAL);
    SelectObject(memory, old);
    DeleteDC(memory);
    ReleaseDC(nullptr, screen);
    return bitmap;
}

void CExplorerBHO::CleanupContextMenuResources() {
    for (HBITMAP bitmap : m_contextMenuBitmaps) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
    }
    m_contextMenuBitmaps.clear();

    for (HMENU submenu : m_contextMenuSubmenus) {
        if (submenu) {
            DestroyMenu(submenu);
        }
    }
    m_contextMenuSubmenus.clear();

    m_contextMenuIconRefs.clear();
    m_contextMenuCommandMap.clear();
    m_contextMenuSelection.Clear();
    m_nextContextCommandId = 0;
}

std::optional<CExplorerBHO::PreparedMenuItem> CExplorerBHO::PrepareMenuItem(
    const ContextMenuItem& item, const ContextMenuSelectionSnapshot& selection, bool allowSubmenuAnchors) {
    if (item.type != ContextMenuItemType::kSeparator && !ShouldDisplayMenuItem(item, selection)) {
        return std::nullopt;
    }

    if (item.type == ContextMenuItemType::kSeparator &&
        !IsSelectionCountAllowed(item.selection, selection.items.size())) {
        return std::nullopt;
    }

    PreparedMenuItem prepared;
    prepared.definition = &item;
    prepared.type = item.type;
    prepared.anchor = allowSubmenuAnchors ? item.anchor : ContextMenuInsertionAnchor::kDefault;
    prepared.label = item.label;

    auto applyIcon = [&](PreparedMenuItem& target) {
        if (item.iconSource.empty()) {
            return;
        }

        IconCache::Reference iconRef = ResolveContextMenuIcon(item.iconSource, SHGFI_SMALLICON);
        if (iconRef) {
            const SIZE size = ResolveMenuIconSize(iconRef);
            HBITMAP bitmap = CreateBitmapFromIcon(iconRef.Get(), size);
            if (bitmap) {
                m_contextMenuIconRefs.push_back(iconRef);
                m_contextMenuBitmaps.push_back(bitmap);
                target.bitmap = bitmap;
            }
            return;
        }

        const std::wstring normalized = NormalizeContextMenuIconSource(item.iconSource);
        if (normalized.empty()) {
            return;
        }

        HBITMAP bitmap = reinterpret_cast<HBITMAP>(
            LoadImageW(nullptr, normalized.c_str(), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE | LR_CREATEDIBSECTION));
        if (bitmap) {
            m_contextMenuBitmaps.push_back(bitmap);
            target.bitmap = bitmap;
        }
    };

    switch (item.type) {
        case ContextMenuItemType::kCommand: {
            prepared.enabled = !item.commandTemplate.empty();
            applyIcon(prepared);
            break;
        }
        case ContextMenuItemType::kSubmenu: {
            if (item.children.empty()) {
                return std::nullopt;
            }

            HMENU submenu = CreatePopupMenu();
            if (!submenu) {
                LogLastError(L"CreatePopupMenu(custom submenu)", GetLastError());
                return std::nullopt;
            }

            if (!PopulateCustomSubmenu(submenu, item.children, selection)) {
                DestroyMenu(submenu);
                return std::nullopt;
            }

            prepared.submenu = submenu;
            m_contextMenuSubmenus.push_back(submenu);
            applyIcon(prepared);
            break;
        }
        case ContextMenuItemType::kSeparator: {
            prepared.enabled = false;
            break;
        }
    }

    return prepared;
}

bool CExplorerBHO::PopulateCustomSubmenu(HMENU submenu, const std::vector<ContextMenuItem>& items,
                                         const ContextMenuSelectionSnapshot& selection) {
    if (!submenu) {
        return false;
    }

    bool insertedAny = false;
    bool lastInsertedSeparator = false;

    auto insertPrepared = [&](PreparedMenuItem& prepared) -> bool {
        auto labelFor = [&]() -> std::wstring {
            if (!prepared.label.empty()) {
                return prepared.label;
            }
            return prepared.definition ? prepared.definition->label : std::wstring();
        };

        if (prepared.type == ContextMenuItemType::kSeparator) {
            const int count = GetMenuItemCount(submenu);
            if (count <= 0 || lastInsertedSeparator) {
                return false;
            }
        }

        MENUITEMINFOW info{};
        info.cbSize = sizeof(info);

        switch (prepared.type) {
            case ContextMenuItemType::kCommand: {
                info.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_ID | MIIM_STATE;
                info.fType = MFT_STRING;
                info.wID = AllocateContextMenuCommandId(submenu);
                prepared.commandId = info.wID;
                std::wstring label = labelFor();
                info.dwTypeData = label.empty() ? nullptr : const_cast<wchar_t*>(label.c_str());
                info.fState = prepared.enabled ? MFS_ENABLED : MFS_DISABLED;
                if (prepared.bitmap) {
                    info.fMask |= MIIM_BITMAP;
                    info.hbmpItem = prepared.bitmap;
                }
                if (!InsertMenuItemW(submenu, static_cast<UINT>(GetMenuItemCount(submenu)), TRUE, &info)) {
                    LogLastError(L"InsertMenuItem(custom submenu command)", GetLastError());
                    return false;
                }
                TrackContextCommand(prepared.commandId, prepared.definition);
                lastInsertedSeparator = false;
                return true;
            }
            case ContextMenuItemType::kSubmenu: {
                info.fMask = MIIM_FTYPE | MIIM_SUBMENU | MIIM_STRING | MIIM_STATE;
                info.fType = MFT_STRING;
                info.hSubMenu = prepared.submenu;
                std::wstring label = labelFor();
                info.dwTypeData = label.empty() ? nullptr : const_cast<wchar_t*>(label.c_str());
                info.fState = MFS_ENABLED;
                if (prepared.bitmap) {
                    info.fMask |= MIIM_BITMAP;
                    info.hbmpItem = prepared.bitmap;
                }
                if (!InsertMenuItemW(submenu, static_cast<UINT>(GetMenuItemCount(submenu)), TRUE, &info)) {
                    LogLastError(L"InsertMenuItem(custom submenu)", GetLastError());
                    return false;
                }
                lastInsertedSeparator = false;
                return true;
            }
            case ContextMenuItemType::kSeparator: {
                info.fMask = MIIM_FTYPE;
                info.fType = MFT_SEPARATOR;
                if (!InsertMenuItemW(submenu, static_cast<UINT>(GetMenuItemCount(submenu)), TRUE, &info)) {
                    LogLastError(L"InsertMenuItem(custom submenu separator)", GetLastError());
                    return false;
                }
                lastInsertedSeparator = true;
                return true;
            }
        }

        return false;
    };

    for (const auto& child : items) {
        auto prepared = PrepareMenuItem(child, selection, false);
        if (!prepared) {
            continue;
        }

        if (prepared->type == ContextMenuItemType::kSeparator && lastInsertedSeparator) {
            continue;
        }

        if (insertPrepared(*prepared)) {
            insertedAny = true;
        }
    }

    if (lastInsertedSeparator) {
        const int count = GetMenuItemCount(submenu);
        if (count > 0) {
            DeleteMenu(submenu, static_cast<UINT>(count - 1), MF_BYPOSITION);
            insertedAny = count > 1;
        }
    }

    return insertedAny;
}

bool CExplorerBHO::PopulateCustomContextMenus(HMENU menu, const ContextMenuSelectionSnapshot& selection,
                                              bool anchorFound, UINT anchorPosition) {
    if (!menu || m_cachedContextMenuItems.empty()) {
        return false;
    }

    struct AnchorState {
        bool anchorFound = false;
        UINT anchorPosition = 0;
        UINT topInsertCount = 0;
        UINT beforeShellCount = 0;
        UINT afterShellCount = 0;
    } state{anchorFound, anchorPosition, 0, 0, 0};

    auto insertPrepared = [&](PreparedMenuItem& prepared, UINT position) -> bool {
        auto labelFor = [&]() -> std::wstring {
            if (!prepared.label.empty()) {
                return prepared.label;
            }
            return prepared.definition ? prepared.definition->label : std::wstring();
        };

        if (prepared.type == ContextMenuItemType::kSeparator) {
            const int count = GetMenuItemCount(menu);
            if ((count <= 0 && position == 0) || (position > 0 && IsSeparatorItem(menu, position - 1)) ||
                (position < static_cast<UINT>(count) && IsSeparatorItem(menu, position))) {
                return false;
            }
        }

        MENUITEMINFOW info{};
        info.cbSize = sizeof(info);

        switch (prepared.type) {
            case ContextMenuItemType::kCommand: {
                info.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_ID | MIIM_STATE;
                info.fType = MFT_STRING;
                info.wID = AllocateContextMenuCommandId(menu);
                prepared.commandId = info.wID;
                std::wstring label = labelFor();
                info.dwTypeData = label.empty() ? nullptr : const_cast<wchar_t*>(label.c_str());
                info.fState = prepared.enabled ? MFS_ENABLED : MFS_DISABLED;
                if (prepared.bitmap) {
                    info.fMask |= MIIM_BITMAP;
                    info.hbmpItem = prepared.bitmap;
                }
                if (!InsertMenuItemW(menu, position, TRUE, &info)) {
                    LogLastError(L"InsertMenuItem(custom command)", GetLastError());
                    return false;
                }
                TrackContextCommand(prepared.commandId, prepared.definition);
                return true;
            }
            case ContextMenuItemType::kSubmenu: {
                info.fMask = MIIM_FTYPE | MIIM_SUBMENU | MIIM_STRING | MIIM_STATE;
                info.fType = MFT_STRING;
                info.hSubMenu = prepared.submenu;
                std::wstring label = labelFor();
                info.dwTypeData = label.empty() ? nullptr : const_cast<wchar_t*>(label.c_str());
                info.fState = MFS_ENABLED;
                if (prepared.bitmap) {
                    info.fMask |= MIIM_BITMAP;
                    info.hbmpItem = prepared.bitmap;
                }
                if (!InsertMenuItemW(menu, position, TRUE, &info)) {
                    LogLastError(L"InsertMenuItem(custom submenu)", GetLastError());
                    return false;
                }
                return true;
            }
            case ContextMenuItemType::kSeparator: {
                info.fMask = MIIM_FTYPE;
                info.fType = MFT_SEPARATOR;
                if (!InsertMenuItemW(menu, position, TRUE, &info)) {
                    LogLastError(L"InsertMenuItem(custom separator)", GetLastError());
                    return false;
                }
                return true;
            }
        }

        return false;
    };

    auto insertWithAnchor = [&](PreparedMenuItem& prepared) -> bool {
        UINT position = 0;
        switch (prepared.anchor) {
            case ContextMenuInsertionAnchor::kTop: {
                position = state.topInsertCount;
                ++state.topInsertCount;
                if (state.anchorFound) {
                    ++state.anchorPosition;
                }
                break;
            }
            case ContextMenuInsertionAnchor::kBeforeShellItems: {
                if (state.anchorFound) {
                    position = state.anchorPosition + state.beforeShellCount;
                    ++state.beforeShellCount;
                    ++state.anchorPosition;
                } else {
                    position = state.topInsertCount;
                    ++state.topInsertCount;
                }
                break;
            }
            case ContextMenuInsertionAnchor::kBottom: {
                position = static_cast<UINT>(GetMenuItemCount(menu));
                break;
            }
            case ContextMenuInsertionAnchor::kAfterShellItems:
            case ContextMenuInsertionAnchor::kDefault:
            default: {
                if (state.anchorFound) {
                    position = state.anchorPosition + 1 + state.afterShellCount;
                    ++state.afterShellCount;
                } else {
                    position = static_cast<UINT>(GetMenuItemCount(menu));
                }
                break;
            }
        }

        return insertPrepared(prepared, position);
    };

    std::vector<PreparedMenuItem> topItems;
    std::vector<PreparedMenuItem> beforeShellItems;
    std::vector<PreparedMenuItem> defaultItems;
    std::vector<PreparedMenuItem> afterShellItems;
    std::vector<PreparedMenuItem> bottomItems;

    for (const auto& definition : m_cachedContextMenuItems) {
        auto prepared = PrepareMenuItem(definition, selection, true);
        if (!prepared) {
            continue;
        }

        switch (prepared->anchor) {
            case ContextMenuInsertionAnchor::kTop:
                topItems.push_back(std::move(*prepared));
                break;
            case ContextMenuInsertionAnchor::kBeforeShellItems:
                beforeShellItems.push_back(std::move(*prepared));
                break;
            case ContextMenuInsertionAnchor::kAfterShellItems:
                afterShellItems.push_back(std::move(*prepared));
                break;
            case ContextMenuInsertionAnchor::kBottom:
                bottomItems.push_back(std::move(*prepared));
                break;
            case ContextMenuInsertionAnchor::kDefault:
            default:
                defaultItems.push_back(std::move(*prepared));
                break;
        }
    }

    bool insertedAny = false;
    auto process = [&](std::vector<PreparedMenuItem>& items) {
        for (auto& prepared : items) {
            if (insertWithAnchor(prepared)) {
                insertedAny = true;
            }
        }
    };

    process(topItems);
    process(beforeShellItems);
    process(defaultItems);
    process(afterShellItems);
    process(bottomItems);

    return insertedAny;
}

std::vector<std::wstring> CExplorerBHO::BuildCommandLines(const ContextMenuItem& item) const {
    std::vector<std::wstring> commands;
    if (item.commandTemplate.empty()) {
        return commands;
    }

    const std::wstring aggregated = ExpandAggregateTokens(item.commandTemplate);
    const bool hasSingularToken = ContainsToken(aggregated, L"%PATH%") || ContainsToken(aggregated, L"%PARENT%") ||
                                  ContainsToken(aggregated, L"%EXT%");
    const bool hasPluralToken = ContainsToken(aggregated, L"%PATHS%") || ContainsToken(aggregated, L"%PARENTS%") ||
                                ContainsToken(aggregated, L"%EXTS%");

    const size_t selectionCount = m_contextMenuSelection.items.size();

    if (hasSingularToken && selectionCount > 1 && !hasPluralToken) {
        for (const auto& selected : m_contextMenuSelection.items) {
            std::wstring expanded = ExpandCommandTemplate(aggregated, &selected);
            commands.push_back(std::move(expanded));
        }
        return commands;
    }

    const ContextMenuSelectionItem* first = selectionCount > 0 ? &m_contextMenuSelection.items.front() : nullptr;
    commands.push_back(ExpandCommandTemplate(aggregated, first));
    return commands;
}

std::wstring CExplorerBHO::ExpandAggregateTokens(const std::wstring& commandTemplate) const {
    std::wstring result = commandTemplate;

    if (ContainsToken(result, L"%COUNT%")) {
        result = ReplaceToken(result, L"%COUNT%", std::to_wstring(m_contextMenuSelection.items.size()));
    }

    if (ContainsToken(result, L"%PATHS%")) {
        std::wstring joined;
        bool first = true;
        for (const auto& item : m_contextMenuSelection.items) {
            if (item.path.empty()) {
                continue;
            }
            if (!first) {
                joined.push_back(L' ');
            }
            joined += QuoteArgument(item.path);
            first = false;
        }
        result = ReplaceToken(result, L"%PATHS%", joined);
    }

    if (ContainsToken(result, L"%PARENTS%")) {
        std::wstring joined;
        bool first = true;
        for (const auto& item : m_contextMenuSelection.items) {
            if (item.parentPath.empty()) {
                continue;
            }
            if (!first) {
                joined.push_back(L' ');
            }
            joined += QuoteArgument(item.parentPath);
            first = false;
        }
        result = ReplaceToken(result, L"%PARENTS%", joined);
    }

    if (ContainsToken(result, L"%EXTS%")) {
        std::vector<std::wstring> extensions;
        extensions.reserve(m_contextMenuSelection.items.size());
        for (const auto& item : m_contextMenuSelection.items) {
            if (!item.extension.empty() &&
                std::find(extensions.begin(), extensions.end(), item.extension) == extensions.end()) {
                extensions.push_back(item.extension);
            }
        }
        std::wstring joined;
        for (size_t i = 0; i < extensions.size(); ++i) {
            if (i > 0) {
                joined.push_back(L' ');
            }
            joined += extensions[i];
        }
        result = ReplaceToken(result, L"%EXTS%", joined);
    }

    return result;
}

std::wstring CExplorerBHO::ExpandCommandTemplate(const std::wstring& commandTemplate,
                                                 const ContextMenuSelectionItem* item) const {
    std::wstring result = commandTemplate;

    const std::wstring path = (item && !item->path.empty()) ? item->path : std::wstring();
    if (ContainsToken(result, L"%PATH%")) {
        result = ReplaceToken(result, L"%PATH%", path);
    }

    const std::wstring parent = (item && !item->parentPath.empty()) ? item->parentPath : std::wstring();
    if (ContainsToken(result, L"%PARENT%")) {
        result = ReplaceToken(result, L"%PARENT%", parent);
    }

    const std::wstring extension = (item && !item->extension.empty()) ? item->extension : std::wstring();
    if (ContainsToken(result, L"%EXT%")) {
        result = ReplaceToken(result, L"%EXT%", extension);
    }

    return result;
}

bool CExplorerBHO::ExecuteCommandLine(const std::wstring& commandLine) const {
    if (commandLine.empty()) {
        return false;
    }

    std::vector<wchar_t> buffer(commandLine.begin(), commandLine.end());
    buffer.push_back(L'\0');

    STARTUPINFOW startupInfo{};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo{};

    if (CreateProcessW(nullptr, buffer.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startupInfo,
                       &processInfo)) {
        CloseHandle(processInfo.hThread);
        CloseHandle(processInfo.hProcess);
        return true;
    }

    const DWORD processError = GetLastError();
    if (processError != ERROR_FILE_NOT_FOUND && processError != ERROR_PATH_NOT_FOUND) {
        LogLastError(L"CreateProcess(custom context command)", processError);
    }

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(commandLine.c_str(), &argc);
    if (!argv || argc <= 0) {
        if (argv) {
            LocalFree(argv);
        }
        return false;
    }

    std::wstring file = argv[0];
    std::wstring parameters;
    for (int i = 1; i < argc; ++i) {
        if (i > 1) {
            parameters.push_back(L' ');
        }
        parameters += QuoteArgument(argv[i]);
    }
    LocalFree(argv);

    SHELLEXECUTEINFOW exec{};
    exec.cbSize = sizeof(exec);
    exec.fMask = SEE_MASK_NOASYNC;
    exec.nShow = SW_SHOWNORMAL;
    exec.lpFile = file.c_str();
    exec.lpParameters = parameters.empty() ? nullptr : parameters.c_str();

    if (!ShellExecuteExW(&exec)) {
        LogLastError(L"ShellExecuteEx(custom context command)", GetLastError());
        return false;
    }

    return true;
}

void CExplorerBHO::ExecuteContextMenuCommand(const ContextMenuItem& item) const {
    const std::vector<std::wstring> commands = BuildCommandLines(item);
    if (commands.empty()) {
        LogMessage(LogLevel::Warning, L"ExecuteContextMenuCommand skipped: no command lines generated");
        return;
    }

    size_t succeeded = 0;
    for (const auto& commandLine : commands) {
        if (commandLine.empty()) {
            continue;
        }

        if (ExecuteCommandLine(commandLine)) {
            ++succeeded;
            LogMessage(LogLevel::Info, L"ExecuteContextMenuCommand launched: %ls", commandLine.c_str());
        } else {
            LogMessage(LogLevel::Warning, L"ExecuteContextMenuCommand failed: %ls", commandLine.c_str());
        }
    }

    if (succeeded == 0) {
        LogMessage(LogLevel::Warning, L"ExecuteContextMenuCommand failed for all generated commands");
    }
}


bool CExplorerBHO::ResolveHighlightFromPidl(PCIDLIST_ABSOLUTE pidl, PaneHighlight* highlight) const {
    if (!pidl) {
        return false;
    }

    std::vector<std::wstring> paths;
    if (!AppendPathFromPidl(pidl, paths) || paths.empty()) {
        return false;
    }

    const std::wstring normalizedPath = NormalizePaneHighlightKey(paths.front());
    if (normalizedPath.empty()) {
        return false;
    }

    return TryGetPaneHighlight(normalizedPath, highlight);
}

bool CExplorerBHO::AppendPathFromPidl(PCIDLIST_ABSOLUTE pidl, std::vector<std::wstring>& paths) const {
    if (!pidl) {
        return false;
    }

    std::wstring value;

    enum class FileSystemFailureReason {
        kNone,
        kBindFailed,
        kAttributeMismatch,
        kPathResolutionFailed,
    } failureReason = FileSystemFailureReason::kNone;

    HRESULT failureHr = S_OK;
    SFGAOF failureAttributes = 0;

    Microsoft::WRL::ComPtr<IShellFolder> parentFolder;
    PCUITEMID_CHILD child = nullptr;
    HRESULT hr = SHBindToParent(pidl, IID_PPV_ARGS(&parentFolder), &child);
    if (FAILED(hr) || !parentFolder || !child) {
        failureReason = FileSystemFailureReason::kBindFailed;
        failureHr = hr;
    } else {
        SFGAOF attributes = SFGAO_FOLDER | SFGAO_FILESYSTEM;
        hr = parentFolder->GetAttributesOf(1, &child, &attributes);
        if (FAILED(hr) || (attributes & SFGAO_FOLDER) == 0 || (attributes & SFGAO_FILESYSTEM) == 0) {
            failureReason = FileSystemFailureReason::kAttributeMismatch;
            failureHr = hr;
            failureAttributes = attributes;
        } else {
            PWSTR path = nullptr;
            hr = SHGetNameFromIDList(pidl, SIGDN_FILESYSPATH, &path);
            if (FAILED(hr) || !path || path[0] == L'\0') {
                failureReason = FileSystemFailureReason::kPathResolutionFailed;
                failureHr = hr;
            } else {
                value.assign(path);
            }
            if (path) {
                CoTaskMemFree(path);
            }
        }
    }

    if (value.empty()) {
        if (auto translated = TranslateVirtualLocation(pidl)) {
            value = std::move(*translated);
        }
    }

    if (value.empty()) {
        switch (failureReason) {
            case FileSystemFailureReason::kBindFailed:
                LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: unable to bind to parent (hr=0x%08lX)", failureHr);
                break;
            case FileSystemFailureReason::kAttributeMismatch:
                LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: attributes=0x%08lX (hr=0x%08lX)", failureAttributes,
                           failureHr);
                break;
            case FileSystemFailureReason::kPathResolutionFailed:
                LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: unable to resolve filesystem path (hr=0x%08lX)",
                           failureHr);
                break;
            case FileSystemFailureReason::kNone:
                LogMessage(LogLevel::Info, L"AppendPathFromPidl skipped: unsupported namespace");
                break;
        }
        return false;
    }

    if (std::find(paths.begin(), paths.end(), value) != paths.end()) {
        return true;
    }

    paths.push_back(std::move(value));
    return true;
}

void CExplorerBHO::DispatchOpenInNewTab(const std::vector<std::wstring>& paths) {
    if (paths.empty()) {
        LogMessage(LogLevel::Info, L"DispatchOpenInNewTab skipped: no paths provided");
        return;
    }

    QueueOpenInNewTabRequests(paths);
    TryDispatchQueuedOpenInNewTabRequests();
}

void CExplorerBHO::QueueOpenInNewTabRequests(const std::vector<std::wstring>& paths) {
    size_t added = 0;
    for (const std::wstring& path : paths) {
        if (path.empty()) {
            LogMessage(LogLevel::Warning, L"QueueOpenInNewTabRequests skipped empty path entry");
            continue;
        }

        m_openInNewTabQueue.push_back(path);
        ++added;
    }

    if (added > 0) {
        LogMessage(LogLevel::Info, L"Queued %zu Open In New Tab request(s); %zu pending", added,
                   m_openInNewTabQueue.size());
    }
}

void CExplorerBHO::TryDispatchQueuedOpenInNewTabRequests() {
    if (m_openInNewTabQueue.empty()) {
        CancelOpenInNewTabRetry();
        return;
    }

    HWND frame = GetTopLevelExplorerWindow();
    if (!frame) {
        LogMessage(LogLevel::Warning,
                   L"Open In New Tab dispatch deferred: explorer frame not found (%zu request(s) pending)",
                   m_openInNewTabQueue.size());
        ScheduleOpenInNewTabRetry();
        return;
    }

    HWND bandWindow = FindDescendantWindow(frame, L"ShellTabsBandWindow");
    if (!bandWindow || !IsWindow(bandWindow)) {
        LogMessage(LogLevel::Info,
                   L"Open In New Tab dispatch deferred: ShellTabs band window missing (frame=%p, pending=%zu)",
                   frame, m_openInNewTabQueue.size());
        m_shouldRetryEnsure = true;
        EnsureBandVisible();
        ScheduleOpenInNewTabRetry();
        return;
    }

    std::vector<std::wstring> pending;
    pending.swap(m_openInNewTabQueue);
    CancelOpenInNewTabRetry();

    for (const std::wstring& path : pending) {
        if (path.empty()) {
            continue;
        }

        OpenFolderMessagePayload payload{path.c_str(), path.size()};
        SendMessageW(bandWindow, WM_SHELLTABS_OPEN_FOLDER, reinterpret_cast<WPARAM>(&payload), 0);
        LogMessage(LogLevel::Info, L"Dispatched Open In New Tab request for %ls", path.c_str());
    }
}

void CExplorerBHO::ClearPendingOpenInNewTabState() {
    CleanupContextMenuResources();
    m_pendingOpenInNewTabPaths.clear();
    m_trackedContextMenu = nullptr;
    m_contextMenuInserted = false;
    LogMessage(LogLevel::Info, L"Cleared Open In New Tab pending state");
}

bool CExplorerBHO::InstallBreadcrumbSubclass(HWND toolbar) {
    if (!toolbar || !IsWindow(toolbar)) {
        return false;
    }

    if (toolbar == m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        return true;
    }

    RemoveBreadcrumbSubclass();

    if (SetWindowSubclass(toolbar, &CExplorerBHO::BreadcrumbSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_breadcrumbToolbar = toolbar;
        m_breadcrumbSubclassInstalled = true;
        m_loggedBreadcrumbToolbarMissing = false;
        LogMessage(LogLevel::Info, L"Installed breadcrumb gradient subclass on hwnd=%p", toolbar);
        InvalidateRect(toolbar, nullptr, TRUE);
        UpdateAddressEditSubclass();
        return true;
    }

    LogLastError(L"SetWindowSubclass(breadcrumb toolbar)", GetLastError());
    return false;
}

bool CExplorerBHO::InstallProgressSubclass(HWND progressWindow) {
    if (!progressWindow || !IsWindow(progressWindow)) {
        return false;
    }

    if (SetWindowSubclass(progressWindow, &CExplorerBHO::ProgressSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_progressWindow = progressWindow;
        m_progressSubclassInstalled = true;
        if (!EnsureProgressGradientResources()) {
            LogMessage(LogLevel::Warning,
                       L"Progress gradient resources unavailable; falling back to on-demand rendering");
        }
        LogMessage(LogLevel::Info, L"Installed progress gradient subclass on hwnd=%p", progressWindow);
        return true;
    }

    LogLastError(L"SetWindowSubclass(progress window)", GetLastError());
    return false;
}

void CExplorerBHO::RemoveBreadcrumbSubclass() {
    if (m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        if (IsWindow(m_breadcrumbToolbar)) {
            RemoveWindowSubclass(m_breadcrumbToolbar, &CExplorerBHO::BreadcrumbSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_breadcrumbToolbar, nullptr, TRUE);
        }
    }
    m_breadcrumbToolbar = nullptr;
    m_breadcrumbSubclassInstalled = false;
    if (m_breadcrumbLogState == BreadcrumbLogState::Searching) {
        m_breadcrumbLogState = BreadcrumbLogState::Unknown;
    }
    m_loggedBreadcrumbToolbarMissing = false;

    RemoveAddressEditSubclass();
    RemoveProgressSubclass();
}

void CExplorerBHO::RemoveProgressSubclass() {
    if (m_progressWindow && m_progressSubclassInstalled) {
        if (IsWindow(m_progressWindow)) {
            RemoveWindowSubclass(m_progressWindow, &CExplorerBHO::ProgressSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_progressWindow, nullptr, TRUE);
        }
    }
    m_progressWindow = nullptr;
    m_progressSubclassInstalled = false;
    DestroyProgressGradientResources();
}

bool CExplorerBHO::EnsureProgressGradientResources() {
    if (!m_useCustomProgressGradientColors) {
        return false;
    }

    if (m_progressGradientBitmap && m_progressGradientBitmapStartColor == m_progressGradientStartColor &&
        m_progressGradientBitmapEndColor == m_progressGradientEndColor && m_progressGradientBits) {
        return true;
    }

    DestroyProgressGradientResources();

    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = kProgressGradientSampleWidth;
    info.bmiHeader.biHeight = -1;  // top-down
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (!bitmap || !bits) {
        if (bitmap) {
            DeleteObject(bitmap);
        }
        return false;
    }

    auto* pixels = static_cast<DWORD*>(bits);
    const BYTE startRed = GetRValue(m_progressGradientStartColor);
    const BYTE startGreen = GetGValue(m_progressGradientStartColor);
    const BYTE startBlue = GetBValue(m_progressGradientStartColor);
    const BYTE endRed = GetRValue(m_progressGradientEndColor);
    const BYTE endGreen = GetGValue(m_progressGradientEndColor);
    const BYTE endBlue = GetBValue(m_progressGradientEndColor);

    for (int x = 0; x < kProgressGradientSampleWidth; ++x) {
        const double t = (kProgressGradientSampleWidth > 1)
                             ? static_cast<double>(x) / static_cast<double>(kProgressGradientSampleWidth - 1)
                             : 0.0;
        const BYTE red = static_cast<BYTE>(std::clamp<int>(
            static_cast<int>(std::lround(static_cast<double>(startRed) +
                                         (static_cast<double>(endRed) - static_cast<double>(startRed)) * t)),
            0, 255));
        const BYTE green = static_cast<BYTE>(std::clamp<int>(
            static_cast<int>(std::lround(static_cast<double>(startGreen) +
                                         (static_cast<double>(endGreen) - static_cast<double>(startGreen)) * t)),
            0, 255));
        const BYTE blue = static_cast<BYTE>(std::clamp<int>(
            static_cast<int>(std::lround(static_cast<double>(startBlue) +
                                         (static_cast<double>(endBlue) - static_cast<double>(startBlue)) * t)),
            0, 255));
        pixels[x] = (static_cast<DWORD>(blue)) | (static_cast<DWORD>(green) << 8) | (static_cast<DWORD>(red) << 16) |
                    0xFF000000;
    }

    m_progressGradientBitmap = bitmap;
    m_progressGradientBits = bits;
    m_progressGradientInfo = info;
    m_progressGradientBitmapStartColor = m_progressGradientStartColor;
    m_progressGradientBitmapEndColor = m_progressGradientEndColor;
    return true;
}

void CExplorerBHO::DestroyProgressGradientResources() {
    if (m_progressGradientBitmap) {
        DeleteObject(m_progressGradientBitmap);
        m_progressGradientBitmap = nullptr;
    }
    m_progressGradientBits = nullptr;
    m_progressGradientInfo = BITMAPINFO{};
    m_progressGradientBitmapStartColor = 0;
    m_progressGradientBitmapEndColor = 0;
}

bool CExplorerBHO::InstallAddressEditSubclass(HWND editWindow) {
    if (!editWindow || !IsWindow(editWindow)) {
        return false;
    }

    if (SetWindowSubclass(editWindow, &CExplorerBHO::AddressEditSubclassProc, reinterpret_cast<UINT_PTR>(this), 0)) {
        m_addressEditWindow = editWindow;
        m_addressEditSubclassInstalled = true;
        ResetAddressEditStateCache();
        RefreshAddressEditFont(editWindow);
        RefreshAddressEditState(editWindow, true, true, true, true);
        LogMessage(LogLevel::Info, L"Installed address edit gradient subclass on hwnd=%p", editWindow);
        return true;
    }

    LogLastError(L"SetWindowSubclass(address edit)", GetLastError());
    return false;
}

void CExplorerBHO::RemoveAddressEditSubclass() {
    if (m_addressEditWindow && m_addressEditSubclassInstalled) {
        ResetAddressEditStateCache();
        if (IsWindow(m_addressEditWindow)) {
            RemoveWindowSubclass(m_addressEditWindow, &CExplorerBHO::AddressEditSubclassProc,
                                 reinterpret_cast<UINT_PTR>(this));
            InvalidateRect(m_addressEditWindow, nullptr, TRUE);
        }
    } else {
        ResetAddressEditStateCache();
    }
    m_addressEditWindow = nullptr;
    m_addressEditSubclassInstalled = false;
}

void CExplorerBHO::RequestAddressEditRedraw(HWND hwnd) const {
    if (!hwnd || !IsWindow(hwnd)) {
        return;
    }

    if (!m_breadcrumbFontGradientEnabled) {
        return;
    }

    if (m_addressEditRedrawPending) {
        if (m_addressEditRedrawTimerActive) {
            if (!SetTimer(hwnd, kAddressEditRedrawTimerId, kAddressEditRedrawCoalesceDelayMs, nullptr)) {
                m_addressEditRedrawTimerActive = false;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
        }
        return;
    }

    m_addressEditRedrawPending = true;

    if (SetTimer(hwnd, kAddressEditRedrawTimerId, kAddressEditRedrawCoalesceDelayMs, nullptr)) {
        m_addressEditRedrawTimerActive = true;
        return;
    }

    m_addressEditRedrawTimerActive = false;
    InvalidateRect(hwnd, nullptr, FALSE);
}

void CExplorerBHO::ResetAddressEditStateCache() {
    if (m_addressEditWindow && m_addressEditRedrawTimerActive) {
        KillTimer(m_addressEditWindow, kAddressEditRedrawTimerId);
    }

    m_addressEditRedrawTimerActive = false;
    m_addressEditRedrawPending = false;
    m_addressEditCachedText.clear();
    m_addressEditCachedSelStart = 0;
    m_addressEditCachedSelEnd = 0;
    m_addressEditCachedHasFocus = false;
    m_addressEditCachedThemeActive = IsThemeActive() != FALSE;
    m_addressEditCachedFont = nullptr;
}

bool CExplorerBHO::RefreshAddressEditState(HWND hwnd, bool updateText, bool updateSelection,
                                           bool updateFocus, bool updateTheme) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    bool changed = false;
    if (updateText) {
        changed |= RefreshAddressEditText(hwnd);
    }
    if (updateSelection) {
        changed |= RefreshAddressEditSelection(hwnd);
    }
    if (updateFocus) {
        changed |= RefreshAddressEditFocus(hwnd);
    }
    if (updateTheme) {
        changed |= RefreshAddressEditTheme();
    }

    return changed;
}

bool CExplorerBHO::RefreshAddressEditText(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    int length = GetWindowTextLengthW(hwnd);
    if (length < 0) {
        length = 0;
    }

    std::wstring text;
    if (length > 0) {
        text.resize(static_cast<size_t>(length) + 1);
        int copied = GetWindowTextW(hwnd, text.data(), length + 1);
        if (copied < 0) {
            copied = 0;
        }
        text.resize(static_cast<size_t>(copied));
    }

    if (text != m_addressEditCachedText) {
        m_addressEditCachedText = std::move(text);
        return true;
    }

    return false;
}

bool CExplorerBHO::RefreshAddressEditSelection(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    DWORD selectionStart = 0;
    DWORD selectionEnd = 0;
    SendMessageW(hwnd, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart),
                 reinterpret_cast<LPARAM>(&selectionEnd));

    if (selectionStart != m_addressEditCachedSelStart || selectionEnd != m_addressEditCachedSelEnd) {
        m_addressEditCachedSelStart = selectionStart;
        m_addressEditCachedSelEnd = selectionEnd;
        return true;
    }

    return false;
}

bool CExplorerBHO::RefreshAddressEditFocus(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        return false;
    }

    bool hasFocus = GetFocus() == hwnd;
    if (hasFocus != m_addressEditCachedHasFocus) {
        m_addressEditCachedHasFocus = hasFocus;
        return true;
    }

    return false;
}

bool CExplorerBHO::RefreshAddressEditTheme() {
    bool themeActive = IsThemeActive() != FALSE;
    if (themeActive != m_addressEditCachedThemeActive) {
        m_addressEditCachedThemeActive = themeActive;
        return true;
    }

    return false;
}

bool CExplorerBHO::RefreshAddressEditFont(HWND hwnd) {
    if (!hwnd || !IsWindow(hwnd)) {
        m_addressEditCachedFont = nullptr;
        return false;
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    if (font != m_addressEditCachedFont) {
        m_addressEditCachedFont = font;
        return true;
    }

    return false;
}

void CExplorerBHO::UpdateAddressEditSubclass() {
    if (!m_breadcrumbFontGradientEnabled || !m_gdiplusInitialized) {
        RemoveAddressEditSubclass();
        return;
    }

    HWND edit = FindAddressEditControl();
    if (!edit) {
        RemoveAddressEditSubclass();
        return;
    }

    if (m_addressEditSubclassInstalled && edit == m_addressEditWindow && IsWindow(edit)) {
        InvalidateRect(edit, nullptr, TRUE);
        return;
    }

    RemoveAddressEditSubclass();
    if (InstallAddressEditSubclass(edit)) {
        InvalidateRect(edit, nullptr, TRUE);
    }
}

void CExplorerBHO::EnsureBreadcrumbHook() {
    if (m_breadcrumbHookRegistered) {
        return;
    }

    const DWORD threadId = GetCurrentThreadId();
    std::lock_guard<std::mutex> lock(g_breadcrumbHookMutex);
    auto& entry = g_breadcrumbHooks[threadId];
    if (std::find(entry.observers.begin(), entry.observers.end(), this) == entry.observers.end()) {
        entry.observers.push_back(this);
    }

    if (!entry.hook) {
        HHOOK hook = SetWindowsHookExW(WH_CBT, BreadcrumbCbtProc, nullptr, threadId);
        if (!hook) {
            LogLastError(L"SetWindowsHookEx(WH_CBT)", GetLastError());
            entry.observers.erase(std::remove(entry.observers.begin(), entry.observers.end(), this), entry.observers.end());
            if (entry.observers.empty()) {
                g_breadcrumbHooks.erase(threadId);
            }
            return;
        }
        entry.hook = hook;
        LogMessage(LogLevel::Info, L"Breadcrumb CBT hook installed for thread %lu", threadId);
    }

    m_breadcrumbHookRegistered = true;
}

void CExplorerBHO::RemoveBreadcrumbHook() {
    if (!m_breadcrumbHookRegistered) {
        return;
    }

    const DWORD threadId = GetCurrentThreadId();
    std::lock_guard<std::mutex> lock(g_breadcrumbHookMutex);
    auto it = g_breadcrumbHooks.find(threadId);
    if (it == g_breadcrumbHooks.end()) {
        m_breadcrumbHookRegistered = false;
        return;
    }

    auto& observers = it->second.observers;
    observers.erase(std::remove(observers.begin(), observers.end(), this), observers.end());
    if (observers.empty()) {
        if (it->second.hook) {
            UnhookWindowsHookEx(it->second.hook);
        }
        g_breadcrumbHooks.erase(it);
        LogMessage(LogLevel::Info, L"Breadcrumb CBT hook removed for thread %lu", threadId);
    }

    m_breadcrumbHookRegistered = false;
}

void CExplorerBHO::UpdateBreadcrumbSubclass() {
    auto& store = OptionsStore::Instance();
    static bool loggedOptionsLoadFailure = false;
    std::wstring errorContext;
    if (!store.Load(&errorContext)) {
        if (!loggedOptionsLoadFailure) {
            if (!errorContext.empty()) {
                LogMessage(LogLevel::Warning,
                           L"CExplorerBHO::UpdateBreadcrumbSubclass failed to load options: %ls",
                           errorContext.c_str());
            } else {
                LogMessage(LogLevel::Warning,
                           L"CExplorerBHO::UpdateBreadcrumbSubclass failed to load options");
            }
            loggedOptionsLoadFailure = true;
        }
    } else if (loggedOptionsLoadFailure) {
        loggedOptionsLoadFailure = false;
    }
    const ShellTabsOptions options = store.Get();
    const bool previousBreadcrumbFontGradientEnabled = m_breadcrumbFontGradientEnabled;
    const int previousBreadcrumbFontBrightness = m_breadcrumbFontBrightness;
    const bool previousUseCustomFontColors = m_useCustomBreadcrumbFontColors;
    const bool previousUseCustomGradientColors = m_useCustomBreadcrumbGradientColors;
    const COLORREF previousBreadcrumbFontGradientStart = m_breadcrumbFontGradientStartColor;
    const COLORREF previousBreadcrumbFontGradientEnd = m_breadcrumbFontGradientEndColor;
    const COLORREF previousBreadcrumbGradientStart = m_breadcrumbGradientStartColor;
    const COLORREF previousBreadcrumbGradientEnd = m_breadcrumbGradientEndColor;
    m_glowCoordinator.Configure(options);
    m_breadcrumbGradientEnabled = options.enableBreadcrumbGradient;
    m_breadcrumbFontGradientEnabled = options.enableBreadcrumbFontGradient;
    m_breadcrumbGradientTransparency = std::clamp<int>(options.breadcrumbGradientTransparency, 0, 100);
    m_breadcrumbFontBrightness = std::clamp<int>(options.breadcrumbFontBrightness, 0, 100);
    m_breadcrumbHighlightAlphaMultiplier =
        std::clamp<int>(options.breadcrumbHighlightAlphaMultiplier, 0, 200);
    m_breadcrumbDropdownAlphaMultiplier =
        std::clamp<int>(options.breadcrumbDropdownAlphaMultiplier, 0, 200);
    m_useCustomBreadcrumbGradientColors = options.useCustomBreadcrumbGradientColors;
    m_breadcrumbGradientStartColor = options.breadcrumbGradientStartColor;
    m_breadcrumbGradientEndColor = options.breadcrumbGradientEndColor;
    m_useCustomBreadcrumbFontColors = options.useCustomBreadcrumbFontColors;
    m_breadcrumbFontGradientStartColor = options.breadcrumbFontGradientStartColor;
    m_breadcrumbFontGradientEndColor = options.breadcrumbFontGradientEndColor;
    m_useCustomProgressGradientColors = options.useCustomProgressBarGradientColors;
    m_progressGradientStartColor = options.progressBarGradientStartColor;
    m_progressGradientEndColor = options.progressBarGradientEndColor;
    const bool breadcrumbFontGradientChanged =
        previousBreadcrumbFontGradientEnabled != m_breadcrumbFontGradientEnabled ||
        previousBreadcrumbFontBrightness != m_breadcrumbFontBrightness ||
        previousUseCustomFontColors != m_useCustomBreadcrumbFontColors ||
        previousUseCustomGradientColors != m_useCustomBreadcrumbGradientColors ||
        previousBreadcrumbFontGradientStart != m_breadcrumbFontGradientStartColor ||
        previousBreadcrumbFontGradientEnd != m_breadcrumbFontGradientEndColor ||
        previousBreadcrumbGradientStart != m_breadcrumbGradientStartColor ||
        previousBreadcrumbGradientEnd != m_breadcrumbGradientEndColor;
    if (breadcrumbFontGradientChanged) {
        RequestHeaderGlowRepaint();
    }
    const bool previousAccentSetting = m_useExplorerAccentColors;
    m_useExplorerAccentColors = options.useExplorerAccentColors;
    if (previousAccentSetting != m_useExplorerAccentColors) {
        RefreshListViewAccentState();
    }

    m_cachedContextMenuItems = options.contextMenuItems;

    ReloadFolderBackgrounds(options);
    UpdateCurrentFolderBackground();

    UpdateProgressSubclass();

    const bool gradientsEnabled = (m_breadcrumbGradientEnabled || m_breadcrumbFontGradientEnabled);
    if (!gradientsEnabled || !m_gdiplusInitialized) {
        if (m_breadcrumbLogState != BreadcrumbLogState::Disabled) {
            LogMessage(LogLevel::Info,
                       L"Breadcrumb gradients inactive (background=%d text=%d gdiplus=%d); ensuring subclass removed",
                       m_breadcrumbGradientEnabled ? 1 : 0, m_breadcrumbFontGradientEnabled ? 1 : 0,
                       m_gdiplusInitialized ? 1 : 0);
            m_breadcrumbLogState = BreadcrumbLogState::Disabled;
        }
        if (m_breadcrumbSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Breadcrumb gradients disabled; removing subclass");
        }
        RemoveBreadcrumbHook();
        RemoveBreadcrumbSubclass();
        m_loggedBreadcrumbToolbarMissing = false;
        return;
    }

    EnsureBreadcrumbHook();

    if (m_breadcrumbLogState != BreadcrumbLogState::Searching) {
        LogMessage(LogLevel::Info,
                   L"Breadcrumb gradients enabled; locating toolbar (installed=%d background=%d text=%d)",
                   m_breadcrumbSubclassInstalled ? 1 : 0, m_breadcrumbGradientEnabled ? 1 : 0,
                   m_breadcrumbFontGradientEnabled ? 1 : 0);
        m_lastBreadcrumbStage = BreadcrumbDiscoveryStage::None;
        m_breadcrumbLogState = BreadcrumbLogState::Searching;
    }

    HWND toolbar = FindBreadcrumbToolbar();
    if (!toolbar) {
        if (!m_loggedBreadcrumbToolbarMissing) {
            LogMessage(LogLevel::Info, L"Breadcrumb toolbar not yet available; will retry");
            m_loggedBreadcrumbToolbarMissing = true;
        }
        if (m_breadcrumbSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Breadcrumb toolbar not found; removing subclass");
        }
        RemoveBreadcrumbSubclass();
        return;
    }

    if (m_loggedBreadcrumbToolbarMissing) {
        LogMessage(LogLevel::Info, L"Breadcrumb toolbar discovered after retry");
    }
    m_loggedBreadcrumbToolbarMissing = false;

    if (toolbar == m_breadcrumbToolbar && m_breadcrumbSubclassInstalled) {
        InvalidateRect(toolbar, nullptr, TRUE);
        UpdateProgressSubclass();
        UpdateAddressEditSubclass();
        return;
    }

    InstallBreadcrumbSubclass(toolbar);
    UpdateProgressSubclass();
    UpdateAddressEditSubclass();
}

void CExplorerBHO::UpdateProgressSubclass() {
    if (!m_useCustomProgressGradientColors) {
        if (m_progressSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Progress gradients disabled; removing subclass");
        }
        RemoveProgressSubclass();
        return;
    }

    HWND progress = FindProgressWindow();
    if (!progress) {
        if (m_progressSubclassInstalled) {
            LogMessage(LogLevel::Info, L"Progress window not found; removing subclass");
        }
        RemoveProgressSubclass();
        return;
    }

    if (m_progressSubclassInstalled && progress == m_progressWindow) {
        InvalidateRect(progress, nullptr, TRUE);
        return;
    }

    RemoveProgressSubclass();
    if (InstallProgressSubclass(progress)) {
        InvalidateRect(progress, nullptr, TRUE);
    }
}

bool CExplorerBHO::HandleBreadcrumbPaint(HWND hwnd) {
    if ((!m_breadcrumbGradientEnabled && !m_breadcrumbFontGradientEnabled) || !m_gdiplusInitialized) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC target = BeginPaint(hwnd, &ps);
    if (!target) {
        return true;
    }

    RECT client{};
    GetClientRect(hwnd, &client);

    BP_PAINTPARAMS params{};
    params.cbSize = sizeof(params);
    params.dwFlags = BPPF_ERASE;

    HDC paintDc = nullptr;
    HPAINTBUFFER buffer = BeginBufferedPaint(target, &client, BPBF_TOPDOWNDIB, &params, &paintDc);
    HDC drawDc = paintDc ? paintDc : target;

    HRESULT backgroundResult = DrawThemeParentBackground(hwnd, drawDc, &client);
    if (FAILED(backgroundResult)) {
        HBRUSH brush = reinterpret_cast<HBRUSH>(GetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND));
        if (!brush) {
            brush = GetSysColorBrush(COLOR_WINDOW);
        }
        FillRect(drawDc, &client, brush);
    }

    Gdiplus::Graphics graphics(drawDc);
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        if (buffer) {
            EndBufferedPaint(buffer, TRUE);
        }
        EndPaint(hwnd, &ps);
        return true;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
    graphics.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

    const int highlightAlphaMultiplier = std::clamp<int>(m_breadcrumbHighlightAlphaMultiplier, 0, 200);
    const int dropdownAlphaMultiplier = std::clamp<int>(m_breadcrumbDropdownAlphaMultiplier, 0, 200);
    auto scaleAlpha = [](BYTE alpha, int multiplier) -> BYTE {
        if (multiplier <= 0) {
            return 0;
        }
        if (multiplier == 100) {
            return alpha;
        }
        const int scaled = static_cast<int>(alpha) * multiplier;
        const int result = (scaled + 50) / 100;
        return static_cast<BYTE>(std::clamp<int>(result, 0, 255));
    };

    auto drawDropdownArrow = [&](const RECT& buttonRect, bool hot, bool pressed, BYTE textAlphaValue,
                                 const Gdiplus::Color& brightFontEndColor, const Gdiplus::Color& arrowTextStart,
                                 const Gdiplus::Color& arrowTextEnd, bool fontGradientEnabled,
                                 bool backgroundGradientEnabled, bool backgroundGradientVisible, BYTE gradientAlpha,
                                 COLORREF highlightColorRef) {
        const float arrowWidth = 6.0f;
        const float arrowHeight = 4.0f;
        const float rectWidth = static_cast<float>(buttonRect.right - buttonRect.left);
        const float rectHeight = static_cast<float>(buttonRect.bottom - buttonRect.top);
        const float centerX = static_cast<float>(buttonRect.left) + rectWidth - 9.0f;
        const float centerY = static_cast<float>(buttonRect.top) + rectHeight / 2.0f;

        if (hot || pressed) {
            const float highlightWidth = arrowWidth + 6.0f;
            constexpr Gdiplus::REAL kHighlightVerticalDeflate = 1.0f;
            const Gdiplus::REAL buttonHeight = static_cast<Gdiplus::REAL>(rectHeight);
            Gdiplus::REAL highlightTop = static_cast<Gdiplus::REAL>(buttonRect.top) + kHighlightVerticalDeflate;
            Gdiplus::REAL highlightHeight = buttonHeight - (kHighlightVerticalDeflate * 2.0f);
            if (highlightHeight < 0.0f) {
                highlightHeight = 0.0f;
            }
            if (highlightHeight < 4.0f) {
                highlightHeight = std::min<Gdiplus::REAL>(4.0f, buttonHeight);
                const Gdiplus::REAL highlightMid =
                    static_cast<Gdiplus::REAL>(buttonRect.top) + (buttonHeight / 2.0f);
                highlightTop = highlightMid - (highlightHeight / 2.0f);
            }
            if (highlightHeight > 0.0f) {
                Gdiplus::RectF highlightRect(centerX - highlightWidth / 2.0f, highlightTop, highlightWidth,
                                             highlightHeight);
                const BYTE highlightAlpha = scaleAlpha(static_cast<BYTE>(pressed ? 160 : 130),
                                                       highlightAlphaMultiplier);
                const Gdiplus::Color highlightBase(highlightAlpha, brightFontEndColor.GetR(),
                                                   brightFontEndColor.GetG(), brightFontEndColor.GetB());
                Gdiplus::Color highlightColor =
                    BrightenBreadcrumbColor(highlightBase, hot, pressed, highlightColorRef);
                Gdiplus::SolidBrush highlightBrush(highlightColor);
                graphics.FillRectangle(&highlightBrush, highlightRect);
            }
        }

        Gdiplus::PointF arrow[3] = {
            {centerX - arrowWidth / 2.0f, centerY - arrowHeight / 2.0f},
            {centerX + arrowWidth / 2.0f, centerY - arrowHeight / 2.0f},
            {centerX, centerY + arrowHeight / 2.0f},
        };

        Gdiplus::RectF arrowRect(centerX - arrowWidth / 2.0f,
                                 centerY - arrowHeight / 2.0f,
                                 arrowWidth,
                                 arrowHeight);
        const bool useArrowGradient = fontGradientEnabled || backgroundGradientEnabled;
        BYTE arrowAlphaBase = textAlphaValue;
        if (backgroundGradientVisible && gradientAlpha > arrowAlphaBase) {
            arrowAlphaBase = gradientAlpha;
        }
        const int arrowBoost = pressed ? 60 : (hot ? 35 : 15);
        const int boostedAlpha = static_cast<int>(arrowAlphaBase) + arrowBoost;
        const BYTE arrowAlpha = scaleAlpha(static_cast<BYTE>(boostedAlpha > 255 ? 255 : boostedAlpha),
                                           dropdownAlphaMultiplier);
        const Gdiplus::Color arrowStartColor(arrowAlpha, arrowTextStart.GetR(), arrowTextStart.GetG(),
                                             arrowTextStart.GetB());
        const Gdiplus::Color arrowEndColor(arrowAlpha, arrowTextEnd.GetR(), arrowTextEnd.GetG(),
                                           arrowTextEnd.GetB());
        if (useArrowGradient) {
            Gdiplus::LinearGradientBrush arrowBrush(arrowRect, arrowStartColor, arrowEndColor,
                                                    Gdiplus::LinearGradientModeHorizontal);
            arrowBrush.SetGammaCorrection(TRUE);
            graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
        } else {
            const BYTE arrowRed = AverageColorChannel(arrowStartColor.GetR(), arrowEndColor.GetR());
            const BYTE arrowGreen = AverageColorChannel(arrowStartColor.GetG(), arrowEndColor.GetG());
            const BYTE arrowBlue = AverageColorChannel(arrowStartColor.GetB(), arrowEndColor.GetB());
            Gdiplus::SolidBrush arrowBrush(Gdiplus::Color(arrowAlpha, arrowRed, arrowGreen, arrowBlue));
            graphics.FillPolygon(&arrowBrush, arrow, ARRAYSIZE(arrow));
        }
    };

    HTHEME theme = nullptr;
    if (IsAppThemed() && IsThemeActive()) {
        SetLastError(ERROR_SUCCESS);
        theme = OpenThemeData(hwnd, L"BreadcrumbBar");
        if (!theme) {
            const DWORD breadcrumbError = GetLastError();
            if (breadcrumbError != ERROR_SUCCESS) {
                LogLastError(L"OpenThemeData(BreadcrumbBar)", breadcrumbError);
            } else {
                LogMessage(LogLevel::Error,
                           L"OpenThemeData(BreadcrumbBar) returned nullptr without extended error.");
            }

            SetLastError(ERROR_SUCCESS);
            theme = OpenThemeData(hwnd, L"Toolbar");
            if (!theme) {
                const DWORD toolbarError = GetLastError();
                if (toolbarError != ERROR_SUCCESS) {
                    LogLastError(L"OpenThemeData(Toolbar)", toolbarError);
                } else {
                    LogMessage(LogLevel::Error,
                               L"OpenThemeData(Toolbar) returned nullptr without extended error.");
                }
            }
        }
    }
    struct ThemeCloser {
        HTHEME handle;
        ~ThemeCloser() {
            if (handle) {
                CloseThemeData(handle);
            }
        }
    } themeCloser{theme};

    const COLORREF highlightBackgroundColor =
        theme ? GetThemeSysColor(theme, COLOR_HIGHLIGHT) : GetSysColor(COLOR_HIGHLIGHT);

    HFONT fontHandle = reinterpret_cast<HFONT>(SendMessage(hwnd, WM_GETFONT, 0, 0));
    if (!fontHandle) {
        fontHandle = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    Gdiplus::Font font(drawDc, fontHandle);
    if (font.GetLastStatus() != Gdiplus::Ok) {
        if (buffer) {
            EndBufferedPaint(buffer, TRUE);
        }
        EndPaint(hwnd, &ps);
        return true;
    }

    Gdiplus::StringFormat format;
    format.SetAlignment(Gdiplus::StringAlignmentNear);
    format.SetLineAlignment(Gdiplus::StringAlignmentCenter);
    format.SetTrimming(Gdiplus::StringTrimmingEllipsisCharacter);
    format.SetFormatFlags(Gdiplus::StringFormatFlagsNoWrap);

    const int gradientTransparency = std::clamp<int>(m_breadcrumbGradientTransparency, 0, 100);
    const int gradientOpacityPercent = 100 - gradientTransparency;
    const int fontBrightness = std::clamp<int>(m_breadcrumbFontBrightness, 0, 100);
    const BYTE textAlphaBase = 255;

    bool buttonIsPressed = false;
    bool buttonIsHot = false;
    bool buttonHasDropdown = false;
    bool buttonUseFontGradient = false;
    BYTE buttonTextAlpha = 0;
    BYTE buttonScaledAlpha = 0;
    bool buttonBackgroundGradientVisible = false;
    Gdiplus::Color buttonBrightFontEnd;
    Gdiplus::Color buttonTextPaintStart;
    Gdiplus::Color buttonTextPaintEnd;
    RECT buttonRect{};

    static const std::array<COLORREF, 7> kRainbowColors = {
        RGB(255, 59, 48),   // red
        RGB(255, 149, 0),   // orange
        RGB(255, 204, 0),   // yellow
        RGB(52, 199, 89),   // green
        RGB(0, 122, 255),   // blue
        RGB(88, 86, 214),   // indigo
        RGB(175, 82, 222)   // violet
    };

    HIMAGELIST imageList = reinterpret_cast<HIMAGELIST>(SendMessage(hwnd, TB_GETIMAGELIST, 0, 0));
    if (!imageList) {
        imageList = reinterpret_cast<HIMAGELIST>(SendMessage(hwnd, TB_GETIMAGELIST, 1, 0));
    }
    int imageWidth = 0;
    int imageHeight = 0;
    if (imageList) {
        ImageList_GetIconSize(imageList, &imageWidth, &imageHeight);
    }

    auto fetchBreadcrumbText = [&](const TBBUTTON& button) -> std::wstring {
        const UINT commandId = static_cast<UINT>(button.idCommand);

        LRESULT textLength = SendMessage(hwnd, TB_GETBUTTONTEXTW, commandId, 0);
        if (textLength > 0) {
            std::wstring text(static_cast<size_t>(textLength) + 1, L'\0');
            LRESULT copied =
                SendMessage(hwnd, TB_GETBUTTONTEXTW, commandId, reinterpret_cast<LPARAM>(text.data()));
            if (copied > 0) {
                text.resize(static_cast<size_t>(copied));
                return text;
            }
        }

        // Some breadcrumb configurations clear the toolbar's stored text. In those cases, query the
        // button information directly so we can render the gradient text ourselves.
        constexpr size_t kMaxBreadcrumbText = 512;
        std::wstring fallback(kMaxBreadcrumbText, L'\0');
        TBBUTTONINFOW info{};
        info.cbSize = sizeof(info);
        info.dwMask = TBIF_TEXT;
        info.pszText = fallback.data();
        info.cchText = static_cast<int>(fallback.size());
        if (SendMessage(hwnd, TB_GETBUTTONINFOW, commandId, reinterpret_cast<LPARAM>(&info))) {
            fallback.resize(std::wcslen(fallback.c_str()));
            if (!fallback.empty()) {
                return fallback;
            }
        }

        return std::wstring();
    };

    const int buttonCount = static_cast<int>(SendMessage(hwnd, TB_BUTTONCOUNT, 0, 0));
    const LRESULT hotItemIndex = SendMessage(hwnd, TB_GETHOTITEM, 0, 0);
    int gradientStartX = 0;
    int gradientEndX = 0;
    if (m_useCustomBreadcrumbFontColors) {
        RECT toolbarRect{};
        if (GetClientRect(hwnd, &toolbarRect)) {
            gradientStartX = toolbarRect.left;
            gradientEndX = toolbarRect.right;
        }
        int detectedLeft = std::numeric_limits<int>::max();
        int detectedRight = std::numeric_limits<int>::min();
        for (int i = 0; i < buttonCount; ++i) {
            TBBUTTON gradientButton{};
            if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&gradientButton))) {
                continue;
            }
            if ((gradientButton.fsStyle & TBSTYLE_SEP) != 0 || (gradientButton.fsState & TBSTATE_HIDDEN) != 0) {
                continue;
            }
            RECT gradientRect{};
            if (!SendMessage(hwnd, TB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&gradientRect))) {
                continue;
            }
            detectedLeft = std::min(detectedLeft, static_cast<int>(gradientRect.left));
            detectedRight = std::max(detectedRight, static_cast<int>(gradientRect.right));
        }
        if (detectedLeft < detectedRight) {
            gradientStartX = detectedLeft;
            gradientEndX = detectedRight;
        }
    }

    auto sampleFontGradientAtX = [&](int x) -> COLORREF {
        if (!m_useCustomBreadcrumbFontColors) {
            return m_breadcrumbFontGradientStartColor;
        }
        if (gradientEndX <= gradientStartX) {
            return x <= gradientStartX ? m_breadcrumbFontGradientStartColor : m_breadcrumbFontGradientEndColor;
        }
        const int clamped = std::clamp<int>(x, gradientStartX, gradientEndX);
        const double position = static_cast<double>(clamped - gradientStartX) /
                                static_cast<double>(gradientEndX - gradientStartX);
        auto interpolateChannel = [&](BYTE start, BYTE end) -> BYTE {
            const double value = static_cast<double>(start) +
                                 (static_cast<double>(end) - static_cast<double>(start)) * position;
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(value)), 0, 255));
        };
        return RGB(interpolateChannel(GetRValue(m_breadcrumbFontGradientStartColor),
                                      GetRValue(m_breadcrumbFontGradientEndColor)),
                   interpolateChannel(GetGValue(m_breadcrumbFontGradientStartColor),
                                      GetGValue(m_breadcrumbFontGradientEndColor)),
                   interpolateChannel(GetBValue(m_breadcrumbFontGradientStartColor),
                                      GetBValue(m_breadcrumbFontGradientEndColor)));
    };
    int colorIndex = 0;
    for (int i = 0; i < buttonCount; ++i) {
        TBBUTTON button{};
        if (!SendMessage(hwnd, TB_GETBUTTON, i, reinterpret_cast<LPARAM>(&button))) {
            continue;
        }
        if ((button.fsStyle & TBSTYLE_SEP) != 0 || (button.fsState & TBSTATE_HIDDEN) != 0) {
            continue;
        }

        if (!SendMessage(hwnd, TB_GETITEMRECT, i, reinterpret_cast<LPARAM>(&buttonRect))) {
            continue;
        }

        buttonIsPressed = (button.fsState & TBSTATE_PRESSED) != 0;
        buttonIsHot = !buttonIsPressed && ((button.fsState & TBSTATE_HOT) != 0 ||
                                           (hotItemIndex >= 0 && i == static_cast<int>(hotItemIndex)));
        buttonHasDropdown = (button.fsStyle & BTNS_DROPDOWN) != 0;
        const bool hasIcon = imageList && imageWidth > 0 && imageHeight > 0 && button.iBitmap >= 0 &&
                              button.iBitmap != I_IMAGENONE;
        buttonUseFontGradient = m_breadcrumbFontGradientEnabled;

        COLORREF startRgb = 0;
        COLORREF endRgb = 0;
        if (m_useCustomBreadcrumbGradientColors) {
            startRgb = m_breadcrumbGradientStartColor;
            endRgb = m_breadcrumbGradientEndColor;
            ++colorIndex;
        } else {
            const size_t startIndex = static_cast<size_t>(colorIndex % kRainbowColors.size());
            const size_t endIndex = static_cast<size_t>((colorIndex + 1) % kRainbowColors.size());
            ++colorIndex;
            startRgb = kRainbowColors[startIndex];
            endRgb = kRainbowColors[endIndex];
        }

        auto darkenChannel = [](BYTE channel) -> BYTE {
            return static_cast<BYTE>(std::clamp<int>(static_cast<int>(channel) * 35 / 100, 0, 255));
        };
        auto transformBackgroundChannel = [&](BYTE channel) -> BYTE {
            return m_useCustomBreadcrumbGradientColors ? channel : darkenChannel(channel);
        };
        auto applyBrightness = [&](BYTE channel) -> BYTE {
            const int boosted = channel + ((255 - channel) * fontBrightness) / 100;
            return static_cast<BYTE>(std::clamp<int>(boosted, 0, 255));
        };

        Gdiplus::RectF rectF(static_cast<Gdiplus::REAL>(buttonRect.left),
                             static_cast<Gdiplus::REAL>(buttonRect.top),
                             static_cast<Gdiplus::REAL>(buttonRect.right - buttonRect.left),
                             static_cast<Gdiplus::REAL>(buttonRect.bottom - buttonRect.top));

        // Deflate the painted area to leave the native outline pixels untouched. Keep
        // the original buttonRect unmodified so hit-testing continues to use the
        // Explorer-provided bounds.
        constexpr Gdiplus::REAL kBreadcrumbVerticalDeflate = 1.0f;
        const Gdiplus::REAL verticalDeflate =
            std::min<Gdiplus::REAL>(kBreadcrumbVerticalDeflate, rectF.Height / 2.0f);
        if (verticalDeflate > 0.0f) {
            rectF.Y += verticalDeflate;
            rectF.Height = std::max<Gdiplus::REAL>(0.0f, rectF.Height - (verticalDeflate * 2.0f));
        }

        BYTE baseAlpha = 200;
        if (buttonIsPressed) {
            baseAlpha = 235;
        } else if (buttonIsHot) {
            baseAlpha = 220;
        }

        buttonScaledAlpha = static_cast<BYTE>(std::clamp<int>(baseAlpha * gradientOpacityPercent / 100, 0, 255));
        buttonBackgroundGradientVisible = (m_breadcrumbGradientEnabled && buttonScaledAlpha > 0);

        Gdiplus::Color backgroundGradientStartColor;
        Gdiplus::Color backgroundGradientEndColor;
        bool hasBackgroundGradientColors = false;
        Gdiplus::Color backgroundSolidColor;
        bool hasBackgroundSolidColor = false;

        if (buttonBackgroundGradientVisible) {
            backgroundGradientStartColor = BrightenBreadcrumbColor(
                Gdiplus::Color(buttonScaledAlpha, transformBackgroundChannel(GetRValue(startRgb)),
                               transformBackgroundChannel(GetGValue(startRgb)),
                               transformBackgroundChannel(GetBValue(startRgb))),
                buttonIsHot, buttonIsPressed, highlightBackgroundColor);
            backgroundGradientEndColor = BrightenBreadcrumbColor(
                Gdiplus::Color(buttonScaledAlpha, transformBackgroundChannel(GetRValue(endRgb)),
                               transformBackgroundChannel(GetGValue(endRgb)),
                               transformBackgroundChannel(GetBValue(endRgb))),
                buttonIsHot, buttonIsPressed, highlightBackgroundColor);
            hasBackgroundGradientColors = true;
            Gdiplus::LinearGradientBrush backgroundBrush(rectF, backgroundGradientStartColor,
                                                         backgroundGradientEndColor, Gdiplus::LinearGradientModeHorizontal);
            backgroundBrush.SetGammaCorrection(TRUE);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
            graphics.FillRectangle(&backgroundBrush, rectF);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
        } else {
            RECT sampleRect = buttonRect;
            const COLORREF averageBackground =
                SampleAverageColor(drawDc, sampleRect).value_or(GetSysColor(COLOR_WINDOW));
            backgroundSolidColor = Gdiplus::Color(255, GetRValue(averageBackground), GetGValue(averageBackground),
                                                  GetBValue(averageBackground));
            hasBackgroundSolidColor = true;
            if (buttonIsHot || buttonIsPressed) {
                graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                const BYTE overlayAlpha = static_cast<BYTE>(buttonIsPressed ? 140 : 100);
                Gdiplus::Color overlayColor(overlayAlpha, GetRValue(highlightBackgroundColor),
                                            GetGValue(highlightBackgroundColor), GetBValue(highlightBackgroundColor));
                Gdiplus::SolidBrush overlayBrush(overlayColor);
                graphics.FillRectangle(&overlayBrush, rectF);
                auto blendOverlayChannel = [&](BYTE base, BYTE overlay) -> BYTE {
                    const int blended = static_cast<int>(std::lround(base + (overlay - base) * (overlayAlpha / 255.0)));
                    return static_cast<BYTE>(std::clamp<int>(blended, 0, 255));
                };
                backgroundSolidColor = Gdiplus::Color(
                    255, blendOverlayChannel(backgroundSolidColor.GetR(), overlayColor.GetR()),
                    blendOverlayChannel(backgroundSolidColor.GetG(), overlayColor.GetG()),
                    blendOverlayChannel(backgroundSolidColor.GetB(), overlayColor.GetB()));
            }
        }

        if (hasIcon) {
            const int iconX = buttonRect.left + 4;
            const LONG verticalSpace = ((buttonRect.bottom - buttonRect.top) - imageHeight) / 2;
            const LONG iconYOffset = std::max<LONG>(0, verticalSpace);
            const int iconY = static_cast<int>(buttonRect.top + iconYOffset);
            ImageList_Draw(imageList, button.iBitmap, drawDc, iconX, iconY, ILD_TRANSPARENT);
        }

        buttonTextAlpha = textAlphaBase;
        if (buttonIsPressed) {
            buttonTextAlpha = static_cast<BYTE>(std::min<int>(255, buttonTextAlpha + 60));
        } else if (buttonIsHot) {
            buttonTextAlpha = static_cast<BYTE>(std::min<int>(255, buttonTextAlpha + 35));
        }
        COLORREF buttonFontStartRgb = startRgb;
        COLORREF buttonFontEndRgb = endRgb;
        if (m_useCustomBreadcrumbFontColors) {
            buttonFontStartRgb = sampleFontGradientAtX(buttonRect.left);
            buttonFontEndRgb = sampleFontGradientAtX(buttonRect.right);
        }

        auto computeBrightFontColorFn = [&](COLORREF rgb) -> Gdiplus::Color {
            const BYTE adjustedRed = applyBrightness(GetRValue(rgb));
            const BYTE adjustedGreen = applyBrightness(GetGValue(rgb));
            const BYTE adjustedBlue = applyBrightness(GetBValue(rgb));
            return BrightenBreadcrumbColor(
                Gdiplus::Color(buttonTextAlpha, adjustedRed, adjustedGreen, adjustedBlue), buttonIsHot,
                buttonIsPressed, highlightBackgroundColor);
        };

        const Gdiplus::Color buttonBrightFontStart = computeBrightFontColorFn(buttonFontStartRgb);
        buttonBrightFontEnd = computeBrightFontColorFn(buttonFontEndRgb);

        auto computeOpaqueFontColorFn = [&](const Gdiplus::Color& fontColor, bool useStart) {
            if (buttonTextAlpha >= 255) {
                return Gdiplus::Color(255, fontColor.GetR(), fontColor.GetG(), fontColor.GetB());
            }
            const double opacity = static_cast<double>(buttonTextAlpha) / 255.0;
            int backgroundRed = 0;
            int backgroundGreen = 0;
            int backgroundBlue = 0;
            if (hasBackgroundGradientColors) {
                const Gdiplus::Color& background = useStart ? backgroundGradientStartColor : backgroundGradientEndColor;
                backgroundRed = background.GetR();
                backgroundGreen = background.GetG();
                backgroundBlue = background.GetB();
            } else if (hasBackgroundSolidColor) {
                backgroundRed = backgroundSolidColor.GetR();
                backgroundGreen = backgroundSolidColor.GetG();
                backgroundBlue = backgroundSolidColor.GetB();
            } else {
                backgroundRed = GetRValue(highlightBackgroundColor);
                backgroundGreen = GetGValue(highlightBackgroundColor);
                backgroundBlue = GetBValue(highlightBackgroundColor);
            }
            auto blendComponent = [&](int foreground, int background) -> BYTE {
                const double value = static_cast<double>(background) +
                                     (static_cast<double>(foreground) - static_cast<double>(background)) * opacity;
                return static_cast<BYTE>(std::clamp<int>(static_cast<int>(std::lround(value)), 0, 255));
            };
            return Gdiplus::Color(255, blendComponent(fontColor.GetR(), backgroundRed),
                                  blendComponent(fontColor.GetG(), backgroundGreen),
                                  blendComponent(fontColor.GetB(), backgroundBlue));
        };

        buttonTextPaintStart = computeOpaqueFontColorFn(buttonBrightFontStart, true);
        buttonTextPaintEnd = computeOpaqueFontColorFn(buttonBrightFontEnd, false);

        constexpr int kTextPadding = 8;
        const int iconReserve = hasIcon ? (imageWidth + 6) : 0;
        const int dropdownReserve = buttonHasDropdown ? 12 : 0;

        std::wstring text = fetchBreadcrumbText(button);
        if (!text.empty()) {
            const int iconAreaLeft = buttonRect.left + iconReserve;
            const int textBaseLeft = iconAreaLeft + kTextPadding;
            RECT textRect = buttonRect;
            textRect.left = std::max(iconAreaLeft, textBaseLeft - 1);
            textRect.right -= kTextPadding;
            if (buttonHasDropdown) {
                textRect.right -= dropdownReserve;
            }

            if (textRect.right <= textRect.left) {
                textRect.left = iconAreaLeft;
                textRect.right = buttonRect.right - (buttonHasDropdown ? dropdownReserve : 0);
            }

            if (textRect.right > textRect.left) {
                Gdiplus::RectF textRectF(static_cast<Gdiplus::REAL>(textRect.left),
                                         static_cast<Gdiplus::REAL>(textRect.top),
                                         static_cast<Gdiplus::REAL>(textRect.right - textRect.left),
                                         static_cast<Gdiplus::REAL>(textRect.bottom - textRect.top));

                COLORREF textFontStartRgb = buttonFontStartRgb;
                COLORREF textFontEndRgb = buttonFontEndRgb;
                if (m_useCustomBreadcrumbFontColors) {
                    textFontStartRgb = sampleFontGradientAtX(textRect.left);
                    textFontEndRgb = sampleFontGradientAtX(textRect.right);
                }
                const Gdiplus::Color brightFontStart = computeBrightFontColorFn(textFontStartRgb);
                const Gdiplus::Color textBrightFontEnd = computeBrightFontColorFn(textFontEndRgb);
                buttonTextPaintStart = computeOpaqueFontColorFn(brightFontStart, true);
                buttonTextPaintEnd = computeOpaqueFontColorFn(textBrightFontEnd, false);

                if (buttonTextAlpha > 0) {
                    if (buttonUseFontGradient) {
                        const auto previousHint = graphics.GetTextRenderingHint();
                        const auto previousMode = graphics.GetCompositingMode();
                        const auto previousPixelOffset = graphics.GetPixelOffsetMode();
                        const auto previousSmoothing = graphics.GetSmoothingMode();

                        graphics.SetTextRenderingHint(Gdiplus::TextRenderingHintAntiAliasGridFit);
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                        graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
                        graphics.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);

                        Gdiplus::LinearGradientBrush textBrush(
                            textRectF,
                            buttonTextPaintStart,
                            buttonTextPaintEnd,
                            Gdiplus::LinearGradientModeHorizontal);
                        textBrush.SetGammaCorrection(TRUE);

                        bool renderedWithPath = false;
                        Gdiplus::FontFamily fontFamily;
                        if (font.GetFamily(&fontFamily) == Gdiplus::Ok) {
                            Gdiplus::GraphicsPath textPath;
                            if (textPath.AddString(text.c_str(), static_cast<INT>(text.size()), &fontFamily,
                                                   font.GetStyle(), font.GetSize(), textRectF, &format) ==
                                Gdiplus::Ok) {
                                graphics.FillPath(&textBrush, &textPath);
                                renderedWithPath = true;
                            }
                        }

                        if (!renderedWithPath) {
                            graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                                &textBrush);
                        }

                        graphics.SetSmoothingMode(previousSmoothing);
                        graphics.SetPixelOffsetMode(previousPixelOffset);
                        graphics.SetCompositingMode(previousMode);
                        graphics.SetTextRenderingHint(previousHint);
                    } else {
                        graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
                        const BYTE avgRed = AverageColorChannel(brightFontStart.GetR(), textBrightFontEnd.GetR());
                        const BYTE avgGreen = AverageColorChannel(brightFontStart.GetG(), textBrightFontEnd.GetG());
                        const BYTE avgBlue = AverageColorChannel(brightFontStart.GetB(), textBrightFontEnd.GetB());
                        Gdiplus::Color solidColor = computeOpaqueFontColorFn(
                            Gdiplus::Color(buttonTextAlpha, avgRed, avgGreen, avgBlue), true);
                        Gdiplus::SolidBrush textBrush(solidColor);
                        graphics.DrawString(text.c_str(), static_cast<INT>(text.size()), &font, textRectF, &format,
                                            &textBrush);
                    }
                    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                }
            }
        }

        Gdiplus::Color arrowTextStart = buttonTextPaintStart;
        Gdiplus::Color arrowTextEnd = buttonTextPaintEnd;
        if (buttonHasDropdown && m_useCustomBreadcrumbFontColors) {
            const int arrowLeft = buttonRect.right - 12;
            const int arrowRight = buttonRect.right - 6;
            auto computeArrowTextColor = [&](int sampleX, bool useStart) {
                const COLORREF sampleRgb = sampleFontGradientAtX(sampleX);
                const Gdiplus::Color brightColor = computeBrightFontColorFn(sampleRgb);
                return computeOpaqueFontColorFn(brightColor, useStart);
            };
            arrowTextStart = computeArrowTextColor(arrowLeft, true);
            arrowTextEnd = computeArrowTextColor(arrowRight, false);
        }

        if (buttonHasDropdown) {
            drawDropdownArrow(buttonRect, buttonIsHot, buttonIsPressed, buttonTextAlpha, buttonBrightFontEnd,
                              arrowTextStart, arrowTextEnd, buttonUseFontGradient, m_breadcrumbGradientEnabled,
                              buttonBackgroundGradientVisible, buttonScaledAlpha, highlightBackgroundColor);
        }
    }

    if (buffer) {
        EndBufferedPaint(buffer, TRUE);
    }
    EndPaint(hwnd, &ps);
    return true;
}

bool CExplorerBHO::HandleProgressPaint(HWND hwnd) {
    if (!m_useCustomProgressGradientColors) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc) {
        return false;
    }

    RECT client{};
    if (!GetClientRect(hwnd, &client)) {
        EndPaint(hwnd, &ps);
        return true;
    }

    RECT inner = client;
    DrawEdge(dc, &inner, EDGE_SUNKEN, BF_RECT | BF_ADJUST);

    FillRect(dc, &inner, GetSysColorBrush(COLOR_WINDOW));

    PBRANGE range{};
    SendMessageW(hwnd, PBM_GETRANGE, TRUE, reinterpret_cast<LPARAM>(&range));
    if (range.iHigh <= range.iLow) {
        range.iLow = 0;
        range.iHigh = 100;
    }

    LRESULT position = SendMessageW(hwnd, PBM_GETPOS, 0, 0);
    const int span = range.iHigh - range.iLow;
    double fraction = 0.0;
    if (span > 0) {
        fraction = static_cast<double>(position - range.iLow) / static_cast<double>(span);
    }
    fraction = std::clamp<double>(fraction, 0.0, 1.0);

    const LONG width = inner.right - inner.left;
    if (fraction > 0.0 && width > 0) {
        const LONG progressWidth = static_cast<LONG>(std::lround(fraction * static_cast<double>(width)));
        if (progressWidth > 0) {
            RECT fillRect = inner;
            fillRect.right = std::min(fillRect.left + progressWidth, inner.right);
            const LONG fillWidth = fillRect.right - fillRect.left;
            const LONG fillHeight = fillRect.bottom - fillRect.top;
            if (fillWidth > 0 && fillHeight > 0) {
                bool rendered = false;
                if (EnsureProgressGradientResources() && m_progressGradientBits &&
                    m_progressGradientInfo.bmiHeader.biWidth > 0) {
                    const int previousMode = SetStretchBltMode(dc, HALFTONE);
                    POINT origin{};
                    if (previousMode != 0) {
                        SetBrushOrgEx(dc, 0, 0, &origin);
                    }
                    const int srcWidth = m_progressGradientInfo.bmiHeader.biWidth;
                    const int srcHeight = (m_progressGradientInfo.bmiHeader.biHeight < 0)
                                              ? -m_progressGradientInfo.bmiHeader.biHeight
                                              : m_progressGradientInfo.bmiHeader.biHeight;
                    const int result = StretchDIBits(dc, fillRect.left, fillRect.top, fillWidth, fillHeight, 0, 0,
                                                     srcWidth, srcHeight, m_progressGradientBits,
                                                     &m_progressGradientInfo, DIB_RGB_COLORS, SRCCOPY);
                    if (previousMode != 0) {
                        SetBrushOrgEx(dc, origin.x, origin.y, nullptr);
                        SetStretchBltMode(dc, previousMode);
                    }
                    rendered = (result != GDI_ERROR);
                }

                if (!rendered) {
                    TRIVERTEX vertex[2] = {};
                    vertex[0].x = fillRect.left;
                    vertex[0].y = fillRect.top;
                    vertex[0].Red = static_cast<COLOR16>(GetRValue(m_progressGradientStartColor) << 8);
                    vertex[0].Green = static_cast<COLOR16>(GetGValue(m_progressGradientStartColor) << 8);
                    vertex[0].Blue = static_cast<COLOR16>(GetBValue(m_progressGradientStartColor) << 8);
                    vertex[0].Alpha = 0xFFFF;

                    vertex[1].x = fillRect.right;
                    vertex[1].y = fillRect.bottom;
                    vertex[1].Red = static_cast<COLOR16>(GetRValue(m_progressGradientEndColor) << 8);
                    vertex[1].Green = static_cast<COLOR16>(GetGValue(m_progressGradientEndColor) << 8);
                    vertex[1].Blue = static_cast<COLOR16>(GetBValue(m_progressGradientEndColor) << 8);
                    vertex[1].Alpha = 0xFFFF;

                    GRADIENT_RECT gradientRect{0, 1};
                    GradientFill(dc, vertex, 2, &gradientRect, 1, GRADIENT_FILL_RECT_H);
                }
            }
        }
    }

    EndPaint(hwnd, &ps);
    return true;
}

bool CExplorerBHO::HandleAddressEditPaint(HWND hwnd) {
    if (!m_breadcrumbFontGradientEnabled || !m_gdiplusInitialized) {
        return false;
    }

    PAINTSTRUCT ps{};
    HDC dc = BeginPaint(hwnd, &ps);
    if (!dc) {
        return false;
    }

    DrawAddressEditContent(hwnd, dc);

    EndPaint(hwnd, &ps);
    return true;
}

bool CExplorerBHO::DrawAddressEditContent(HWND hwnd, HDC dc) {
    if (!hwnd || !dc) {
        return false;
    }

    RECT client{};
    if (!GetClientRect(hwnd, &client)) {
        return true;
    }

    const UINT windowDpi = GetWindowDpi(hwnd);
    const UINT initialDcDpiX = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSX)));
    const UINT initialDcDpiY = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSY)));
    const bool dpiMismatch = (windowDpi != 0) &&
                             (initialDcDpiX != windowDpi || initialDcDpiY != windowDpi);

#ifdef DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    const ScopedThreadDpiAwarenessContext dpiScope(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, dpiMismatch);
#else
    const ScopedThreadDpiAwarenessContext dpiScope(nullptr, false);
#endif

    const BOOL caretHidden = HideCaret(hwnd);

    LRESULT eraseResult = SendMessageW(hwnd, WM_ERASEBKGND, reinterpret_cast<WPARAM>(dc), 0);
    if (eraseResult == 0) {
        HBRUSH backgroundBrush = reinterpret_cast<HBRUSH>(GetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND));
        if (!backgroundBrush) {
            backgroundBrush = GetSysColorBrush(COLOR_WINDOW);
        }
        FillRect(dc, &client, backgroundBrush);
    }

    std::wstring text;
    const int length = GetWindowTextLengthW(hwnd);
    if (length > 0) {
        text.resize(static_cast<size_t>(length) + 1);
        int copied = GetWindowTextW(hwnd, text.data(), length + 1);
        if (copied < 0) {
            copied = 0;
        }
        text.resize(static_cast<size_t>(std::clamp<int>(copied, 0, length)));
    }

    RECT formatRect = client;
    SendMessageW(hwnd, EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formatRect));
    if (formatRect.right <= formatRect.left) {
        formatRect = client;
    }

    const UINT effectiveDcDpiX = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSX)));
    const UINT effectiveDcDpiY = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSY)));
    const double scaleX = (windowDpi != 0) ? static_cast<double>(effectiveDcDpiX) /
                                                static_cast<double>(windowDpi)
                                          : 1.0;
    const double scaleY = (windowDpi != 0) ? static_cast<double>(effectiveDcDpiY) /
                                                static_cast<double>(windowDpi)
                                          : 1.0;

    auto scaleCoordinateX = [&](int value) -> double {
        return static_cast<double>(value) * scaleX;
    };
    auto scaleCoordinateY = [&](int value) -> double {
        return static_cast<double>(value) * scaleY;
    };

    auto scaleRectToDevice = [&](const RECT& rect) -> RECT {
        const double left = scaleCoordinateX(rect.left);
        const double right = scaleCoordinateX(rect.right);
        const double top = scaleCoordinateY(rect.top);
        const double bottom = scaleCoordinateY(rect.bottom);

        RECT result{};
        result.left = static_cast<LONG>(std::lround(std::min(left, right)));
        result.right = static_cast<LONG>(std::lround(std::max(left, right)));
        result.top = static_cast<LONG>(std::lround(std::min(top, bottom)));
        result.bottom = static_cast<LONG>(std::lround(std::max(top, bottom)));
        return result;
    };

    RECT scaledFormatRect = scaleRectToDevice(formatRect);

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    if (!font) {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    HFONT oldFont = nullptr;
    if (font) {
        oldFont = static_cast<HFONT>(SelectObject(dc, font));
    }

    const auto& gradientConfig = m_glowCoordinator.BreadcrumbFontGradient();
    const BreadcrumbGradientPalette gradientPalette = ResolveBreadcrumbGradientPalette(gradientConfig);

    const COLORREF previousTextColor = GetTextColor(dc);
    const int previousBkMode = SetBkMode(dc, TRANSPARENT);

    struct CharacterMetrics {
        int index;
        double x;
        double y;
        double width;
    };

    std::vector<CharacterMetrics> characters;
    characters.reserve(text.size());

    double gradientLeft = static_cast<double>(scaledFormatRect.right);
    double gradientRight = static_cast<double>(scaledFormatRect.left);

    for (int i = 0; i < static_cast<int>(text.size()); ++i) {
        LRESULT pos = SendMessageW(hwnd, EM_POSFROMCHAR, i, 0);
        if (pos == -1) {
            continue;
        }

        const int rawCharX = static_cast<SHORT>(LOWORD(static_cast<DWORD_PTR>(pos)));
        const int rawCharY = static_cast<SHORT>(HIWORD(static_cast<DWORD_PTR>(pos)));
        const double charX = scaleCoordinateX(rawCharX);
        const double charY = scaleCoordinateY(rawCharY);

        LRESULT nextPos = SendMessageW(hwnd, EM_POSFROMCHAR, i + 1, 0);
        double nextX = charX;
        if (nextPos != -1) {
            const int rawNextX = static_cast<SHORT>(LOWORD(static_cast<DWORD_PTR>(nextPos)));
            nextX = scaleCoordinateX(rawNextX);
        }
        if (nextPos == -1 || nextX <= charX) {
            SIZE extent{};
            if (GetTextExtentPoint32W(dc, &text[i], 1, &extent) && extent.cx > 0) {
                nextX = charX + static_cast<double>(extent.cx);
            } else {
                nextX = charX + 1.0;
            }
        }

        const double charWidth = std::max(1.0, nextX - charX);
        characters.push_back(CharacterMetrics{i, charX, charY, charWidth});

        gradientLeft = std::min(gradientLeft, charX);
        gradientRight = std::max(gradientRight, charX + charWidth);
    }

    if (!characters.empty()) {
        gradientLeft = std::min(gradientLeft, characters.front().x - characters.front().width * 0.5);
        gradientRight = std::max(gradientRight, characters.back().x + characters.back().width * 0.5);
    }

    const double scaledLeftBound = static_cast<double>(scaledFormatRect.left);
    const double scaledRightBound = static_cast<double>(scaledFormatRect.right);

    gradientLeft = std::clamp<double>(gradientLeft, scaledLeftBound, scaledRightBound);
    gradientRight = std::clamp<double>(gradientRight, scaledLeftBound, scaledRightBound);

    if (gradientRight <= gradientLeft) {
        gradientLeft = scaledLeftBound;
        gradientRight = std::max(scaledLeftBound + 1.0, scaledRightBound);
    }

    const double gradientWidth = std::max(1.0, gradientRight - gradientLeft);

    DWORD selectionStart = 0;
    DWORD selectionEnd = 0;
    if (SendMessageW(hwnd, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart),
                     reinterpret_cast<LPARAM>(&selectionEnd)) == 0) {
        selectionStart = selectionEnd = 0;
    }
    if (selectionEnd < selectionStart) {
        std::swap(selectionStart, selectionEnd);
    }
    const bool hasSelection = (selectionEnd > selectionStart);
    const COLORREF highlightTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT);

    RECT clipRect = scaledFormatRect;

    if (hasSelection) {
        const COLORREF highlightColor = GetSysColor(COLOR_HIGHLIGHT);
        HBRUSH selectionBrush = CreateSolidBrush(highlightColor);
        bool deleteSelectionBrush = true;
        if (!selectionBrush) {
            selectionBrush = GetSysColorBrush(COLOR_HIGHLIGHT);
            deleteSelectionBrush = false;
        }

        TEXTMETRICW metrics{};
        int lineHeight = 0;
        if (GetTextMetricsW(dc, &metrics)) {
            lineHeight = metrics.tmHeight + metrics.tmExternalLeading;
        }
        if (lineHeight <= 0) {
            lineHeight = std::max<LONG>(1, scaledFormatRect.bottom - scaledFormatRect.top);
        }

        auto paintSelectionRun = [&](double startX, double endX, int top, int bottom) {
            if (!selectionBrush) {
                return;
            }
            RECT selectionRect{};
            selectionRect.left = static_cast<LONG>(std::floor(startX));
            selectionRect.right = static_cast<LONG>(std::ceil(endX));
            selectionRect.top = top;
            selectionRect.bottom = bottom;

            RECT intersected{};
            if (IntersectRect(&intersected, &selectionRect, &clipRect)) {
                FillRect(dc, &intersected, selectionBrush);
            }
        };

        bool runActive = false;
        double runStartX = 0.0;
        double runEndX = 0.0;
        int runTop = clipRect.top;
        int runBottom = clipRect.top;

        auto flushRun = [&]() {
            if (!runActive) {
                return;
            }
            paintSelectionRun(runStartX, runEndX, runTop, runBottom);
            runActive = false;
        };

        const LONG clipTop = clipRect.top;
        const LONG clipBottom = clipRect.bottom;

        for (const CharacterMetrics& character : characters) {
            const bool isSelected = character.index >= static_cast<int>(selectionStart) &&
                                   character.index < static_cast<int>(selectionEnd);
            if (!isSelected) {
                flushRun();
                continue;
            }

            const double charLeft = character.x;
            const double charRight = character.x + character.width;
            int charTop = static_cast<int>(std::floor(character.y));
            int charBottom = charTop + lineHeight;
            charTop = static_cast<int>(
                std::clamp<LONG>(static_cast<LONG>(charTop), clipTop, clipBottom));
            charBottom = static_cast<int>(
                std::clamp<LONG>(static_cast<LONG>(charBottom), clipTop, clipBottom));
            if (charBottom <= charTop) {
                charBottom = std::min<int>(static_cast<int>(clipBottom), charTop + lineHeight);
            }
            if (charBottom <= charTop) {
                charBottom = std::min<int>(static_cast<int>(clipBottom), charTop + 1);
            }

            if (!runActive) {
                runActive = true;
                runStartX = charLeft;
                runTop = charTop;
                runBottom = charBottom;
            }

            runEndX = charRight;
            runTop = std::min(runTop, charTop);
            runBottom = std::max(runBottom, charBottom);
        }

        flushRun();

        if (deleteSelectionBrush && selectionBrush) {
            DeleteObject(selectionBrush);
        }
    }

    for (const CharacterMetrics& character : characters) {
        bool isSelected = hasSelection && character.index >= static_cast<int>(selectionStart) &&
                          character.index < static_cast<int>(selectionEnd);

        if (isSelected) {
            SetTextColor(dc, highlightTextColor);
        } else {
            const double centerX = character.x + character.width / 2.0;
            double position = (centerX - gradientLeft) / gradientWidth;
            position = std::clamp<double>(position, 0.0, 1.0);

            const COLORREF gradientColor = EvaluateBreadcrumbGradientColor(gradientPalette, position);
            SetTextColor(dc, gradientColor);
        }

        const int drawX = static_cast<int>(std::lround(character.x));
        const int drawY = static_cast<int>(std::lround(character.y));

        ExtTextOutW(dc, drawX, drawY, ETO_CLIPPED, &clipRect, text.data() + character.index, 1, nullptr);
    }

    SetBkMode(dc, previousBkMode);
    SetTextColor(dc, previousTextColor);

    if (oldFont) {
        SelectObject(dc, oldFont);
    }

    if (caretHidden) {
        ShowCaret(hwnd);
    }

    return true;
}

LRESULT CALLBACK CExplorerBHO::BreadcrumbCbtProc(int code, WPARAM wParam, LPARAM lParam) {
    HHOOK hookHandle = nullptr;

    if (code == HCBT_CREATEWND) {
        HWND hwnd = reinterpret_cast<HWND>(wParam);
        const CBT_CREATEWNDW* create = reinterpret_cast<CBT_CREATEWNDW*>(lParam);

        const wchar_t* className = nullptr;
        if (create && create->lpcs) {
            if (HIWORD(create->lpcs->lpszClass)) {
                className = create->lpcs->lpszClass;
            }
        }

        wchar_t classBuffer[64]{};
        if (!className) {
            if (GetClassNameW(hwnd, classBuffer, ARRAYSIZE(classBuffer)) > 0) {
                className = classBuffer;
            }
        }

        if (!className) {
            return CallNextHookEx(nullptr, code, wParam, lParam);
        }

        const bool isToolbar = (_wcsicmp(className, TOOLBARCLASSNAMEW) == 0);
        const bool isCombo = (_wcsicmp(className, L"ComboBoxEx32") == 0);
        const bool isEdit = (_wcsicmp(className, L"Edit") == 0);
        if (!isToolbar && !isCombo && !isEdit) {
            return CallNextHookEx(nullptr, code, wParam, lParam);
        }

        std::vector<CExplorerBHO*> observers;
        {
            std::lock_guard<std::mutex> lock(g_breadcrumbHookMutex);
            auto it = g_breadcrumbHooks.find(GetCurrentThreadId());
            if (it != g_breadcrumbHooks.end()) {
                observers = it->second.observers;
                hookHandle = it->second.hook;
            }
        }

        if (!observers.empty()) {
            for (CExplorerBHO* observer : observers) {
                if (!observer || !observer->m_gdiplusInitialized) {
                    continue;
                }

                if (isToolbar) {
                    if (!observer->m_breadcrumbGradientEnabled && !observer->m_breadcrumbFontGradientEnabled) {
                        continue;
                    }
                    HWND start = hwnd;
                    if (create && create->lpcs && create->lpcs->hwndParent) {
                        start = create->lpcs->hwndParent;
                    }

                    if (!observer->IsBreadcrumbToolbarAncestor(start)) {
                        continue;
                    }
                    if (!observer->IsWindowOwnedByThisExplorer(hwnd)) {
                        continue;
                    }

                    if (observer->InstallBreadcrumbSubclass(hwnd)) {
                        observer->LogBreadcrumbStage(BreadcrumbDiscoveryStage::Discovered,
                                                     L"Breadcrumb toolbar subclassed via CBT hook (hwnd=%p)", hwnd);
                    }
                    continue;
                }

                if (!observer->m_breadcrumbFontGradientEnabled) {
                    continue;
                }

                HWND ancestryCheck = hwnd;
                if (create && create->lpcs && create->lpcs->hwndParent) {
                    ancestryCheck = create->lpcs->hwndParent;
                }
                if (!observer->IsBreadcrumbToolbarAncestor(ancestryCheck) &&
                    !observer->IsBreadcrumbToolbarAncestor(hwnd)) {
                    continue;
                }
                if (!observer->IsWindowOwnedByThisExplorer(hwnd)) {
                    continue;
                }

                observer->UpdateAddressEditSubclass();
            }
        }
    }

    return CallNextHookEx(hookHandle, code, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::BreadcrumbSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                      UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_PAINT:
            if (self->HandleBreadcrumbPaint(hwnd)) {
                return 0;
            }
            break;
        case WM_ERASEBKGND:
            if (self->m_breadcrumbGradientEnabled) {
                return 1;
            }
            break;
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_UPDATEUISTATE:
            if (self->m_breadcrumbGradientEnabled || self->m_breadcrumbFontGradientEnabled) {
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        case WM_NCDESTROY:
            self->RemoveBreadcrumbSubclass();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::ProgressSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                    UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_PAINT:
            if (self->HandleProgressPaint(hwnd)) {
                return 0;
            }
            break;
        case WM_ERASEBKGND:
            if (self->m_useCustomProgressGradientColors) {
                return 1;
            }
            break;
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
            if (self->m_useCustomProgressGradientColors) {
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            break;
        case WM_NCDESTROY:
            self->RemoveProgressSubclass();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::AddressEditSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                       UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    auto requestOnStateChange = [&](bool updateText, bool updateSelection, bool updateFocus, bool updateTheme) {
        bool changed = self->RefreshAddressEditState(hwnd, updateText, updateSelection, updateFocus, updateTheme);
        if (self->m_breadcrumbFontGradientEnabled && changed) {
            self->RequestAddressEditRedraw(hwnd);
        }
    };

    switch (msg) {
        case WM_PAINT: {
            if (self->m_addressEditRedrawTimerActive) {
                KillTimer(hwnd, kAddressEditRedrawTimerId);
                self->m_addressEditRedrawTimerActive = false;
            }
            self->m_addressEditRedrawPending = false;
            if (self->HandleAddressEditPaint(hwnd)) {
                return 0;
            }
            break;
        }
        case WM_PRINTCLIENT: {
            if (self->m_breadcrumbFontGradientEnabled) {
                HDC dc = reinterpret_cast<HDC>(wParam);
                if (self->DrawAddressEditContent(hwnd, dc)) {
                    return 0;
                }
            }
            break;
        }
        case WM_TIMER:
            if (wParam == kAddressEditRedrawTimerId) {
                KillTimer(hwnd, kAddressEditRedrawTimerId);
                self->m_addressEditRedrawTimerActive = false;
                InvalidateRect(hwnd, nullptr, FALSE);
                return 0;
            }
            break;
        case WM_SETTEXT:
        case EM_REPLACESEL:
        case WM_CUT:
        case WM_PASTE:
        case WM_UNDO:
        case WM_CLEAR: {
            LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            requestOnStateChange(true, true, false, false);
            return result;
        }
        case EM_SETSEL: {
            LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            requestOnStateChange(false, true, false, false);
            return result;
        }
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_CHAR:
        case WM_KEYDOWN:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_MBUTTONDOWN:
        case WM_MBUTTONUP:
        case WM_MOUSEMOVE: {
            LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);

            switch (msg) {
                case WM_SETTEXT:
                case EM_REPLACESEL:
                case WM_CUT:
                case WM_PASTE:
                case WM_UNDO:
                case WM_CLEAR:
                    requestOnStateChange(true, true, false, false);
                    break;
                case WM_THEMECHANGED:
                case WM_SETTINGCHANGE:
                    requestOnStateChange(false, false, false, true);
                    break;
                case WM_SETFOCUS:
                case WM_KILLFOCUS:
                    requestOnStateChange(false, false, true, false);
                    break;
                case WM_CHAR:
                    requestOnStateChange(true, true, false, false);
                    break;
                case WM_KEYDOWN: {
                    bool updateText = (wParam == VK_BACK || wParam == VK_DELETE);
                    requestOnStateChange(updateText, true, false, false);
                    break;
                }
                case WM_LBUTTONDOWN:
                case WM_LBUTTONUP:
                case WM_LBUTTONDBLCLK:
                case WM_RBUTTONDOWN:
                case WM_RBUTTONUP:
                case WM_MBUTTONDOWN:
                case WM_MBUTTONUP:
                    requestOnStateChange(false, true, false, false);
                    break;
                case WM_MOUSEMOVE:
                    if ((wParam & (MK_LBUTTON | MK_MBUTTON | MK_RBUTTON)) != 0) {
                        requestOnStateChange(false, true, false, false);
                    }
                    break;
                case EM_SETSEL:
                    requestOnStateChange(false, true, false, false);
                    break;
                default:
                    break;
            }

            return result;
        }
        case WM_SETFONT: {
            LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            bool fontChanged = self->RefreshAddressEditFont(hwnd);
            if (self->m_breadcrumbFontGradientEnabled && fontChanged) {
                self->RequestAddressEditRedraw(hwnd);
            }
            return result;
        }
        case WM_NCDESTROY:
            self->RemoveAddressEditSubclass();
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::ExplorerViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                       UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    LRESULT result = 0;
    if (self->HandleExplorerViewMessage(hwnd, msg, wParam, lParam, &result)) {
        return result;
    }

    if (msg == WM_NCDESTROY) {
        if (hwnd == self->m_listView) {
            self->m_listView = nullptr;
            self->m_listViewSubclassInstalled = false;
        } else if (hwnd == self->m_listViewControlWindow) {
            self->m_listViewControlWindow = nullptr;
            self->m_listViewControl.reset();
            if (self->m_nativeListView && IsWindow(self->m_nativeListView)) {
                EnableWindow(self->m_nativeListView, TRUE);
                ShowWindow(self->m_nativeListView, SW_SHOW);
            }
            self->m_nativeListView = nullptr;
        } else if (hwnd == self->m_directUiView) {
            self->m_directUiView = nullptr;
            self->m_directUiSubclassInstalled = false;
        } else if (hwnd == self->m_treeView) {
            self->m_treeView = nullptr;
            self->m_treeViewSubclassInstalled = false;
            self->m_paneHooks.SetTreeView(nullptr);
        } else if (hwnd == self->m_frameWindow) {
            self->m_frameWindow = nullptr;
            self->m_frameSubclassInstalled = false;
        } else if (hwnd == self->m_shellViewWindow) {
            self->m_shellViewWindowSubclassInstalled = false;
            self->m_shellViewWindow = nullptr;
        } else {
            self->m_listViewHostSubclassed.erase(hwnd);
        }

        if (!self->m_listView && !self->m_treeView && !self->m_shellViewWindow) {
            self->m_shellView.Reset();
            self->ClearPendingOpenInNewTabState();
        }

        self->UnregisterGlowSurface(hwnd);
        RemoveWindowSubclass(hwnd, &CExplorerBHO::ExplorerViewSubclassProc, reinterpret_cast<UINT_PTR>(self));
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::ScrollbarGlowSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                         UINT_PTR subclassId, DWORD_PTR) {
    auto* self = reinterpret_cast<CExplorerBHO*>(subclassId);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_NCPAINT:
        case WM_PRINTCLIENT: {
            if (!self->ShouldSuppressScrollbarDrawing(hwnd)) {
                self->RestoreScrollbarTransparency(hwnd);
                break;
            }

            self->EnsureScrollbarTransparency(hwnd);

            bool painted = false;
            if (msg == WM_PRINTCLIENT) {
                painted = self->PaintScrollbarGlow(hwnd, reinterpret_cast<HDC>(wParam), nullptr);
            } else {
                painted = self->PaintScrollbarGlow(hwnd, nullptr, reinterpret_cast<HRGN>(wParam));
            }

            if (!painted) {
                self->RestoreScrollbarTransparency(hwnd);
                break;
            }

            return 0;
        }
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
        case WM_DPICHANGED: {
            LRESULT result = DefSubclassProc(hwnd, msg, wParam, lParam);
            if (self->ShouldSuppressScrollbarDrawing(hwnd)) {
                self->EnsureScrollbarTransparency(hwnd);
                self->RequestScrollbarGlowRepaint(hwnd);
            } else {
                self->RestoreScrollbarTransparency(hwnd);
            }
            return result;
        }
        case WM_NCDESTROY:
            self->RestoreScrollbarTransparency(hwnd);
            self->m_scrollbarGlowSubclassed.erase(hwnd);
            RemoveWindowSubclass(hwnd, &CExplorerBHO::ScrollbarGlowSubclassProc, reinterpret_cast<UINT_PTR>(self));
            break;
        default:
            break;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK CExplorerBHO::StatusBarSubclassProc(
    HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData) {
    auto* self = reinterpret_cast<CExplorerBHO*>(refData);
    if (!self) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    bool handled = false;
    LRESULT result = self->HandleStatusBarMessage(hwnd, msg, wParam, lParam, &handled);
    if (handled) {
        return result;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

}  // namespace shelltabs
