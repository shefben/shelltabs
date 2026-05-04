/*
 * ShellTabs - File Color Hooks Implementation
 *
 * Subclasses each Explorer frame's SHELLDLL_DefView and intercepts
 * NM_CUSTOMDRAW notifications from its child SysListView32 (Details / List /
 * Tile / Content / SmallIcons / Icons views) to override the text color of
 * matching items.
 *
 * Modern Explorer also uses DirectUIHWND in some configurations. That control
 * does not produce NM_CUSTOMDRAW notifications, so the override here applies
 * only when the legacy SysListView32 is in use. ExplorerPatcher and several
 * other tweaks routinely fall back to the SysListView32 surface.
 */

#include "../include/FileColorHooks.h"
#include "../include/Logging.h"

#include <algorithm>
#include <cwctype>

namespace shelltabs {

namespace {

std::wstring ToLowerCopy(std::wstring text) {
    std::transform(text.begin(), text.end(), text.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return text;
}

std::wstring NormalizePath(const std::wstring& path) {
    std::wstring out = ToLowerCopy(path);
    // Strip trailing slash for consistent comparisons.
    while (out.size() > 1 && (out.back() == L'\\' || out.back() == L'/')) {
        out.pop_back();
    }
    // Normalize separators to backslash.
    std::replace(out.begin(), out.end(), L'/', L'\\');
    return out;
}

std::wstring ExtractBasename(const std::wstring& path) {
    const size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring::npos) {
        return path;
    }
    return path.substr(pos + 1);
}

}  // namespace

FileColorHooks& FileColorHooks::Instance() {
    static FileColorHooks instance;
    return instance;
}

FileColorHooks::~FileColorHooks() {
    Shutdown();
}

bool FileColorHooks::Initialize() {
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (m_initialized) {
            return true;
        }
        m_initialized = true;
    }

    // Default test entry requested in the task brief: paint README.md red.
    SetColorForPath(L"F:\\development\\steam\\emulator_bot\\shelltabs\\README.md",
                    RGB(255, 0, 0));

    LogMessage(LogLevel::Info, L"FileColorHooks initialized");
    return true;
}

void FileColorHooks::Shutdown() {
    UnsubclassAll();
    std::lock_guard<std::mutex> lock(m_mutex);
    if (!m_initialized) {
        return;
    }
    m_pathToColor.clear();
    m_basenameToColor.clear();
    m_frameFolders.clear();
    m_initialized = false;
    LogMessage(LogLevel::Info, L"FileColorHooks shutdown");
}

void FileColorHooks::SetColorForPath(const std::wstring& absolutePath, COLORREF color) {
    if (absolutePath.empty()) {
        return;
    }
    const std::wstring normalized = NormalizePath(absolutePath);
    const std::wstring basename = ToLowerCopy(ExtractBasename(absolutePath));

    std::lock_guard<std::mutex> lock(m_mutex);
    if (color == CLR_INVALID) {
        m_pathToColor.erase(normalized);
        // Only erase the basename lookup if no other path uses this basename.
        bool stillReferenced = false;
        for (const auto& [path, _] : m_pathToColor) {
            if (ExtractBasename(path) == basename) {
                stillReferenced = true;
                break;
            }
        }
        if (!stillReferenced) {
            m_basenameToColor.erase(basename);
        }
    } else {
        m_pathToColor[normalized] = color;
        m_basenameToColor[basename] = color;
    }
}

void FileColorHooks::ClearAll() {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_pathToColor.clear();
    m_basenameToColor.clear();
}

void FileColorHooks::SetFrameCurrentFolder(HWND explorerFrame, const std::wstring& folderPath) {
    if (!explorerFrame) {
        return;
    }
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_frameFolders[explorerFrame] = NormalizePath(folderPath);
    }
    TrySubclassDefView(explorerFrame);
}

void FileColorHooks::ClearFrame(HWND explorerFrame) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_frameFolders.erase(explorerFrame);
}

namespace {

HWND FindDefViewInFrame(HWND explorerFrame) {
    HWND found = nullptr;
    EnumChildWindows(explorerFrame, [](HWND hwnd, LPARAM lp) -> BOOL {
        wchar_t cls[64] = {};
        GetClassNameW(hwnd, cls, 64);
        if (_wcsicmp(cls, L"SHELLDLL_DefView") == 0) {
            *reinterpret_cast<HWND*>(lp) = hwnd;
            return FALSE;
        }
        return TRUE;
    }, reinterpret_cast<LPARAM>(&found));
    return found;
}

}  // namespace

void FileColorHooks::TrySubclassDefView(HWND explorerFrame) {
    if (!explorerFrame || !IsWindow(explorerFrame)) {
        return;
    }
    HWND defView = FindDefViewInFrame(explorerFrame);
    if (!defView) {
        return;
    }
    SubclassHwnd(defView, explorerFrame);
}

void FileColorHooks::SubclassHwnd(HWND defView, HWND explorerFrame) {
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        if (m_subclassedDefViews.count(defView)) {
            return;
        }
    }
    if (!SetWindowSubclass(defView, &FileColorHooks::DefViewSubclassProc,
                           reinterpret_cast<UINT_PTR>(this),
                           reinterpret_cast<DWORD_PTR>(explorerFrame))) {
        LogMessage(LogLevel::Warning, L"FileColor: SetWindowSubclass failed defView=%p", defView);
        return;
    }
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_subclassedDefViews.insert(defView);
    }
    LogMessage(LogLevel::Info, L"FileColor: subclassed SHELLDLL_DefView %p (frame=%p)",
               defView, explorerFrame);
}

void FileColorHooks::UnsubclassAll() {
    std::unordered_set<HWND> toRemove;
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        toRemove = m_subclassedDefViews;
        m_subclassedDefViews.clear();
    }
    for (HWND hwnd : toRemove) {
        if (!IsWindow(hwnd)) {
            continue;
        }
        RemoveWindowSubclass(hwnd, &FileColorHooks::DefViewSubclassProc,
                             reinterpret_cast<UINT_PTR>(this));
    }
}

bool FileColorHooks::LookupColorForItemText(HWND /*defView*/, const std::wstring& displayName,
                                             COLORREF* outColor) const {
    if (displayName.empty() || !outColor) {
        return false;
    }
    const std::wstring key = ToLowerCopy(displayName);
    std::lock_guard<std::mutex> lock(m_mutex);
    auto it = m_basenameToColor.find(key);
    if (it == m_basenameToColor.end()) {
        return false;
    }
    *outColor = it->second;
    return true;
}

LRESULT CALLBACK FileColorHooks::DefViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                     UINT_PTR id, DWORD_PTR /*refData*/) {
    if (msg == WM_NCDESTROY) {
        FileColorHooks* self = reinterpret_cast<FileColorHooks*>(id);
        if (self) {
            std::lock_guard<std::mutex> lock(self->m_mutex);
            self->m_subclassedDefViews.erase(hwnd);
        }
        RemoveWindowSubclass(hwnd, &FileColorHooks::DefViewSubclassProc, id);
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    if (msg == WM_NOTIFY) {
        NMHDR* nmhdr = reinterpret_cast<NMHDR*>(lParam);
        if (nmhdr && nmhdr->code == NM_CUSTOMDRAW) {
            wchar_t cls[64] = {};
            GetClassNameW(nmhdr->hwndFrom, cls, 64);
            if (_wcsicmp(cls, WC_LISTVIEWW) == 0) {
                NMLVCUSTOMDRAW* lvcd = reinterpret_cast<NMLVCUSTOMDRAW*>(lParam);
                FileColorHooks* self = reinterpret_cast<FileColorHooks*>(id);

                switch (lvcd->nmcd.dwDrawStage) {
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;

                case CDDS_ITEMPREPAINT: {
                    if (!self) {
                        break;
                    }
                    // Read the item's display name.
                    wchar_t buf[MAX_PATH] = {};
                    LVITEMW item{};
                    item.iItem = static_cast<int>(lvcd->nmcd.dwItemSpec);
                    item.iSubItem = 0;
                    item.mask = LVIF_TEXT;
                    item.pszText = buf;
                    item.cchTextMax = ARRAYSIZE(buf);
                    SendMessageW(nmhdr->hwndFrom, LVM_GETITEMTEXTW,
                                 lvcd->nmcd.dwItemSpec, reinterpret_cast<LPARAM>(&item));
                    if (buf[0] == L'\0') {
                        break;
                    }
                    COLORREF color = 0;
                    if (self->LookupColorForItemText(hwnd, buf, &color)) {
                        lvcd->clrText = color;
                        return CDRF_NEWFONT;
                    }
                    break;
                }
                default:
                    break;
                }
            }
        }
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

bool InitializeFileColorHooks() {
    return FileColorHooks::Instance().Initialize();
}

void ShutdownFileColorHooks() {
    FileColorHooks::Instance().Shutdown();
}

}  // namespace shelltabs
