/*
 * ShellTabs - File Color Hooks
 *
 * Subclasses SHELLDLL_DefView so we can intercept NM_CUSTOMDRAW notifications
 * from the SysListView32 child and override per-item text colors.
 */

#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <commctrl.h>

#include <mutex>
#include <string>
#include <unordered_map>
#include <unordered_set>

namespace shelltabs {

class FileColorHooks {
public:
    static FileColorHooks& Instance();

    bool Initialize();
    void Shutdown();
    bool IsInitialized() const { return m_initialized; }

    // Set the override color for a file by absolute path. Use CLR_INVALID to
    // clear the override.
    void SetColorForPath(const std::wstring& absolutePath, COLORREF color);

    // Convenience: clear all stored color overrides.
    void ClearAll();

    // Called from CExplorerBHO when a frame's current folder is known. The
    // BHO may call this multiple times (e.g., on every navigation); we use it
    // both to record the per-frame folder and to install the subclass on the
    // SHELLDLL_DefView if it has appeared since the last call.
    void SetFrameCurrentFolder(HWND explorerFrame, const std::wstring& folderPath);

    // Called when the explorer frame is going away.
    void ClearFrame(HWND explorerFrame);

private:
    FileColorHooks() = default;
    ~FileColorHooks();
    FileColorHooks(const FileColorHooks&) = delete;
    FileColorHooks& operator=(const FileColorHooks&) = delete;

    void TrySubclassDefView(HWND explorerFrame);
    void SubclassHwnd(HWND defView, HWND explorerFrame);
    void UnsubclassAll();

    static LRESULT CALLBACK DefViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                                UINT_PTR id, DWORD_PTR refData);

    bool LookupColorForItemText(HWND defView, const std::wstring& displayName, COLORREF* outColor) const;

    bool m_initialized = false;
    mutable std::mutex m_mutex;

    // path (lowercased, normalized) -> color
    std::unordered_map<std::wstring, COLORREF> m_pathToColor;
    // basename (lowercased) -> color, for fast NM_CUSTOMDRAW matching
    std::unordered_map<std::wstring, COLORREF> m_basenameToColor;

    // explorerFrame HWND -> normalized (lowercased) folder path
    std::unordered_map<HWND, std::wstring> m_frameFolders;

    // SHELLDLL_DefView HWND set tracked here so we can clean up.
    std::unordered_set<HWND> m_subclassedDefViews;
};

bool InitializeFileColorHooks();
void ShutdownFileColorHooks();

}  // namespace shelltabs
