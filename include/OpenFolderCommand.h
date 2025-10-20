#pragma once

#include <atomic>
#include <string>
#include <vector>

#include <windows.h>
#include <shobjidl_core.h>
#include <wrl/client.h>

namespace shelltabs {

class OpenFolderCommand : public IExplorerCommand, public IObjectWithSite {
public:
    OpenFolderCommand();
    ~OpenFolderCommand();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IExplorerCommand
    IFACEMETHODIMP GetTitle(IShellItemArray* itemArray, LPWSTR* name) override;
    IFACEMETHODIMP GetIcon(IShellItemArray* itemArray, LPWSTR* icon) override;
    IFACEMETHODIMP GetToolTip(IShellItemArray* itemArray, LPWSTR* infoTip) override;
    IFACEMETHODIMP GetCanonicalName(GUID* guidCommandName) override;
    IFACEMETHODIMP GetState(IShellItemArray* itemArray, BOOL okToBeSlow, EXPCMDSTATE* state) override;
    IFACEMETHODIMP Invoke(IShellItemArray* itemArray, IBindCtx* bindContext) override;
    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) override;
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** enumCommands) override;

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* site) override;
    IFACEMETHODIMP GetSite(REFIID riid, void** site) override;

private:
    void UpdateFrameWindow();
    bool HasOpenableFolder(IShellItemArray* items) const;
    bool CollectOpenablePaths(IShellItemArray* items, std::vector<std::wstring>* paths) const;
    bool OpenPathsInNewTabs(const std::vector<std::wstring>& paths) const;
    HWND FindBandWindow() const;
    HWND FindBandWindowRecursive(HWND parent) const;

    std::atomic<long> m_refCount{1};
    Microsoft::WRL::ComPtr<IServiceProvider> m_site;
    HWND m_frameWindow = nullptr;
};

}  // namespace shelltabs

