#pragma once

// Ensure Win7+ feature set regardless of who included Windows.h first.
#include <sdkddkver.h>

#if !defined(_WIN32_WINNT) || (_WIN32_WINNT < 0x0601)
#  undef _WIN32_WINNT
#  define _WIN32_WINNT 0x0601
#endif

#if !defined(NTDDI_VERSION) || (NTDDI_VERSION < 0x06010000)
#  undef NTDDI_VERSION
#  define NTDDI_VERSION 0x06010000
#endif

#include <Windows.h>
#include <servprov.h>   // IServiceProvider
#include <shobjidl.h>   // INamespaceTreeControlCustomDraw, NSTCCUSTOMDRAW
#include <wrl/client.h>
#include <string>
#include <unordered_map>
#include <unordered_set>


namespace shelltabs {

        class NamespaceTreeColorizer :
                public INameSpaceTreeControlCustomDraw,
                public INameSpaceTreeControlEvents
        {
        public:
                NamespaceTreeColorizer();
                virtual ~NamespaceTreeColorizer();

		// Call with IServiceProvider for the current Explorer window.
		bool Attach(Microsoft::WRL::ComPtr<IServiceProvider> sp);
		void Detach();

		// IUnknown
		IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
		IFACEMETHODIMP_(ULONG) AddRef() override;
		IFACEMETHODIMP_(ULONG) Release() override;

                // INameSpaceTreeControlCustomDraw
                IFACEMETHODIMP PrePaint(HDC, RECT*, LRESULT*) override;
                IFACEMETHODIMP PostPaint(HDC, RECT*) override;
                IFACEMETHODIMP ItemPrePaint(HDC hdc, RECT* bounds, NSTCCUSTOMDRAW* info,
                                        COLORREF* textColor, COLORREF* backgroundColor,
                                        LRESULT* result) override;
                IFACEMETHODIMP ItemPostPaint(HDC, RECT*, NSTCCUSTOMDRAW*) override;

                // INameSpaceTreeControlEvents
                IFACEMETHODIMP OnItemClick(IShellItem*, NSTCEHITTEST, NSTCECLICKTYPE) override;
                IFACEMETHODIMP OnPropertyItemCommit(IShellItem*) override;
                IFACEMETHODIMP OnItemStateChanging(IShellItem*, NSTCITEMSTATE, NSTCITEMSTATE) override;
                IFACEMETHODIMP OnItemStateChanged(IShellItem*, NSTCITEMSTATE, NSTCITEMSTATE) override;
                IFACEMETHODIMP OnSelectionChanged(IShellItemArray*) override;
                IFACEMETHODIMP OnKeyboardInput(UINT, WPARAM, LPARAM) override;
                IFACEMETHODIMP OnBeforeExpand(IShellItem*) override;
                IFACEMETHODIMP OnAfterExpand(IShellItem*) override;
                IFACEMETHODIMP OnBeginLabelEdit(IShellItem*) override;
                IFACEMETHODIMP OnEndLabelEdit(IShellItem*) override;
                IFACEMETHODIMP OnGetToolTip(IShellItem*, LPWSTR, int) override;
                IFACEMETHODIMP OnBeforeItemDelete(IShellItem*) override;
                IFACEMETHODIMP OnItemAdded(IShellItem*, BOOL) override;
                IFACEMETHODIMP OnItemDeleted(IShellItem*, BOOL) override;
                IFACEMETHODIMP OnBeforeContextMenu(IShellItem*, REFIID, void**) override;
                IFACEMETHODIMP OnAfterContextMenu(IShellItem*, IContextMenu*, REFIID, void**) override;
                IFACEMETHODIMP OnGetDefaultIconIndex(IShellItem*, int*, int*) override;
                IFACEMETHODIMP OnBeforeStateImageChange(IShellItem*) override;

        private:
                struct PendingItemPaint {
                        bool fillBackground = false;
                        COLORREF background = CLR_INVALID;
                };

                ULONG ref_ = 1;
                DWORD cookie_ = 0;
                Microsoft::WRL::ComPtr<INameSpaceTreeControl> nstc_;
                std::unordered_map<DWORD_PTR, PendingItemPaint> pending_paints_;
                std::unordered_set<HFONT> owned_fonts_;
                std::wstring pending_tree_rename_;

                bool ResolveNSTC(Microsoft::WRL::ComPtr<IServiceProvider> sp);
                bool ItemPathFromShellItem(IShellItem* psi, std::wstring* out) const;
        };

} // namespace shelltabs
