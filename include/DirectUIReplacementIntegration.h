#pragma once

#include <Windows.h>

namespace ShellTabs {
    class CustomFileListView;
}

class ExplorerGlowCoordinator;

namespace shelltabs {

// Integration layer for DirectUI replacement system
class DirectUIReplacementIntegration {
public:
    // Initialize the DirectUI replacement system
    // Call this once during BHO initialization
    static bool Initialize();

    // Shutdown the DirectUI replacement system
    // Call this during BHO shutdown
    static void Shutdown();

    // Check if replacement system is active
    static bool IsEnabled();

    // Enable or disable the replacement (can be toggled via settings)
    static void SetEnabled(bool enabled);

    // Get the custom view instance for a given window
    static ShellTabs::CustomFileListView* GetCustomViewForWindow(HWND hwnd);

    // Callback for when a custom view is created
    // This allows CExplorerBHO to configure it with coordinators, etc.
    static void SetCustomViewCreatedCallback(
        void (*callback)(ShellTabs::CustomFileListView* view, HWND hwnd, void* context),
        void* context);

private:
    static bool s_initialized;
    static bool s_enabled;
    static void (*s_viewCreatedCallback)(ShellTabs::CustomFileListView*, HWND, void*);
    static void* s_viewCreatedContext;

    friend class ShellTabs::CustomFileListView;
    static void NotifyViewCreated(ShellTabs::CustomFileListView* view, HWND hwnd);
};

// Helper function to check if a window is a DirectUIHWND or custom replacement
bool IsDirectUIWindow(HWND hwnd);

// Helper to get DirectUI window for a shell view
HWND FindDirectUIWindow(HWND shellViewWindow);

} // namespace shelltabs
