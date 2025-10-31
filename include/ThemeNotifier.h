#pragma once

#include <windows.h>

#include <functional>
#include <memory>
#include <mutex>

namespace shelltabs {

class ThemeNotifier {
public:
    struct ThemeColors {
        COLORREF background = RGB(255, 255, 255);
        COLORREF foreground = RGB(0, 0, 0);
        bool valid = false;
    };

    ThemeNotifier();
    ~ThemeNotifier();

    ThemeNotifier(const ThemeNotifier&) = delete;
    ThemeNotifier& operator=(const ThemeNotifier&) = delete;

    bool Initialize(HWND window, std::function<void()> callback);
    void Shutdown();

    ThemeColors GetThemeColors() const;

    // Forces a synchronous refresh of the cached system colors. This is used when
    // Explorer broadcasts legacy theme messages (WM_THEMECHANGED/WM_SYSCOLORCHANGE)
    // so the band can update even if WinRT notifications are unavailable.
    void RefreshColorsFromSystem();

    bool HandleSessionChange(WPARAM sessionEvent, LPARAM lParam);

#if defined(SHELLTABS_ENABLE_THEME_TEST_HOOKS)
    void SimulateColorChangeForTest();
    void SimulateSessionEventForTest(DWORD sessionEvent);
#endif

private:
    void NotifyThemeChanged();
    void UpdateColorSnapshot();

    HWND m_window = nullptr;
    std::function<void()> m_callback;

    struct UiSettingsState;
    std::unique_ptr<UiSettingsState> m_uiSettings;

    bool m_wtsRegistered = false;

    mutable std::mutex m_mutex;
    ThemeColors m_cachedColors;
};

}  // namespace shelltabs

