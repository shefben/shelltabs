#pragma once

#include <windows.h>

#include <functional>
#include <mutex>
#include <unordered_map>
#include <vector>

namespace shelltabs {

struct GlowColorSet;

class DirectUiHooks {
public:
    using PaintCallback = std::function<void(const std::vector<RECT>& rectangles)>;

    static DirectUiHooks& Instance();

    void RegisterHost(HWND host);
    void UnregisterHost(HWND host);

    bool PaintHost(HWND host, const RECT& clientRect, const PaintCallback& callback);
    bool EnumerateRectangles(HWND host, const RECT& clientRect, std::vector<RECT>& rectangles) const;

private:
    DirectUiHooks() = default;
    ~DirectUiHooks() = default;

    DirectUiHooks(const DirectUiHooks&) = delete;
    DirectUiHooks& operator=(const DirectUiHooks&) = delete;

    struct HostEntry {
        HWND hwnd = nullptr;
        WNDPROC windowProc = nullptr;
        WNDPROC classProc = nullptr;
        bool attempted = false;
        bool resolved = false;
    };

    bool TryResolveHostLocked(HostEntry& entry);

    mutable std::mutex m_lock;
    std::unordered_map<HWND, HostEntry> m_hosts;
};

}  // namespace shelltabs

