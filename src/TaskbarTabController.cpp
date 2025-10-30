#include "TaskbarTabController.h"

#include <VersionHelpers.h>

#include <algorithm>
#include <utility>
#include <cwchar>

#include "Logging.h"
#include "TabBand.h"
#include "TaskbarProxyWindow.h"

namespace shelltabs {

namespace {

struct ProxyLocationEquals {
    bool operator()(const TabLocation& a, const TabLocation& b) const noexcept {
        return a.groupIndex == b.groupIndex && a.tabIndex == b.tabIndex;
    }
};

bool LocationsEqual(const TabLocation& a, const TabLocation& b) {
    return ProxyLocationEquals{}(a, b);
}

bool TabsEqual(const std::vector<TaskbarTabController::CachedTab>& a,
               const std::vector<TaskbarTabController::CachedTab>& b) {
    if (a.size() != b.size()) {
        return false;
    }
    for (size_t i = 0; i < a.size(); ++i) {
        const auto& lhs = a[i];
        const auto& rhs = b[i];
        if (!LocationsEqual(lhs.location, rhs.location) || lhs.name != rhs.name ||
            lhs.tooltip != rhs.tooltip || lhs.selected != rhs.selected) {
            return false;
        }
    }
    return true;
}

}  // namespace

TaskbarTabController::TaskbarTabController(TabBand* owner) : m_owner(owner) {
    m_thumbButtonIcon = LoadIconW(nullptr, IDI_INFORMATION);
    if (!m_thumbButtonIcon) {
        m_thumbButtonIcon = LoadIconW(nullptr, IDI_APPLICATION);
    }
}

TaskbarTabController::~TaskbarTabController() { Reset(); }

bool TaskbarTabController::IsSupported() noexcept { return IsWindows7OrGreater() != FALSE; }

bool TaskbarTabController::EnsureTaskbar() {
    if (m_taskbar) {
        return true;
    }
    if (!IsSupported()) {
        return false;
    }

    Microsoft::WRL::ComPtr<ITaskbarList3> taskbar;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&taskbar));
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TaskbarTabController CoCreateInstance failed (hr=0x%08X)",
                   static_cast<unsigned int>(hr));
        return false;
    }
    hr = taskbar->HrInit();
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TaskbarTabController HrInit failed (hr=0x%08X)",
                   static_cast<unsigned int>(hr));
        return false;
    }
    m_taskbar = std::move(taskbar);
    return true;
}

void TaskbarTabController::SyncFrameSummary(const std::vector<TabViewItem>& items, TabLocation active,
                                            HWND frame) {
    if (!IsSupported()) {
        Reset();
        return;
    }

    if (!frame) {
        m_frame = nullptr;
        m_cachedTabs.clear();
        m_activeLocation = {};
        m_frameTooltip.clear();
        return;
    }

    if (!EnsureTaskbar()) {
        return;
    }

    if (frame != m_frame) {
        m_cachedTabs.clear();
        m_frameTooltip.clear();
        m_activeLocation = {};
        m_thumbButtonFrame = nullptr;
        m_thumbButtonAdded = false;
    }

    std::vector<CachedTab> tabs;
    tabs.reserve(items.size());
    for (const auto& item : items) {
        if (item.type != TabViewItemType::kTab) {
            continue;
        }
        CachedTab tab;
        tab.location = item.location;
        tab.name = item.name;
        tab.tooltip = item.tooltip;
        tab.selected = item.selected;
        tabs.push_back(std::move(tab));
    }

    TabLocation resolvedActive = active;
    if (!resolvedActive.IsValid()) {
        auto it = std::find_if(tabs.begin(), tabs.end(), [](const CachedTab& tab) { return tab.selected; });
        if (it != tabs.end()) {
            resolvedActive = it->location;
        }
    }

    for (auto& tab : tabs) {
        if (LocationsEqual(tab.location, resolvedActive)) {
            tab.selected = true;
        }
    }

    std::vector<FrameTabEntry> frameEntries;
    frameEntries.reserve(tabs.size());
    for (const auto& tab : tabs) {
        FrameTabEntry entry;
        entry.location = tab.location;
        entry.name = tab.name;
        entry.tooltip = tab.tooltip;
        entry.selected = tab.selected;
        frameEntries.push_back(std::move(entry));
    }

    const std::wstring tooltip = BuildFrameTooltip(frameEntries);

    const bool tabsChanged = !TabsEqual(m_cachedTabs, tabs);
    const bool tooltipChanged = tooltip != m_frameTooltip;

    m_frame = frame;
    if (tabsChanged) {
        m_cachedTabs = std::move(tabs);
    }
    m_activeLocation = resolvedActive;

    if (tooltipChanged) {
        RefreshFrameTooltip(frame, tooltip);
    }

    EnsureThumbnailButton(frame);
}

void TaskbarTabController::Reset() {
    m_taskbar.Reset();
    m_frame = nullptr;
    m_cachedTabs.clear();
    m_activeLocation = {};
    m_frameTooltip.clear();
    m_thumbButtonFrame = nullptr;
    m_thumbButtonAdded = false;
}

void TaskbarTabController::RefreshFrameTooltip(HWND frame, const std::wstring& tooltip) {
    if (!m_taskbar || !frame) {
        m_frameTooltip = tooltip;
        return;
    }

    HRESULT hr = m_taskbar->SetThumbnailTooltip(frame, tooltip.c_str());
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TaskbarTabController SetThumbnailTooltip failed (hr=0x%08X)",
                   static_cast<unsigned int>(hr));
        return;
    }

    m_frameTooltip = tooltip;
}

void TaskbarTabController::HandleThumbnailButton(const POINT& anchor) {
    if (!m_owner) {
        return;
    }

    POINT resolved = anchor;
    if (resolved.x == 0 && resolved.y == 0) {
        GetCursorPos(&resolved);
    }

    m_owner->ShowTaskbarPopup(resolved);
}

void TaskbarTabController::EnsureThumbnailButton(HWND frame) {
    if (!m_taskbar || !frame) {
        return;
    }

    THUMBBUTTON button{};
    button.iId = kThumbnailToolbarCommandId;
    button.dwMask = THB_FLAGS | THB_TOOLTIP;
    button.dwFlags = THBF_ENABLED;
    if (m_thumbButtonIcon) {
        button.dwMask |= THB_ICON;
        button.hIcon = m_thumbButtonIcon;
    }
    constexpr wchar_t kTooltip[] = L"Switch tabs";
    wcsncpy_s(button.szTip, ARRAYSIZE(button.szTip), kTooltip, _TRUNCATE);

    HRESULT hr = S_OK;
    if (!m_thumbButtonAdded || m_thumbButtonFrame != frame) {
        hr = m_taskbar->ThumbBarAddButtons(frame, 1, &button);
        if (FAILED(hr)) {
            if (hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)) {
                hr = m_taskbar->ThumbBarUpdateButtons(frame, 1, &button);
            }
        }
        if (FAILED(hr)) {
            LogMessage(LogLevel::Warning,
                       L"TaskbarTabController ThumbBarAddButtons failed (hr=0x%08X)",
                       static_cast<unsigned int>(hr));
            return;
        }
        m_thumbButtonAdded = true;
        m_thumbButtonFrame = frame;
        return;
    }

    hr = m_taskbar->ThumbBarUpdateButtons(frame, 1, &button);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"TaskbarTabController ThumbBarUpdateButtons failed (hr=0x%08X)",
                   static_cast<unsigned int>(hr));
    }
}

}  // namespace shelltabs

