#include "TaskbarProxyWindow.h"

#include <algorithm>

namespace shelltabs {

std::wstring BuildFrameTooltip(const std::vector<FrameTabEntry>& entries) {
    if (entries.empty()) {
        return std::wstring();
    }

    auto selected = std::find_if(entries.begin(), entries.end(), [](const FrameTabEntry& entry) {
        return entry.selected;
    });
    if (selected != entries.end()) {
        if (!selected->tooltip.empty()) {
            return selected->tooltip;
        }
        if (!selected->name.empty()) {
            return selected->name;
        }
    }

    for (const auto& entry : entries) {
        if (!entry.tooltip.empty()) {
            return entry.tooltip;
        }
    }

    return entries.front().name;
}

}  // namespace shelltabs

