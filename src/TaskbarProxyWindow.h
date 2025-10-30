#pragma once

#include <string>
#include <vector>

#include "TabManager.h"

namespace shelltabs {

struct FrameTabEntry {
    TabLocation location;
    std::wstring name;
    std::wstring tooltip;
    bool selected = false;
};

std::wstring BuildFrameTooltip(const std::vector<FrameTabEntry>& entries);

}  // namespace shelltabs

