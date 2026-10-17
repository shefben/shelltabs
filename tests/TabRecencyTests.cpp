#include "TabBandWindow.h"
#include "ColorUtils.h"

#include <algorithm>
#include <cmath>
#include <iostream>

namespace shelltabs {
struct TabBandWindowDiffTestHarness {
    static double Opacity(const TabViewItem& tab, const std::vector<TabViewItem>& tabs) {
        return TabBandWindow::ComputeTabRecencyOpacity(tab, tabs);
    }
};
}

int main() {
    using namespace shelltabs;
    std::vector<TabViewItem> tabs(7);
    for (size_t i = 0; i < tabs.size(); ++i) {
        tabs[i].type = TabViewItemType::kTab;
        tabs[i].location = {0, static_cast<int>(i)};
        tabs[i].activationOrdinal = i + 1;
    }
    bool success = true;
    auto expect = [&](const TabViewItem& tab, double opacity, const char* message) {
        if (std::abs(TabBandWindowDiffTestHarness::Opacity(tab, tabs) - opacity) > 0.0001) {
            std::cerr << message << '\n';
            success = false;
        }
    };
    for (size_t i = 0; i < tabs.size(); ++i) {
        expect(tabs[i], i < 2 ? 0.0 : static_cast<double>(i - 1) / 5.0,
               "Only the five newest viewed tabs should have descending opacity");
    }
    tabs[0].activationOrdinal = 8;
    expect(tabs[0], 1.0, "Revisiting a tab should make it newest");
    expect(tabs[2], 0.0, "Revisiting should evict the oldest indicator");
    std::reverse(tabs.begin(), tabs.end());
    expect(tabs.back(), 1.0, "Reordering tabs should preserve recency");
    tabs.pop_back();
    expect(tabs.front(), 1.0, "Closing the newest tab should promote remaining tabs");
    TabViewItem backgroundTab{};
    backgroundTab.type = TabViewItemType::kTab;
    tabs.push_back(backgroundTab);
    expect(tabs.back(), 0.0, "Never-viewed background tabs should have no indicator");
    TabViewItem header = tabs.front();
    header.type = TabViewItemType::kGroupHeader;
    header.activationOrdinal = 100;
    tabs.push_back(header);
    expect(tabs.back(), 0.0, "Group headers should have no recency indicator");
    expect(tabs.front(), 1.0, "Group headers should not consume recency ranks");
    tabs.resize(1);
    tabs.front().activationOrdinal = 0;
    tabs.front().lastActivatedTick = 123;
    expect(tabs.front(), 1.0, "Timestamp-only activation should remain supported");
    tabs.assign(7, tabs.front());
    for (size_t i = 0; i < tabs.size(); ++i) {
        tabs[i].stableId = i + 1;
    }
    size_t indicators = 0;
    for (const auto& tab : tabs) {
        if (TabBandWindowDiffTestHarness::Opacity(tab, tabs) > 0.0) {
            ++indicators;
        }
    }
    if (indicators != 5) {
        std::cerr << "Tied activation timestamps must still respect the five-tab limit\n";
        success = false;
    }
    // Exercise light/dark themes and tab/accent colors that previously hid the line.
    const COLORREF backgrounds[] = {RGB(0, 0, 0), RGB(32, 32, 32), RGB(255, 255, 255),
        RGB(240, 240, 240), RGB(128, 128, 128), RGB(180, 20, 30), RGB(10, 90, 180),
        RGB(20, 140, 60), RGB(250, 210, 20)};
    const COLORREF accents[] = {RGB(0, 0, 0), RGB(255, 255, 255), RGB(0, 120, 215),
        RGB(180, 20, 30), RGB(20, 140, 60)};
    for (const auto background : backgrounds) {
        for (const auto accent : accents) {
            double previousContrast = 0.0;
            COLORREF previousColor = background;
            const COLORREF fullColor = ResolveTabRecencyColor(background, accent, 1.0);
            for (int rank = 1; rank <= 5; ++rank) {
                const double opacity = rank / 5.0;
                const auto color = ResolveTabRecencyColor(background, accent, opacity);
                const auto blendChannel = [=](BYTE base, BYTE full) {
                    return static_cast<BYTE>(std::lround(base * (1.0 - opacity) + full * opacity));
                };
                const COLORREF expected = RGB(blendChannel(GetRValue(background), GetRValue(fullColor)),
                    blendChannel(GetGValue(background), GetGValue(fullColor)),
                    blendChannel(GetBValue(background), GetBValue(fullColor)));
                if (color != expected) {
                    std::cerr << "Dot colors must use actual 20/40/60/80/100 percent blending\n";
                    success = false;
                    break;
                }
                const int separation = std::max({
                    std::abs(static_cast<int>(GetRValue(color)) - GetRValue(previousColor)),
                    std::abs(static_cast<int>(GetGValue(color)) - GetGValue(previousColor)),
                    std::abs(static_cast<int>(GetBValue(color)) - GetBValue(previousColor))});
                const double contrast = ComputeContrastRatio(ComputeColorLuminance(background),
                                                              ComputeColorLuminance(color));
                if (separation < 16 || contrast <= previousContrast) {
                    std::cerr << "Dot levels must have clearly separated colors and increasing contrast\n";
                    success = false;
                    break;
                }
                previousContrast = contrast;
                previousColor = color;
            }
        }
    }
    if (ResolveTabRecencyColor(RGB(255, 255, 255), RGB(0, 0, 128), 1.0) != RGB(0, 0, 128)) {
        std::cerr << "An accent with sufficient contrast should retain its color\n";
        success = false;
    }
    return success ? 0 : 1;
}
