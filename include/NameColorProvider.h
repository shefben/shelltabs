#pragma once

#include <windows.h>
#include <CommCtrl.h>

#include <optional>
#include <string>
#include <vector>

namespace shelltabs {

// Central resolver that unifies persistent filename overrides, ephemeral
// highlights, and tag-derived colours into a single lookup surface.
class NameColorProvider {
public:
    struct ItemAppearance {
        std::optional<COLORREF> textColor;
        std::optional<COLORREF> backgroundColor;
        HFONT font = nullptr;
        bool ownsFont = false;
        bool applyWhenSelected = false;
        bool applyWhenHot = false;

        bool HasOverrides() const {
            return textColor.has_value() || backgroundColor.has_value() || font != nullptr;
        }

        bool AllowsForState(UINT state) const {
            const bool isSelected = (state & (CDIS_SELECTED | CDIS_MARKED | CDIS_DROPHILITED)) != 0;
            const bool isHot = (state & CDIS_HOT) != 0;
            if (isSelected && !applyWhenSelected) {
                return false;
            }
            if (isHot && !applyWhenHot) {
                return false;
            }
            return true;
        }
    };

    static NameColorProvider& Instance();

    // Returns true when a highlight colour exists for |path|. The resolver first
    // honours explicit per-path overrides (persistent or ephemeral) and falls
    // back to tag-derived colours when no override is present.
    bool TryGetColorForPath(const std::wstring& path, COLORREF* color) const;

    ItemAppearance GetAppearanceForPath(const std::wstring& path) const;

    // Convenience helper that returns the resolved colour and any tag labels
    // associated with |path|. If an override colour exists it takes precedence
    // but the tag metadata is still surfaced for UI consumers.
    bool TryGetColorAndTags(const std::wstring& path, COLORREF* color,
                            std::vector<std::wstring>* tags) const;

private:
    NameColorProvider() = default;
};

}  // namespace shelltabs

