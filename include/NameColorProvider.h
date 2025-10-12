#pragma once

#include <windows.h>

#include <string>
#include <vector>

namespace shelltabs {

// Central resolver that unifies persistent filename overrides, ephemeral
// highlights, and tag-derived colours into a single lookup surface.
class NameColorProvider {
public:
    static NameColorProvider& Instance();

    // Returns true when a highlight colour exists for |path|. The resolver first
    // honours explicit per-path overrides (persistent or ephemeral) and falls
    // back to tag-derived colours when no override is present.
    bool TryGetColorForPath(const std::wstring& path, COLORREF* color) const;

    // Convenience helper that returns the resolved colour and any tag labels
    // associated with |path|. If an override colour exists it takes precedence
    // but the tag metadata is still surfaced for UI consumers.
    bool TryGetColorAndTags(const std::wstring& path, COLORREF* color,
                            std::vector<std::wstring>* tags) const;

private:
    NameColorProvider() = default;
};

}  // namespace shelltabs

