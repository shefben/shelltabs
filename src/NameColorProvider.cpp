#include "NameColorProvider.h"

#include "FileColorOverrides.h"
#include "Tagging.h"

namespace shelltabs {

NameColorProvider& NameColorProvider::Instance() {
    static NameColorProvider instance;
    return instance;
}

bool NameColorProvider::TryGetColorForPath(const std::wstring& path, COLORREF* color) const {
    const ItemAppearance appearance = GetAppearanceForPath(path);
    if (!appearance.textColor.has_value()) {
        return false;
    }
    if (color) {
        *color = *appearance.textColor;
    }
    return true;
}

NameColorProvider::ItemAppearance NameColorProvider::GetAppearanceForPath(
        const std::wstring& path) const {
    ItemAppearance appearance;
    if (path.empty()) {
        return appearance;
    }

    COLORREF overrideColor = 0;
    if (FileColorOverrides::Instance().TryGetColor(path, &overrideColor)) {
        appearance.textColor = overrideColor;
        appearance.applyWhenHot = true;
        appearance.applyWhenSelected = true;
        return appearance;
    }

    COLORREF tagColor = 0;
    if (TagStore::Instance().TryGetColorForPath(path, &tagColor)) {
        appearance.textColor = tagColor;
        appearance.applyWhenHot = true;
        appearance.applyWhenSelected = true;
    }

    return appearance;
}

bool NameColorProvider::TryGetColorAndTags(const std::wstring& path, COLORREF* color,
                                           std::vector<std::wstring>* tags) const {
    if (tags) {
        tags->clear();
    }

    if (path.empty()) {
        return false;
    }

    std::vector<std::wstring> resolvedTags;
    COLORREF tagColor = 0;
    const bool hasTagColor = TagStore::Instance().TryGetColorAndTags(path, &tagColor, &resolvedTags);

    if (tags) {
        *tags = std::move(resolvedTags);
        if (!hasTagColor && tags->empty()) {
            *tags = TagStore::Instance().GetTagsForPath(path);
        }
    }

    const ItemAppearance appearance = GetAppearanceForPath(path);
    if (appearance.textColor.has_value()) {
        if (color) {
            *color = *appearance.textColor;
        }
        return true;
    }

    if (hasTagColor) {
        if (color) {
            *color = tagColor;
        }
        return true;
    }

    return false;
}

}  // namespace shelltabs

