#include "NameColorProvider.h"

#include "FileColorOverrides.h"
#include "Tagging.h"

namespace shelltabs {

NameColorProvider& NameColorProvider::Instance() {
    static NameColorProvider instance;
    return instance;
}

bool NameColorProvider::TryGetColorForPath(const std::wstring& path, COLORREF* color) const {
    if (path.empty()) {
        return false;
    }

    if (FileColorOverrides::Instance().TryGetColor(path, color)) {
        return true;
    }

    return TagStore::Instance().TryGetColorForPath(path, color);
}

bool NameColorProvider::TryGetColorAndTags(const std::wstring& path, COLORREF* color,
                                           std::vector<std::wstring>* tags) const {
    if (tags) {
        tags->clear();
    }

    if (path.empty()) {
        return false;
    }

    COLORREF overrideColor = 0;
    const bool hasOverride = FileColorOverrides::Instance().TryGetColor(path, &overrideColor);

    std::vector<std::wstring> resolvedTags;
    COLORREF tagColor = 0;
    const bool hasTagColor = TagStore::Instance().TryGetColorAndTags(path, &tagColor, &resolvedTags);

    if (tags) {
        *tags = std::move(resolvedTags);
        if (!hasTagColor && tags->empty()) {
            *tags = TagStore::Instance().GetTagsForPath(path);
        }
    }

    if (hasOverride) {
        if (color) {
            *color = overrideColor;
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

