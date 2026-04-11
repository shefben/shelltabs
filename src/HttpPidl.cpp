#include "HttpPidl.h"

#include <objbase.h>
#include <shlwapi.h>

#include <algorithm>
#include <cstring>
#include <functional>
#include <limits>
#include <string_view>

namespace shelltabs::http {

namespace {

#pragma pack(push, 1)
struct ItemHeader {
    std::uint32_t signature = kItemSignature;
    std::uint8_t version = kItemVersion;
    std::uint8_t type = 0;
    std::uint8_t componentCount = 0;
    std::uint8_t reserved = 0;
};

struct ComponentHeader {
    std::uint8_t type = 0;
    std::uint8_t reserved = 0;
    std::uint16_t size = 0;
};
#pragma pack(pop)

constexpr std::size_t kAlignment = alignof(std::uint16_t);

std::size_t AlignSize(std::size_t size) {
    return (size + (kAlignment - 1)) & ~(kAlignment - 1);
}

std::span<std::byte> Reserve(std::vector<std::byte>* buffer, std::size_t size) {
    const std::size_t offset = buffer->size();
    buffer->resize(offset + size);
    return {buffer->data() + offset, size};
}

const std::byte* Advance(const std::byte* cursor, const std::byte* end, std::size_t amount) {
    if (cursor > end || amount > static_cast<std::size_t>(end - cursor)) {
        return nullptr;
    }
    return cursor + amount;
}

std::wstring_view AsWideString(std::span<const std::byte> payload) {
    if (payload.empty()) {
        return {};
    }
    if (payload.size() % sizeof(wchar_t) != 0) {
        return {};
    }
    const wchar_t* data = reinterpret_cast<const wchar_t*>(payload.data());
    return {data, payload.size() / sizeof(wchar_t)};
}

bool ForEachComponent(const SHITEMID& item,
                      const std::function<bool(const ComponentHeader&, std::span<const std::byte>)>& callback) {
    if (!IsHttpItemId(item)) {
        return false;
    }
    const auto* header = reinterpret_cast<const ItemHeader*>(item.abID);
    const std::byte* cursor = reinterpret_cast<const std::byte*>(header + 1);
    const std::byte* end = reinterpret_cast<const std::byte*>(&item) + item.cb;
    for (std::uint8_t index = 0; index < header->componentCount; ++index) {
        const auto* componentHeader = reinterpret_cast<const ComponentHeader*>(cursor);
        cursor = Advance(cursor, end, sizeof(ComponentHeader));
        if (!cursor) {
            return false;
        }
        const std::byte* payloadEnd = Advance(cursor, end, componentHeader->size);
        if (!payloadEnd) {
            return false;
        }
        std::span<const std::byte> payload(cursor, componentHeader->size);
        cursor = Advance(cursor, end, AlignSize(componentHeader->size));
        if (!cursor) {
            return false;
        }
        if (!callback(*componentHeader, payload)) {
            break;
        }
    }
    return true;
}

std::vector<std::wstring> SplitPath(const std::wstring& path) {
    std::vector<std::wstring> segments;
    if (path.empty()) {
        return segments;
    }
    std::size_t index = 0;
    while (index < path.size()) {
        std::size_t next = path.find(L'/', index);
        std::wstring segment;
        if (next == std::wstring::npos) {
            segment = path.substr(index);
            index = path.size();
        } else {
            segment = path.substr(index, next - index);
            index = next + 1;
        }
        if (!segment.empty()) {
            segments.push_back(std::move(segment));
        }
    }
    return segments;
}

std::wstring BuildCanonicalUrl(const HttpUrlParts& parts) {
    std::wstring url = parts.useHttps ? L"https://" : L"http://";
    url += parts.host;
    const unsigned short defaultPort = parts.useHttps ? 443 : 80;
    if (parts.port != defaultPort && parts.port != 0) {
        url += L":" + std::to_wstring(parts.port);
    }
    if (!parts.basePath.empty()) {
        if (parts.basePath.front() != L'/') {
            url.push_back(L'/');
        }
        url += parts.basePath;
    } else {
        url.push_back(L'/');
    }
    if (!url.empty() && url.back() != L'/') {
        url.push_back(L'/');
    }
    return url;
}

}  // namespace

PidlBuilder::PidlBuilder() = default;
PidlBuilder::~PidlBuilder() = default;

HRESULT PidlBuilder::Append(ItemType type, std::span<const ComponentDefinition> components) {
    std::size_t payloadSize = sizeof(ItemHeader);
    for (const auto& component : components) {
        if (!component.data && component.size != 0) {
            return E_INVALIDARG;
        }
        if (component.size > std::numeric_limits<std::uint16_t>::max()) {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        payloadSize += sizeof(ComponentHeader);
        payloadSize += AlignSize(component.size);
    }
    if (payloadSize > std::numeric_limits<std::uint16_t>::max()) {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    const std::size_t totalSize = payloadSize + sizeof(std::uint16_t);
    auto span = Reserve(&buffer_, totalSize);
    auto* item = reinterpret_cast<SHITEMID*>(span.data());
    item->cb = static_cast<std::uint16_t>(totalSize);
    auto* header = reinterpret_cast<ItemHeader*>(item->abID);
    header->signature = kItemSignature;
    header->version = kItemVersion;
    header->type = static_cast<std::uint8_t>(type);
    header->componentCount = static_cast<std::uint8_t>(components.size());
    header->reserved = 0;

    std::byte* cursor = reinterpret_cast<std::byte*>(header + 1);
    for (const auto& component : components) {
        auto* componentHeader = reinterpret_cast<ComponentHeader*>(cursor);
        componentHeader->type = static_cast<std::uint8_t>(component.type);
        componentHeader->reserved = 0;
        componentHeader->size = static_cast<std::uint16_t>(component.size);
        cursor += sizeof(ComponentHeader);
        if (component.size != 0) {
            std::memcpy(cursor, component.data, component.size);
        }
        cursor += AlignSize(component.size);
    }

    ++itemCount_;
    return S_OK;
}

UniquePidl PidlBuilder::Finalize() {
    if (buffer_.empty()) {
        buffer_.resize(sizeof(std::uint16_t));
    } else {
        buffer_.resize(buffer_.size() + sizeof(std::uint16_t));
    }
    auto* terminator = reinterpret_cast<std::uint16_t*>(buffer_.data() + buffer_.size() - sizeof(std::uint16_t));
    *terminator = 0;

    auto* raw = static_cast<PIDLIST_ABSOLUTE>(CoTaskMemAlloc(buffer_.size()));
    if (!raw) {
        return nullptr;
    }
    std::memcpy(raw, buffer_.data(), buffer_.size());
    return UniquePidl(raw);
}

bool IsHttpItemId(const SHITEMID& item) noexcept {
    if (item.cb < sizeof(SHITEMID) + sizeof(ItemHeader)) {
        return false;
    }
    const auto* header = reinterpret_cast<const ItemHeader*>(item.abID);
    return header->signature == kItemSignature && header->version == kItemVersion;
}

ItemType GetItemType(const SHITEMID& item) noexcept {
    if (!IsHttpItemId(item)) {
        return ItemType::Directory;
    }
    const auto* header = reinterpret_cast<const ItemHeader*>(item.abID);
    return static_cast<ItemType>(header->type);
}

bool TryGetComponentString(const SHITEMID& item, ComponentType component, std::wstring* value) {
    if (!value) {
        return false;
    }
    bool found = false;
    ForEachComponent(item, [&](const ComponentHeader& header, std::span<const std::byte> payload) -> bool {
        if (static_cast<ComponentType>(header.type) == component) {
            auto view = AsWideString(payload);
            *value = std::wstring(view);
            found = true;
            return false;
        }
        return true;
    });
    return found;
}

bool TryGetComponentUint16(const SHITEMID& item, ComponentType component, std::uint16_t* value) {
    if (!value) {
        return false;
    }
    bool found = false;
    ForEachComponent(item, [&](const ComponentHeader& header, std::span<const std::byte> payload) -> bool {
        if (static_cast<ComponentType>(header.type) == component) {
            if (payload.size() == sizeof(std::uint16_t)) {
                std::memcpy(value, payload.data(), sizeof(std::uint16_t));
                found = true;
            }
            return false;
        }
        return true;
    });
    return found;
}

bool TryGetFindData(const SHITEMID& item, WIN32_FIND_DATAW* data) {
    if (!data) {
        return false;
    }
    bool found = false;
    ForEachComponent(item, [&](const ComponentHeader& header, std::span<const std::byte> payload) -> bool {
        if (static_cast<ComponentType>(header.type) == ComponentType::FindData) {
            if (payload.size() == sizeof(WIN32_FIND_DATAW)) {
                std::memcpy(data, payload.data(), sizeof(WIN32_FIND_DATAW));
                found = true;
            }
            return false;
        }
        return true;
    });
    return found;
}

bool TryGetFindData(PCUIDLIST_RELATIVE pidl, WIN32_FIND_DATAW* data) {
    if (!pidl || !data) {
        return false;
    }
    if (pidl->mkid.cb == 0) {
        return false;
    }
    return TryGetFindData(pidl->mkid, data);
}

bool TryParseHttpPidl(PCIDLIST_ABSOLUTE pidl, HttpUrlParts* parts, std::vector<std::wstring>* segments,
                      bool* terminalIsDirectory) {
    if (!pidl || !parts || !segments) {
        return false;
    }
    segments->clear();
    if (terminalIsDirectory) {
        *terminalIsDirectory = true;
    }

    // Walk past any non-HTTP items (namespace prefix from Explorer).
    const BYTE* cursor = reinterpret_cast<const BYTE*>(pidl);
    bool foundRoot = false;

    while (true) {
        const auto* item = reinterpret_cast<const SHITEMID*>(cursor);
        if (item->cb == 0) {
            break;
        }
        if (IsHttpItemId(*item)) {
            ItemType type = GetItemType(*item);
            if (type == ItemType::Root && !foundRoot) {
                foundRoot = true;
                TryGetComponentString(*item, ComponentType::Host, &parts->host);
                std::uint16_t port = 0;
                if (TryGetComponentUint16(*item, ComponentType::Port, &port)) {
                    parts->port = port;
                }
                TryGetComponentString(*item, ComponentType::BasePath, &parts->basePath);
                std::wstring scheme;
                if (TryGetComponentString(*item, ComponentType::Scheme, &scheme)) {
                    parts->useHttps = (scheme == L"https");
                }
                TryGetComponentString(*item, ComponentType::Name, &parts->displayName);
            } else if (type == ItemType::Directory) {
                std::wstring name;
                if (TryGetComponentString(*item, ComponentType::Name, &name)) {
                    segments->push_back(std::move(name));
                }
                if (terminalIsDirectory) {
                    *terminalIsDirectory = true;
                }
            } else if (type == ItemType::File) {
                std::wstring name;
                if (TryGetComponentString(*item, ComponentType::Name, &name)) {
                    segments->push_back(std::move(name));
                }
                if (terminalIsDirectory) {
                    *terminalIsDirectory = false;
                }
            }
        }
        cursor += item->cb;
    }

    if (foundRoot) {
        parts->canonicalUrl = BuildCanonicalUrl(*parts);
    }
    return foundRoot;
}

std::wstring BuildUrlFromHttpPidl(PCIDLIST_ABSOLUTE pidl) {
    HttpUrlParts parts;
    std::vector<std::wstring> segments;
    bool isDirectory = true;
    if (!TryParseHttpPidl(pidl, &parts, &segments, &isDirectory)) {
        return {};
    }
    std::wstring url = parts.useHttps ? L"https://" : L"http://";
    url += parts.host;
    const unsigned short defaultPort = parts.useHttps ? 443 : 80;
    if (parts.port != defaultPort && parts.port != 0) {
        url += L":" + std::to_wstring(parts.port);
    }
    if (!parts.basePath.empty()) {
        if (parts.basePath.front() != L'/') {
            url.push_back(L'/');
        }
        url += parts.basePath;
    }
    for (const auto& segment : segments) {
        if (url.empty() || url.back() != L'/') {
            url.push_back(L'/');
        }
        url += segment;
    }
    if (isDirectory && !url.empty() && url.back() != L'/') {
        url.push_back(L'/');
    }
    return url;
}

UniquePidl CreatePidlFromHttpUrl(const HttpUrlParts& parts) {
    PidlBuilder builder;

    std::wstring scheme = parts.useHttps ? L"https" : L"http";
    std::uint16_t port = parts.port;

    ComponentDefinition hostComp{ComponentType::Host, parts.host.c_str(), parts.host.size() * sizeof(wchar_t)};
    ComponentDefinition portComp{ComponentType::Port, &port, sizeof(port)};
    ComponentDefinition basePathComp{ComponentType::BasePath, parts.basePath.c_str(),
                                     parts.basePath.size() * sizeof(wchar_t)};
    ComponentDefinition schemeComp{ComponentType::Scheme, scheme.c_str(), scheme.size() * sizeof(wchar_t)};
    ComponentDefinition nameComp{ComponentType::Name, parts.displayName.c_str(),
                                  parts.displayName.size() * sizeof(wchar_t)};

    HRESULT hr = builder.Append(ItemType::Root, {hostComp, portComp, basePathComp, schemeComp, nameComp});
    if (FAILED(hr)) {
        return nullptr;
    }

    return builder.Finalize();
}

UniquePidl CloneRelativeHttpPidl(PCUIDLIST_RELATIVE pidl) {
    if (!pidl) {
        return nullptr;
    }
    return UniquePidl(ILClone(pidl));
}

std::vector<std::uint8_t> SerializeHttpPidl(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return {};
    }
    UINT size = ILGetSize(pidl);
    if (size == 0) {
        return {};
    }
    std::vector<std::uint8_t> bytes(size);
    std::memcpy(bytes.data(), pidl, size);
    return bytes;
}

}  // namespace shelltabs::http
