#include "FtpPidl.h"

#include <objbase.h>
#include <shlwapi.h>
#include <wininet.h>

#include <algorithm>
#include <cstring>
#include <functional>
#include <limits>
#include <string_view>

namespace shelltabs::ftp {

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

std::wstring BuildCanonicalUrl(const FtpUrlParts& parts) {
    std::wstring url = L"ftp://";
    if (!parts.userName.empty()) {
        url += parts.userName;
        if (!parts.password.empty()) {
            url += L":" + parts.password;
        }
        url += L"@";
    }
    url += parts.host;
    if (parts.port != INTERNET_DEFAULT_FTP_PORT && parts.port != 0) {
        url += L":" + std::to_wstring(parts.port);
    }
    if (!parts.path.empty()) {
        if (parts.path.front() != L'/') {
            url += L"/";
        }
        url += parts.path.substr(parts.path.front() == L'/' ? 1 : 0);
    } else {
        url += L"/";
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

    const std::size_t totalSize = payloadSize + sizeof(std::uint16_t);  // include SHITEMID.cb
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

bool IsFtpItemId(const SHITEMID& item) noexcept {
    if (item.cb < sizeof(SHITEMID) + sizeof(ItemHeader)) {
        return false;
    }
    const auto* header = reinterpret_cast<const ItemHeader*>(item.abID);
    return header->signature == kItemSignature && header->version == kItemVersion;
}

ItemType GetItemType(const SHITEMID& item) noexcept {
    if (!IsFtpItemId(item)) {
        return ItemType::Directory;
    }
    const auto* header = reinterpret_cast<const ItemHeader*>(item.abID);
    return static_cast<ItemType>(header->type);
}

bool ForEachComponent(const SHITEMID& item,
                      const std::function<bool(const ComponentHeader&, std::span<const std::byte>)>& callback) {
    if (!IsFtpItemId(item)) {
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

bool TryGetComponentString(const SHITEMID& item, ComponentType component, std::wstring* value) {
    if (!value) {
        return false;
    }
    bool found = false;
    ForEachComponent(item, [&](const ComponentHeader& header, std::span<const std::byte> payload) {
        if (header.type != static_cast<std::uint8_t>(component)) {
            return true;
        }
        std::wstring_view view = AsWideString(payload);
        if (view.empty() && !payload.empty()) {
            return false;
        }
        value->assign(view.begin(), view.end());
        found = true;
        return false;
    });
    return found;
}

bool TryGetComponentUint16(const SHITEMID& item, ComponentType component, std::uint16_t* value) {
    if (!value) {
        return false;
    }
    bool found = false;
    ForEachComponent(item, [&](const ComponentHeader& header, std::span<const std::byte> payload) {
        if (header.type != static_cast<std::uint8_t>(component)) {
            return true;
        }
        if (payload.size() != sizeof(std::uint16_t)) {
            return false;
        }
        std::uint16_t data = 0;
        std::memcpy(&data, payload.data(), sizeof(data));
        *value = data;
        found = true;
        return false;
    });
    return found;
}

bool TryParseFtpPidl(PCIDLIST_ABSOLUTE pidl, FtpUrlParts* parts, std::vector<std::wstring>* segments,
                     bool* terminalIsDirectory) {
    if (!pidl || !parts || !segments) {
        return false;
    }
    segments->clear();
    *parts = {};
    if (terminalIsDirectory) {
        *terminalIsDirectory = true;
    }

    const auto* current = pidl;
    bool rootSeen = false;
    while (current && current->mkid.cb) {
        const SHITEMID& item = current->mkid;
        if (!IsFtpItemId(item)) {
            return false;
        }
        ItemType type = GetItemType(item);
        switch (type) {
            case ItemType::Root: {
                if (rootSeen) {
                    return false;
                }
                rootSeen = true;
                if (!TryGetComponentString(item, ComponentType::Host, &parts->host)) {
                    return false;
                }
                if (!TryGetComponentString(item, ComponentType::UserName, &parts->userName)) {
                    parts->userName.clear();
                }
                if (!TryGetComponentString(item, ComponentType::Password, &parts->password)) {
                    parts->password.clear();
                }
                std::uint16_t port = INTERNET_DEFAULT_FTP_PORT;
                if (!TryGetComponentUint16(item, ComponentType::Port, &port)) {
                    port = INTERNET_DEFAULT_FTP_PORT;
                }
                parts->port = port;
                break;
            }
            case ItemType::Directory:
            case ItemType::File: {
                std::wstring name;
                if (!TryGetComponentString(item, ComponentType::Name, &name)) {
                    return false;
                }
                segments->push_back(std::move(name));
                if (terminalIsDirectory) {
                    *terminalIsDirectory = type != ItemType::File;
                }
                break;
            }
        }
        current = reinterpret_cast<PCIDLIST_ABSOLUTE>(reinterpret_cast<const std::byte*>(current) + current->mkid.cb);
    }
    if (!rootSeen) {
        return false;
    }

    std::wstring normalized = L"/";
    if (!segments->empty()) {
        normalized.clear();
        for (std::size_t index = 0; index < segments->size(); ++index) {
            if (index != 0) {
                normalized.push_back(L'/');
            }
            normalized += (*segments)[index];
        }
        if (terminalIsDirectory && *terminalIsDirectory) {
            normalized.push_back(L'/');
        }
        normalized.insert(normalized.begin(), L'/');
    }
    parts->path = normalized;
    parts->canonicalUrl = BuildCanonicalUrl(*parts);
    if (parts->userName.empty()) {
        parts->userName = L"anonymous";
    }
    return true;
}

std::wstring BuildUrlFromFtpPidl(PCIDLIST_ABSOLUTE pidl) {
    FtpUrlParts parts;
    std::vector<std::wstring> segments;
    bool directory = true;
    if (!TryParseFtpPidl(pidl, &parts, &segments, &directory)) {
        return {};
    }
    if (!segments.empty()) {
        parts.path.clear();
        for (std::size_t index = 0; index < segments.size(); ++index) {
            if (index != 0 || parts.path.empty()) {
                parts.path.push_back(L'/');
            }
            parts.path += segments[index];
        }
        if (directory && !segments.empty()) {
            parts.path.push_back(L'/');
        }
    }
    return BuildCanonicalUrl(parts);
}

UniquePidl CreatePidlFromFtpUrl(const FtpUrlParts& parts) {
    if (parts.host.empty()) {
        return nullptr;
    }
    PidlBuilder builder;
    std::wstring host = parts.host;
    std::wstring user = parts.userName;
    std::wstring password = parts.password;
    std::uint16_t port = parts.port == 0 ? INTERNET_DEFAULT_FTP_PORT : parts.port;

    ComponentDefinition rootComponents[] = {
        {ComponentType::Host, host.data(), host.size() * sizeof(wchar_t)},
        {ComponentType::UserName, user.data(), user.size() * sizeof(wchar_t)},
        {ComponentType::Password, password.data(), password.size() * sizeof(wchar_t)},
        {ComponentType::Port, &port, sizeof(port)},
    };
    if (FAILED(builder.Append(ItemType::Root, rootComponents))) {
        return nullptr;
    }

    std::wstring normalizedPath = parts.path.empty() ? L"/" : parts.path;
    if (!normalizedPath.empty() && normalizedPath.front() == L'/') {
        normalizedPath.erase(normalizedPath.begin());
    }
    bool lastIsDirectory = true;
    if (!normalizedPath.empty() && normalizedPath.back() != L'/') {
        lastIsDirectory = false;
    }
    if (!normalizedPath.empty() && normalizedPath.back() == L'/') {
        normalizedPath.pop_back();
        lastIsDirectory = true;
    }

    auto segments = SplitPath(normalizedPath);
    for (std::size_t index = 0; index < segments.size(); ++index) {
        ItemType type = (index + 1 == segments.size() && !lastIsDirectory) ? ItemType::File : ItemType::Directory;
        const std::wstring& segment = segments[index];
        ComponentDefinition component{ComponentType::Name, segment.data(), segment.size() * sizeof(wchar_t)};
        if (FAILED(builder.Append(type, {component}))) {
            return nullptr;
        }
    }

    return builder.Finalize();
}

UniquePidl CloneRelativeFtpPidl(PCUIDLIST_RELATIVE pidl) {
    if (!pidl) {
        return nullptr;
    }
    return UniquePidl(ILClone(pidl));
}

std::vector<std::uint8_t> SerializeFtpPidl(PCIDLIST_ABSOLUTE pidl) {
    std::vector<std::uint8_t> bytes;
    if (!pidl) {
        return bytes;
    }
    const UINT size = ILGetSize(pidl);
    if (size == 0) {
        return bytes;
    }
    bytes.resize(size);
    std::memcpy(bytes.data(), pidl, size);
    return bytes;
}

}  // namespace shelltabs::ftp

