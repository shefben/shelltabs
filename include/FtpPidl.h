#pragma once

#include "Utilities.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <span>
#include <string>
#include <vector>

#include <shlobj.h>

namespace shelltabs::ftp {

// Signature embedded in each Ftp PIDL item to identify our custom format.
inline constexpr std::uint32_t kItemSignature = 'PTFS';  // "SFTP" reversed for little-endian.
inline constexpr std::uint8_t kItemVersion = 1;

enum class ItemType : std::uint8_t {
    Root = 1,
    Directory = 2,
    File = 3,
};

enum class ComponentType : std::uint8_t {
    Host = 1,
    Port = 2,
    UserName = 3,
    Password = 4,
    Name = 5,
    Flags = 6,
    FindData = 7,
};

struct ComponentDefinition {
    ComponentType type = ComponentType::Name;
    const void* data = nullptr;
    std::size_t size = 0;
};

// Builder used to assemble ITEMIDLIST structures from component definitions.
class PidlBuilder {
public:
    PidlBuilder();
    ~PidlBuilder();

    PidlBuilder(const PidlBuilder&) = delete;
    PidlBuilder& operator=(const PidlBuilder&) = delete;

    HRESULT Append(ItemType type, std::span<const ComponentDefinition> components);
    HRESULT Append(ItemType type, std::initializer_list<ComponentDefinition> components) {
        return Append(type, std::span<const ComponentDefinition>(components.begin(), components.size()));
    }

    UniquePidl Finalize();

    std::size_t item_count() const noexcept { return itemCount_; }

private:
    std::vector<std::byte> buffer_;
    std::size_t itemCount_ = 0;
};

bool IsFtpItemId(const SHITEMID& item) noexcept;
ItemType GetItemType(const SHITEMID& item) noexcept;

bool TryGetComponentString(const SHITEMID& item, ComponentType component, std::wstring* value);
bool TryGetComponentUint16(const SHITEMID& item, ComponentType component, std::uint16_t* value);
bool TryGetFindData(const SHITEMID& item, WIN32_FIND_DATAW* data);
bool TryGetFindData(PCUIDLIST_RELATIVE pidl, WIN32_FIND_DATAW* data);

bool TryParseFtpPidl(PCIDLIST_ABSOLUTE pidl, FtpUrlParts* parts, std::vector<std::wstring>* segments,
                     bool* terminalIsDirectory);
std::wstring BuildUrlFromFtpPidl(PCIDLIST_ABSOLUTE pidl);
UniquePidl CreatePidlFromFtpUrl(const FtpUrlParts& parts);
UniquePidl CloneRelativeFtpPidl(PCUIDLIST_RELATIVE pidl);
std::vector<std::uint8_t> SerializeFtpPidl(PCIDLIST_ABSOLUTE pidl);

}  // namespace shelltabs::ftp

