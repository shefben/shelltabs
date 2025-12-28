#include "ContextMenuRegistry.h"

#include <shlobj.h>
#include <algorithm>
#include <cwctype>

#include "Logging.h"

namespace shelltabs {

namespace {

// RAII wrapper for registry keys
struct ScopedRegKey {
    ScopedRegKey() = default;
    explicit ScopedRegKey(HKEY value) : handle(value) {}
    ~ScopedRegKey() {
        if (handle) {
            RegCloseKey(handle);
        }
    }

    ScopedRegKey(const ScopedRegKey&) = delete;
    ScopedRegKey& operator=(const ScopedRegKey&) = delete;

    ScopedRegKey(ScopedRegKey&& other) noexcept : handle(other.handle) {
        other.handle = nullptr;
    }

    ScopedRegKey& operator=(ScopedRegKey&& other) noexcept {
        if (this != &other) {
            if (handle) {
                RegCloseKey(handle);
            }
            handle = other.handle;
            other.handle = nullptr;
        }
        return *this;
    }

    HKEY get() const { return handle; }
    HKEY* put() { return &handle; }

private:
    HKEY handle = nullptr;
};

HRESULT WriteRegistryStringValue(HKEY key, const wchar_t* valueName, const wchar_t* value) {
    const DWORD length = static_cast<DWORD>((wcslen(value) + 1) * sizeof(wchar_t));
    const LONG status = RegSetValueExW(key, valueName, 0, REG_SZ, reinterpret_cast<const BYTE*>(value), length);
    return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
}

// Sanitize a string for use as a registry key name
std::wstring SanitizeKeyName(const std::wstring& input) {
    std::wstring result;
    result.reserve(input.size());
    for (wchar_t c : input) {
        // Only allow alphanumeric, spaces, underscores, and hyphens
        if (std::iswalnum(static_cast<wint_t>(c)) || c == L' ' || c == L'_' || c == L'-') {
            result += c;
        }
    }
    // Replace spaces with underscores
    for (wchar_t& c : result) {
        if (c == L' ') {
            c = L'_';
        }
    }
    return result;
}

// Delete a registry key and all its subkeys
bool DeleteRegistryKeyRecursive(HKEY root, const std::wstring& keyPath) {
    LONG status = RegDeleteTreeW(root, keyPath.c_str());
    return (status == ERROR_SUCCESS || status == ERROR_FILE_NOT_FOUND);
}

// Enumerate all ShellTabs keys under a given path and delete them
bool ClearShellTabsKeysUnderPath(const std::wstring& basePath) {
    ScopedRegKey baseKey;
    LONG status = RegOpenKeyExW(HKEY_CURRENT_USER, basePath.c_str(), 0, KEY_READ | KEY_WRITE, baseKey.put());
    if (status == ERROR_FILE_NOT_FOUND) {
        return true;  // Nothing to clear
    }
    if (status != ERROR_SUCCESS) {
        LogMessage(LogLevel::Warning, L"ClearShellTabsKeysUnderPath: Failed to open %ls (0x%08X)",
                   basePath.c_str(), status);
        return false;
    }

    std::vector<std::wstring> keysToDelete;
    wchar_t keyName[256];
    DWORD keyNameLen;
    DWORD index = 0;

    // Enumerate all subkeys and find ones with our prefix
    while (true) {
        keyNameLen = ARRAYSIZE(keyName);
        status = RegEnumKeyExW(baseKey.get(), index, keyName, &keyNameLen, nullptr, nullptr, nullptr, nullptr);
        if (status == ERROR_NO_MORE_ITEMS) {
            break;
        }
        if (status != ERROR_SUCCESS) {
            LogMessage(LogLevel::Warning, L"ClearShellTabsKeysUnderPath: RegEnumKeyEx failed (0x%08X)", status);
            break;
        }

        // Check if this key starts with our prefix
        if (wcsncmp(keyName, kContextMenuKeyPrefix, wcslen(kContextMenuKeyPrefix)) == 0) {
            keysToDelete.push_back(keyName);
        }
        index++;
    }

    // Delete the found keys
    bool success = true;
    for (const auto& keyToDelete : keysToDelete) {
        std::wstring fullPath = basePath + L"\\" + keyToDelete;
        if (!DeleteRegistryKeyRecursive(HKEY_CURRENT_USER, fullPath)) {
            LogMessage(LogLevel::Warning, L"ClearShellTabsKeysUnderPath: Failed to delete %ls", fullPath.c_str());
            success = false;
        } else {
            LogMessage(LogLevel::Info, L"ClearShellTabsKeysUnderPath: Deleted %ls", fullPath.c_str());
        }
    }

    return success;
}

}  // namespace

std::wstring GetContextMenuRegistryPath(ContextMenuTarget target) {
    switch (target) {
        case ContextMenuTarget::kFiles:
            return L"Software\\Classes\\*\\shell";
        case ContextMenuTarget::kFolders:
            return L"Software\\Classes\\Directory\\shell";
        case ContextMenuTarget::kDrives:
            return L"Software\\Classes\\Drive\\shell";
        case ContextMenuTarget::kBackground:
            return L"Software\\Classes\\Directory\\Background\\shell";
        default:
            return L"";
    }
}

// Additional folder path for complete folder coverage
std::wstring GetContextMenuRegistryPathFolder() {
    return L"Software\\Classes\\Folder\\shell";
}

std::wstring GenerateContextMenuKeyName(const ContextMenuItem& item, size_t index) {
    std::wstring sanitizedLabel = SanitizeKeyName(item.label);
    if (sanitizedLabel.empty()) {
        sanitizedLabel = L"Item";
    }
    // Truncate if too long
    if (sanitizedLabel.length() > 32) {
        sanitizedLabel = sanitizedLabel.substr(0, 32);
    }

    // Add prefix and index for uniqueness
    wchar_t buffer[64];
    swprintf_s(buffer, L"%ls%ls_%zu", kContextMenuKeyPrefix, sanitizedLabel.c_str(), index);
    return buffer;
}

bool ShouldRegisterForTarget(const ContextMenuItem& item, ContextMenuTarget target) {
    if (!item.enabled) {
        return false;
    }

    switch (target) {
        case ContextMenuTarget::kFiles:
            return item.visibility.showForFiles;
        case ContextMenuTarget::kFolders:
        case ContextMenuTarget::kDrives:
            return item.visibility.showForFolders;
        case ContextMenuTarget::kBackground:
            // Background is typically for folder operations (like "Open Terminal Here")
            return item.visibility.showForFolders;
        default:
            return false;
    }
}

bool RegisterContextMenuCommand(
    const ContextMenuItem& item,
    const std::wstring& keyName,
    ContextMenuTarget target) {

    std::wstring basePath = GetContextMenuRegistryPath(target);
    if (basePath.empty()) {
        return false;
    }

    std::wstring keyPath = basePath + L"\\" + keyName;

    // Create the main key
    ScopedRegKey key;
    LONG status = RegCreateKeyExW(HKEY_CURRENT_USER, keyPath.c_str(), 0, nullptr,
                                   REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, key.put(), nullptr);
    if (status != ERROR_SUCCESS) {
        LogMessage(LogLevel::Warning, L"RegisterContextMenuCommand: Failed to create key %ls (0x%08X)",
                   keyPath.c_str(), status);
        return false;
    }

    // Set the display label
    if (!item.label.empty()) {
        WriteRegistryStringValue(key.get(), nullptr, item.label.c_str());
        WriteRegistryStringValue(key.get(), L"MUIVerb", item.label.c_str());
    }

    // Set the icon if specified
    if (!item.iconSource.empty()) {
        WriteRegistryStringValue(key.get(), L"Icon", item.iconSource.c_str());
    }

    // Handle submenus vs commands
    if (item.type == ContextMenuItemType::kSubmenu && !item.children.empty()) {
        // For submenus, set SubCommands value
        std::wstring subCommands;
        for (size_t i = 0; i < item.children.size(); ++i) {
            if (i > 0) subCommands += L";";
            std::wstring childKeyName = GenerateContextMenuKeyName(item.children[i], i);
            subCommands += childKeyName;
        }
        WriteRegistryStringValue(key.get(), L"SubCommands", subCommands.c_str());

        // Register each child command in the shell key
        for (size_t i = 0; i < item.children.size(); ++i) {
            std::wstring childKeyName = GenerateContextMenuKeyName(item.children[i], i);
            RegisterContextMenuCommand(item.children[i], childKeyName, target);
        }
    } else {
        // For commands, create the command subkey
        std::wstring commandKeyPath = keyPath + L"\\command";
        ScopedRegKey commandKey;
        status = RegCreateKeyExW(HKEY_CURRENT_USER, commandKeyPath.c_str(), 0, nullptr,
                                  REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, commandKey.put(), nullptr);
        if (status != ERROR_SUCCESS) {
            LogMessage(LogLevel::Warning, L"RegisterContextMenuCommand: Failed to create command key (0x%08X)", status);
            return false;
        }

        // Build the command string
        std::wstring command;
        if (!item.executable.empty()) {
            // Use executable and arguments separately
            command = L"\"" + item.executable + L"\"";
            if (!item.arguments.empty()) {
                command += L" " + item.arguments;
            }
        } else if (!item.commandTemplate.empty()) {
            // Legacy: use commandTemplate directly
            command = item.commandTemplate;
        }

        if (!command.empty()) {
            WriteRegistryStringValue(commandKey.get(), nullptr, command.c_str());
        }

        // Set runas if needed for admin elevation
        if (item.runAsAdmin) {
            WriteRegistryStringValue(key.get(), L"HasLUAShield", L"");
            // For actual elevation, we need to modify the command to use runas
            // or create a verb subkey with extended properties
        }
    }

    LogMessage(LogLevel::Info, L"RegisterContextMenuCommand: Registered %ls for target %d", keyName.c_str(), static_cast<int>(target));
    return true;
}

bool UnregisterContextMenuCommand(
    const std::wstring& keyName,
    ContextMenuTarget target) {

    std::wstring basePath = GetContextMenuRegistryPath(target);
    if (basePath.empty()) {
        return false;
    }

    std::wstring keyPath = basePath + L"\\" + keyName;
    return DeleteRegistryKeyRecursive(HKEY_CURRENT_USER, keyPath);
}

bool ClearContextMenuRegistry() {
    LogMessage(LogLevel::Info, L"ClearContextMenuRegistry: Clearing all ShellTabs context menu entries");

    bool success = true;

    // Clear from all targets
    const std::wstring paths[] = {
        GetContextMenuRegistryPath(ContextMenuTarget::kFiles),
        GetContextMenuRegistryPath(ContextMenuTarget::kFolders),
        GetContextMenuRegistryPath(ContextMenuTarget::kDrives),
        GetContextMenuRegistryPath(ContextMenuTarget::kBackground),
        GetContextMenuRegistryPathFolder(),  // Also clear from Folder\shell
    };

    for (const auto& path : paths) {
        if (!path.empty()) {
            if (!ClearShellTabsKeysUnderPath(path)) {
                success = false;
            }
        }
    }

    // Notify shell of the change
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);

    return success;
}

bool SyncContextMenuRegistry(const std::vector<ContextMenuItem>& items) {
    LogMessage(LogLevel::Info, L"SyncContextMenuRegistry: Syncing %zu context menu items", items.size());

    // First, clear all existing ShellTabs entries
    if (!ClearContextMenuRegistry()) {
        LogMessage(LogLevel::Warning, L"SyncContextMenuRegistry: Failed to clear existing entries");
    }

    // Now register each item to its appropriate targets
    bool success = true;
    size_t index = 0;

    for (const auto& item : items) {
        if (!item.enabled) {
            continue;
        }

        std::wstring keyName = GenerateContextMenuKeyName(item, index);

        // Register to Files target (*\shell)
        if (ShouldRegisterForTarget(item, ContextMenuTarget::kFiles)) {
            if (!RegisterContextMenuCommand(item, keyName, ContextMenuTarget::kFiles)) {
                success = false;
            }
        }

        // Register to Folders target (Directory\shell)
        if (ShouldRegisterForTarget(item, ContextMenuTarget::kFolders)) {
            if (!RegisterContextMenuCommand(item, keyName, ContextMenuTarget::kFolders)) {
                success = false;
            }

            // Also register to Folder\shell for complete coverage
            std::wstring folderPath = GetContextMenuRegistryPathFolder();
            std::wstring keyPath = folderPath + L"\\" + keyName;

            ScopedRegKey key;
            LONG status = RegCreateKeyExW(HKEY_CURRENT_USER, keyPath.c_str(), 0, nullptr,
                                           REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, key.put(), nullptr);
            if (status == ERROR_SUCCESS) {
                if (!item.label.empty()) {
                    WriteRegistryStringValue(key.get(), nullptr, item.label.c_str());
                    WriteRegistryStringValue(key.get(), L"MUIVerb", item.label.c_str());
                }
                if (!item.iconSource.empty()) {
                    WriteRegistryStringValue(key.get(), L"Icon", item.iconSource.c_str());
                }

                // Create command subkey
                std::wstring commandKeyPath = keyPath + L"\\command";
                ScopedRegKey commandKey;
                status = RegCreateKeyExW(HKEY_CURRENT_USER, commandKeyPath.c_str(), 0, nullptr,
                                          REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, commandKey.put(), nullptr);
                if (status == ERROR_SUCCESS) {
                    std::wstring command;
                    if (!item.executable.empty()) {
                        command = L"\"" + item.executable + L"\"";
                        if (!item.arguments.empty()) {
                            command += L" " + item.arguments;
                        }
                    } else if (!item.commandTemplate.empty()) {
                        command = item.commandTemplate;
                    }
                    if (!command.empty()) {
                        WriteRegistryStringValue(commandKey.get(), nullptr, command.c_str());
                    }
                }
            }
        }

        // Register to Drives target
        if (ShouldRegisterForTarget(item, ContextMenuTarget::kDrives)) {
            if (!RegisterContextMenuCommand(item, keyName, ContextMenuTarget::kDrives)) {
                success = false;
            }
        }

        // Register to Background target (right-click in empty space)
        if (ShouldRegisterForTarget(item, ContextMenuTarget::kBackground)) {
            if (!RegisterContextMenuCommand(item, keyName, ContextMenuTarget::kBackground)) {
                success = false;
            }
        }

        index++;
    }

    // Notify shell of the changes
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);

    LogMessage(LogLevel::Info, L"SyncContextMenuRegistry: Completed with %s", success ? L"success" : L"some failures");
    return success;
}

}  // namespace shelltabs
