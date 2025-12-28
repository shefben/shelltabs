#pragma once

#include <windows.h>
#include <string>
#include <vector>

#include "OptionsStore.h"

namespace shelltabs {

// Context Menu Registry Target
enum class ContextMenuTarget {
    kFiles = 0,      // Software\Classes\*\shell
    kFolders,        // Software\Classes\Directory\shell and Software\Classes\Folder\shell
    kDrives,         // Software\Classes\Drive\shell
    kBackground,     // Software\Classes\Directory\Background\shell (folder background)
};

// Registry key prefix for ShellTabs context menu items
constexpr wchar_t kContextMenuKeyPrefix[] = L"ShellTabs.";

// Synchronize context menu items from options to registry
// This removes old ShellTabs entries and registers new ones based on the options
bool SyncContextMenuRegistry(const std::vector<ContextMenuItem>& items);

// Remove all ShellTabs context menu entries from the registry
bool ClearContextMenuRegistry();

// Register a single context menu command
bool RegisterContextMenuCommand(
    const ContextMenuItem& item,
    const std::wstring& keyName,
    ContextMenuTarget target);

// Unregister a single context menu command
bool UnregisterContextMenuCommand(
    const std::wstring& keyName,
    ContextMenuTarget target);

// Get the registry path for a context menu target
std::wstring GetContextMenuRegistryPath(ContextMenuTarget target);

// Generate a unique key name for a context menu item
std::wstring GenerateContextMenuKeyName(const ContextMenuItem& item, size_t index);

// Check if an item should be registered for a given target
bool ShouldRegisterForTarget(const ContextMenuItem& item, ContextMenuTarget target);

}  // namespace shelltabs
