#pragma once

#include <windows.h>

namespace shelltabs {

// Attempts to notify the user that ShellTabs automation was disabled by policy.
// Returns true when a toast notification was successfully queued, false otherwise.
bool NotifyAutomationDisabledByPolicy(HRESULT hr) noexcept;

}  // namespace shelltabs
