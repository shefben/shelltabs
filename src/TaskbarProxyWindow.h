#pragma once

#include <windows.h>

#include "TabManager.h"

namespace shelltabs {

class TaskbarTabController;

struct TaskbarProxyConfig {
    TaskbarTabController* controller = nullptr;
    TabLocation location;
    HWND owner = nullptr;
};

HWND CreateTaskbarProxyWindow(const TaskbarProxyConfig& config);
void DestroyTaskbarProxyWindow(HWND hwnd);
void UpdateTaskbarProxyLocation(HWND hwnd, const TabLocation& location);

}  // namespace shelltabs

