# DirectUI Replacement Architecture

## Overview

This document describes the new architecture for replacing Windows Explorer's native DirectUIHWND with a custom implementation that provides full control over the main panel view, including background rendering and file/folder item display.

## Architecture Goals

1. **Complete Control**: Replace the native DirectUIHWND window with our own custom window class
2. **Seamless Integration**: Maintain identical external behavior to DirectUIHWND
3. **Custom Rendering**: Full control over background painting and item rendering
4. **Shell Integration**: Maintain compatibility with IShellView and Explorer's shell interfaces
5. **Easy Customization**: Simple API for applying custom backgrounds, gradients, and item styles

## Opt-in requirement

The DirectUI replacement is **disabled by default** to keep Explorer stable while the custom view matures. To enable it for
experimentation, opt in through one of the following mechanisms (checked in this order):

1. Set the environment variable `SHELLTABS_ENABLE_DIRECTUI_REPLACEMENT` to `1`, `true`, or `on` before launching Explorer
2. Create/update `HKCU\Software\ShellTabs\EnableDirectUIReplacement` (DWORD) and set it to `1`
3. Add `directui_replacement|1` to `%AppData%\ShellTabs\options.db`

Any of these sources can also force the feature off by writing `0`, `false`, or `off`. When no opt-in is found the integration
initializes in a disabled state and the legacy DirectUI window is left untouched.

## Key Components

### 1. CustomFileListView (`include/CustomFileListView.h`, `src/CustomFileListView.cpp`)

The core custom window class that replaces DirectUIHWND.

**Features:**
- Custom window class registered as `"ShellTabsFileListView"`
- Complete message handling (WM_PAINT, WM_SIZE, WM_MOUSEMOVE, etc.)
- Multiple view modes (Details, Icons, List, Tiles, etc.)
- Direct2D/DirectWrite integration for advanced rendering
- Item management with full state tracking
- IShellView integration for file system synchronization

**Key Methods:**
```cpp
// Create the custom window
HWND Create(DWORD dwExStyle, DWORD dwStyle, int x, int y, int width, int height, HWND hWndParent, HINSTANCE hInstance);

// Attach to shell view for item synchronization
bool AttachToShellView(IShellView* shellView);

// Set custom rendering callbacks
void SetBackgroundPaintCallback(callback, context);
void SetItemPaintCallback(callback, context);

// Configure theme integration
void SetGlowCoordinator(ExplorerGlowCoordinator* coordinator);
void SetColorDescriptor(const SurfaceColorDescriptor* descriptor);

// Item management
void AddItem(const FileListItem& item);
void RemoveItem(int index);
void RefreshItems();

// View configuration
void SetViewMode(FileListViewMode mode);
void SelectItem(int index, bool addToSelection);
```

### 2. DirectUIReplacementHook (in CustomFileListView.cpp)

MinHook-based interceptor that replaces DirectUIHWND creation.

**How It Works:**
1. Hooks `CreateWindowExW` API using MinHook
2. Detects when DirectUIHWND is being created
3. Creates our `CustomFileListView` instead
4. Returns the custom window handle to the caller

**Key Methods:**
```cpp
static bool Initialize();           // Initialize the hook
static void Shutdown();            // Cleanup
static bool IsEnabled();           // Check if active
static void RegisterInstance(HWND hwnd, CustomFileListView* view);
static CustomFileListView* GetInstance(HWND hwnd);
```

### 3. DirectUIReplacementIntegration (`include/DirectUIReplacementIntegration.h`, `src/DirectUIReplacementIntegration.cpp`)

Integration layer that connects the custom file list view with CExplorerBHO.

**Features:**
- Initialization and lifecycle management
- Callback system for view creation notification
- Enable/disable toggle for the replacement system
- Helper functions for DirectUI window detection

**Key Methods:**
```cpp
static bool Initialize();                    // Initialize the system
static void Shutdown();                      // Cleanup
static void SetEnabled(bool enabled);        // Toggle on/off
static CustomFileListView* GetCustomViewForWindow(HWND hwnd);
static void SetCustomViewCreatedCallback(callback, context);
```

### 4. CExplorerBHO Integration

**Modified Files:**
- `include/CExplorerBHO.h` - Added forward declarations and member variables
- `src/CExplorerBHO.cpp` - Added initialization and configuration

**Integration Points:**

**Constructor Initialization:**
```cpp
CExplorerBHO::CExplorerBHO() {
    // Initialize DirectUI replacement system
    DirectUIReplacementIntegration::Initialize();

    // Set callback for when custom views are created
    DirectUIReplacementIntegration::SetCustomViewCreatedCallback(callback, this);
}
```

**OnCustomFileListViewCreated Callback:**
```cpp
void CExplorerBHO::OnCustomFileListViewCreated(CustomFileListView* view, HWND hwnd) {
    // Store the custom view instance
    m_customFileListView = view;
    m_directUiView = hwnd;

    // Configure with glow coordinator
    view->SetGlowCoordinator(&m_glowCoordinator);

    // Set color descriptor from glow surface
    view->SetColorDescriptor(&surface->GetDescriptor());

    // Attach to shell view
    view->AttachToShellView(m_shellView.Get());

    // Set background paint callback
    view->SetBackgroundPaintCallback(callback, this);

    // Register as glow surface
    RegisterGlowSurface(hwnd, ExplorerSurfaceKind::DirectUi, false);
}
```

## Data Flow

### Window Creation Flow

```
Explorer creates DirectUIHWND
    ↓
CreateWindowExW called
    ↓
DirectUIReplacementHook intercepts
    ↓
Is className == "DirectUIHWND"?
    ↓ Yes
Create CustomFileListView instead
    ↓
CustomFileListView::Create()
    ↓
Notify DirectUIReplacementIntegration
    ↓
Call view created callback
    ↓
CExplorerBHO::OnCustomFileListViewCreated()
    ↓
Configure view with coordinators, callbacks, etc.
    ↓
Return custom HWND to Explorer
```

### Rendering Flow

```
Windows sends WM_PAINT
    ↓
CustomFileListView::HandlePaint()
    ↓
Create double-buffered DC
    ↓
RenderBackground()
    ├─ Check custom callback
    ├─ Use glow coordinator colors
    └─ Paint gradient/solid
    ↓
RenderItems()
    ├─ Calculate visible items
    ├─ For each item:
    │   ├─ Check custom item callback
    │   ├─ Render selection background
    │   ├─ Render icon/thumbnail
    │   └─ Render text
    └─ Complete
    ↓
Blit to screen
```

### Item Synchronization Flow

```
AttachToShellView(IShellView*)
    ↓
Query IShellFolderView interface
    ↓
QueryShellViewItems()
    ↓
Enumerate items from shell view
    ↓
For each item:
    ├─ Get PIDL
    ├─ Get display name
    ├─ Get icon
    ├─ Get thumbnail (optional)
    ├─ Get file info
    └─ Create FileListItem
    ↓
Add to m_items vector
    ↓
Calculate layout
    ↓
Render items
```

## File Structure

### New Files Created

```
include/
├── CustomFileListView.h              # Main custom window class
├── DirectUIReplacementIntegration.h  # Integration layer

src/
├── CustomFileListView.cpp              # Implementation
└── DirectUIReplacementIntegration.cpp  # Integration implementation
```

### Modified Files

```
include/
└── CExplorerBHO.h                     # Added forward declarations, member variables

src/
├── CExplorerBHO.cpp                   # Added initialization, OnCustomFileListViewCreated
└── CMakeLists.txt                     # Added new source files, dwrite library
```

## Build Configuration

### CMakeLists.txt Changes

**Added Source Files:**
```cmake
src/CustomFileListView.cpp
src/DirectUIReplacementIntegration.cpp
```

**Added Libraries:**
```cmake
dwrite  # DirectWrite for text rendering
```

## Custom Rendering API

### Background Customization

You can provide a custom background paint callback:

```cpp
view->SetBackgroundPaintCallback(
    [](HDC dc, HWND hwnd, const RECT& rect, void* context) -> bool {
        // Paint custom background
        // Return true if handled, false to use default

        // Example: Custom gradient
        TRIVERTEX vertices[2] = {
            { rect.left, rect.top, 0xFF00, 0x8000, 0x0000, 0xFF00 },
            { rect.right, rect.bottom, 0x0000, 0x0000, 0xFF00, 0xFF00 }
        };
        GRADIENT_RECT gradRect = { 0, 1 };
        GdiGradientFill(dc, vertices, 2, &gradRect, 1, GRADIENT_FILL_RECT_V);
        return true;
    },
    context
);
```

### Item Customization

Provide a custom item paint callback:

```cpp
view->SetItemPaintCallback(
    [](HDC dc, const FileListItem& item, const RECT& rect, void* context) -> bool {
        // Paint custom item
        // Return true if handled, false to use default

        // Example: Custom selection highlight
        if (item.isSelected) {
            HBRUSH brush = CreateSolidBrush(RGB(100, 200, 255));
            FillRect(dc, &rect, brush);
            DeleteObject(brush);
        }

        // Draw icon
        if (item.icon) {
            DrawIconEx(dc, rect.left + 4, rect.top + 4,
                      item.icon, 16, 16, 0, nullptr, DI_NORMAL);
        }

        // Draw text
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, RGB(0, 0, 0));
        DrawTextW(dc, item.displayName.c_str(), -1, &rect,
                 DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        return true;
    },
    context
);
```

### Theme Integration

The custom view integrates with the existing ExplorerGlowCoordinator and SurfaceColorDescriptor system:

```cpp
// Automatically configured by CExplorerBHO
view->SetGlowCoordinator(&m_glowCoordinator);
view->SetColorDescriptor(&surface->GetDescriptor());

// The view will use these for:
// - Background colors (gradient or solid)
// - Text colors
// - Selection highlight colors
```

## Advanced Features

### View Modes

The custom view supports all standard Explorer view modes:

```cpp
enum class FileListViewMode {
    ExtraLargeIcons,  // 256x256 icons
    LargeIcons,       // 96x96 icons
    MediumIcons,      // 48x48 icons
    SmallIcons,       // 16x16 icons
    List,             // List view
    Details,          // Details view with columns
    Tiles,            // Tile view
    Content           // Content view
};

view->SetViewMode(FileListViewMode::Details);
```

### Selection Management

```cpp
// Single selection
view->SelectItem(index, false);

// Multi-selection
view->SelectItem(index1, false);
view->SelectItem(index2, true);  // Add to selection

// Get selected items
auto indices = view->GetSelectedIndices();

// Deselect all
view->DeselectAll();
```

### Scrolling and Navigation

```cpp
// Scroll to specific item
view->ScrollTo(index);

// Ensure item is visible
view->EnsureVisible(index);

// Programmatic scrolling
view->DoScroll(dx, dy);
```

## Enable/Disable the Replacement

The replacement system can be toggled:

```cpp
// Disable replacement (use native DirectUIHWND)
DirectUIReplacementIntegration::SetEnabled(false);

// Enable replacement (use custom view)
DirectUIReplacementIntegration::SetEnabled(true);

// Check status
bool enabled = DirectUIReplacementIntegration::IsEnabled();
```

## Future Enhancements

### Planned Features

1. **Shell Item Enumeration**: Complete IShellFolderView integration for automatic item synchronization
2. **Thumbnail Support**: Full WIC-based thumbnail rendering for large icon modes
3. **Context Menu Integration**: IContextMenu support for right-click operations
4. **Drag & Drop**: IDropTarget implementation for file operations
5. **Column Headers**: Sortable column headers for Details view
6. **Grouping**: Item grouping support
7. **Search Integration**: Search box and filtering
8. **Performance Optimization**: Virtual scrolling for large directories
9. **Accessibility**: UI Automation support

### Implementation Roadmap

**Phase 1 (Current)**: ✅ Complete
- Basic window replacement
- Custom rendering
- MinHook integration
- CExplorerBHO integration
- Build system configuration

**Phase 2**: Item Synchronization
- Implement QueryShellViewItems()
- IShellFolderView enumeration
- Shell notification handling
- Icon and thumbnail loading

**Phase 3**: User Interaction
- Context menus
- Drag and drop
- Rename operations
- File operations

**Phase 4**: Advanced Features
- Column sorting
- Grouping
- Search/filter
- Virtual scrolling

## Debugging and Troubleshooting

### Logging

The system uses ShellTabs' logging infrastructure:

```cpp
LogMessage(LogLevel::Info, L"Custom file list view created (hwnd=%p)", hwnd);
LogMessage(LogLevel::Warning, L"Failed to initialize DirectUI replacement system");
```

Check the ShellTabs log file for diagnostic information.

### Common Issues

**Issue**: Custom view not being created
- Check that `DirectUIReplacementIntegration::Initialize()` succeeded
- Verify MinHook initialization didn't fail
- Check that `SetEnabled(true)` was called

**Issue**: Items not appearing
- Verify `AttachToShellView()` was called successfully
- Check that `QueryShellViewItems()` is implemented
- Ensure `RefreshItems()` is being called

**Issue**: Rendering glitches
- Check that double-buffering is working (memDC created successfully)
- Verify Direct2D resources initialized properly
- Check for WM_SIZE handling to resize render target

## Testing

### Manual Testing Steps

1. Build the project with the new files
2. Register the ShellTabs.dll
3. Open Windows Explorer
4. Navigate to a folder
5. Verify custom file list view is created (check logs)
6. Test view modes (Details, Icons, etc.)
7. Test selection (single, multi, keyboard)
8. Test scrolling (mouse wheel, keyboard)
9. Test custom backgrounds (if configured)

### Automated Testing

Add tests for:
- Window creation hook interception
- Item management (add, remove, update)
- Layout calculation
- Hit testing
- Selection management

## Performance Considerations

### Optimizations

1. **Double Buffering**: All rendering uses off-screen DC to prevent flicker
2. **Visible Item Culling**: Only render items in visible area
3. **Lazy Loading**: Icons and thumbnails loaded on demand
4. **Layout Caching**: Layout calculations cached until invalidated
5. **Direct2D**: Hardware-accelerated rendering for advanced effects

### Benchmarks

(To be measured in Phase 2)
- Item rendering time
- Scroll performance
- Memory usage for large directories
- Thumbnail loading performance

## API Reference

See header files for complete API documentation:
- `include/CustomFileListView.h` - Main custom window class
- `include/DirectUIReplacementIntegration.h` - Integration layer

## License and Credits

Part of the ShellTabs project.

MinHook library: https://github.com/TsudaKageyu/minhook
