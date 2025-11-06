#pragma once

#include <Windows.h>
#include <d2d1.h>
#include <dwrite.h>
#include <wincodec.h>
#include <ShlObj.h>
#include <vector>
#include <memory>
#include <string>
#include <unordered_map>

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "windowscodecs.lib")

// Forward declarations
class ExplorerGlowCoordinator;
struct SurfaceColorDescriptor;

namespace ShellTabs {

// Represents a single file or folder item in the view
struct FileListItem {
    std::wstring displayName;
    std::wstring fullPath;
    HICON icon = nullptr;
    IWICBitmapSource* thumbnail = nullptr;
    bool isFolder = false;
    bool isSelected = false;
    bool isHovered = false;
    FILETIME dateModified{};
    ULONGLONG fileSize = 0;
    RECT bounds{};  // Item position in view
    int itemIndex = -1;
    LPITEMIDLIST pidl = nullptr;  // Item identifier for shell operations
};

// View mode matching Explorer's view modes
enum class FileListViewMode {
    ExtraLargeIcons,
    LargeIcons,
    MediumIcons,
    SmallIcons,
    List,
    Details,
    Tiles,
    Content
};

// Custom window class that replaces DirectUIHWND
class CustomFileListView {
public:
    CustomFileListView();
    ~CustomFileListView();

    // Creates and registers the custom window class
    static bool RegisterWindowClass(HINSTANCE hInstance);

    // Creates an instance of the custom file list view
    HWND Create(DWORD dwExStyle, DWORD dwStyle, int x, int y,
                int width, int height, HWND hWndParent, HINSTANCE hInstance);

    // Associates with IShellView for shell integration
    bool AttachToShellView(IShellView* shellView);

    // Item management
    void AddItem(const FileListItem& item);
    void RemoveItem(int index);
    void UpdateItem(int index, const FileListItem& item);
    void ClearItems();
    void RefreshItems();  // Sync with shell view

    // View configuration
    void SetViewMode(FileListViewMode mode);
    FileListViewMode GetViewMode() const { return m_viewMode; }

    // Selection management
    void SelectItem(int index, bool addToSelection = false);
    void DeselectAll();
    std::vector<int> GetSelectedIndices() const;

    // Custom rendering control
    void SetBackgroundPaintCallback(
        bool (*callback)(HDC dc, HWND hwnd, const RECT& rect, void* context),
        void* context);

    void SetItemPaintCallback(
        bool (*callback)(HDC dc, const FileListItem& item, const RECT& rect, void* context),
        void* context);

    // Coordinator for theme integration
    void SetGlowCoordinator(ExplorerGlowCoordinator* coordinator);
    void SetColorDescriptor(const SurfaceColorDescriptor* descriptor);

    // Scrolling and layout
    void ScrollTo(int index);
    void EnsureVisible(int index);
    void InvalidateLayout();

    // Hit testing
    int HitTest(POINT pt) const;
    FileListItem* GetItem(int index);

    HWND GetHWND() const { return m_hwnd; }

private:
    // Window procedure for message handling
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

    // Message handlers
    LRESULT HandleCreate(HWND hwnd, LPCREATESTRUCT createStruct);
    LRESULT HandlePaint();
    LRESULT HandleSize(int width, int height);
    LRESULT HandleMouseMove(POINT pt);
    LRESULT HandleLeftButtonDown(POINT pt);
    LRESULT HandleLeftButtonUp(POINT pt);
    LRESULT HandleRightButtonDown(POINT pt);
    LRESULT HandleMouseWheel(int delta);
    LRESULT HandleKeyDown(WPARAM key);
    LRESULT HandleDestroy();

    // Rendering methods
    void InitializeD2DResources();
    void ReleaseD2DResources();
    void RenderBackground(HDC dc, const RECT& clientRect);
    void RenderItems(HDC dc, const RECT& clientRect);
    void RenderItem(HDC dc, const FileListItem& item, bool useCustomCallback);
    void RenderItemIcon(HDC dc, const FileListItem& item);
    void RenderItemText(HDC dc, const FileListItem& item);
    void RenderItemSelection(HDC dc, const FileListItem& item);
    void RenderItemThumbnail(HDC dc, const FileListItem& item);
    void RenderColumnsHeader(HDC dc, const RECT& clientRect);

    // Layout calculation
    void RecalculateLayout();
    void CalculateItemPositions();
    RECT GetItemRect(int index) const;
    SIZE GetItemSize() const;
    int GetItemsPerRow() const;
    int GetVisibleItemCount() const;

    // Scrolling
    void UpdateScrollbars();
    void DoScroll(int dx, int dy);

    // Shell integration
    void SyncWithShellView();
    void QueryShellViewItems();
    HICON GetShellIcon(LPITEMIDLIST pidl, bool large) const;
    IWICBitmapSource* GetShellThumbnail(LPITEMIDLIST pidl) const;

    // Selection handling
    void UpdateSelection(int index, bool ctrl, bool shift);
    void NotifySelectionChanged();

    // Context menu
    void ShowContextMenu(POINT pt, int itemIndex);

private:
    HWND m_hwnd = nullptr;
    HINSTANCE m_hInstance = nullptr;

    // View state
    std::vector<FileListItem> m_items;
    FileListViewMode m_viewMode = FileListViewMode::Details;
    int m_scrollY = 0;
    int m_scrollX = 0;
    int m_hoveredIndex = -1;
    int m_lastSelectedIndex = -1;
    POINT m_lastMousePos{};

    // Shell integration
    IShellView* m_shellView = nullptr;
    IShellFolder* m_shellFolder = nullptr;
    IShellFolderView* m_shellFolderView = nullptr;

    // Custom rendering
    bool (*m_backgroundPaintCallback)(HDC, HWND, const RECT&, void*) = nullptr;
    void* m_backgroundPaintContext = nullptr;
    bool (*m_itemPaintCallback)(HDC, const FileListItem&, const RECT&, void*) = nullptr;
    void* m_itemPaintContext = nullptr;

    // Theme integration
    ExplorerGlowCoordinator* m_coordinator = nullptr;
    const SurfaceColorDescriptor* m_colorDescriptor = nullptr;

    // Direct2D resources for advanced rendering
    ID2D1Factory* m_d2dFactory = nullptr;
    ID2D1HwndRenderTarget* m_renderTarget = nullptr;
    IDWriteFactory* m_dwriteFactory = nullptr;
    IDWriteTextFormat* m_textFormat = nullptr;
    IWICImagingFactory* m_wicFactory = nullptr;

    // Cached resources
    HFONT m_font = nullptr;
    int m_iconSize = 32;
    int m_itemSpacing = 8;
    bool m_layoutDirty = true;

    static constexpr const wchar_t* WINDOW_CLASS_NAME = L"ShellTabsFileListView";
};

// Global hook management
class DirectUIReplacementHook {
public:
    static bool Initialize();
    static void Shutdown();
    static bool IsEnabled() { return s_enabled; }

    // Register a custom file list view instance
    static void RegisterInstance(HWND hwnd, CustomFileListView* view);
    static void UnregisterInstance(HWND hwnd);
    static CustomFileListView* GetInstance(HWND hwnd);

private:
    static HWND WINAPI CreateWindowExW_Hook(
        DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
        DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
        HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam);

    static bool IsDirectUIClassName(LPCWSTR className);
    static HWND CreateReplacementWindow(DWORD dwExStyle, DWORD dwStyle,
                                       int X, int Y, int nWidth, int nHeight,
                                       HWND hWndParent, HINSTANCE hInstance);

    static bool s_enabled;
    static void* s_originalCreateWindowExW;
    static std::unordered_map<HWND, CustomFileListView*> s_instances;
};

} // namespace ShellTabs
