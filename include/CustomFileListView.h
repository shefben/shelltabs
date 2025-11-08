#pragma once

#include <Windows.h>
#include <d2d1.h>
#include <dwrite.h>
#include <wincodec.h>
#include <ShlObj.h>
#include <ShObjIdl.h>
#include <wrl/client.h>
#include <vector>
#include <memory>
#include <string>
#include <unordered_map>

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "windowscodecs.lib")

// Custom message for shell change notifications
#define WM_SHELL_NOTIFY (WM_USER + 100)

// Forward declarations
namespace shelltabs {
    class ExplorerGlowCoordinator;
    struct SurfaceColorDescriptor;
}

namespace ShellTabs {

// Column definitions for Details view
enum class ColumnType {
    Name,
    Size,
    Type,
    DateModified,
    DateCreated,
    Attributes
};

struct ColumnInfo {
    ColumnType type;
    std::wstring title;
    int width;
    bool visible;
};

// Sort state
struct SortState {
    ColumnType column = ColumnType::Name;
    bool ascending = true;
};

// Represents a single file or folder item in the view
struct FileListItem {
    std::wstring displayName;
    std::wstring fullPath;
    std::wstring fileType;  // File extension or "Folder"
    HICON icon = nullptr;
    IWICBitmapSource* thumbnail = nullptr;
    bool isFolder = false;
    bool isSelected = false;
    bool isHovered = false;
    FILETIME dateModified{};
    FILETIME dateCreated{};
    ULONGLONG fileSize = 0;
    DWORD attributes = 0;
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

    // Sorting and columns
    void SortBy(ColumnType column, bool ascending);
    void SortItems();
    void ToggleSort(ColumnType column);
    const SortState& GetSortState() const { return m_sortState; }
    void SetColumns(const std::vector<ColumnInfo>& columns);
    const std::vector<ColumnInfo>& GetColumns() const { return m_columns; }

    // Selection management
    void SelectItem(int index, bool addToSelection = false);
    void DeselectAll();
    std::vector<int> GetSelectedIndices() const;

    // Search and filter
    void SetFilter(const std::wstring& filter);
    void ClearFilter();
    const std::wstring& GetFilter() const { return m_filterText; }
    void BeginTypeAhead();
    void AddTypeAheadChar(wchar_t ch);
    void EndTypeAhead();

    // Custom rendering control
    void SetBackgroundPaintCallback(
        bool (*callback)(HDC dc, HWND hwnd, const RECT& rect, void* context),
        void* context);

    void SetItemPaintCallback(
        bool (*callback)(HDC dc, const FileListItem& item, const RECT& rect, void* context),
        void* context);

    // Coordinator for theme integration
    void SetGlowCoordinator(shelltabs::ExplorerGlowCoordinator* coordinator);
    void SetColorDescriptor(const shelltabs::SurfaceColorDescriptor* descriptor);

    // Scrolling and layout
    void ScrollTo(int index);
    void EnsureVisible(int index);
    void InvalidateLayout();

    // Hit testing
    int HitTest(POINT pt) const;
    FileListItem* GetItem(int index);

    HWND GetHWND() const { return m_hwnd; }

    // View mode synchronization
    void SyncViewModeFromShellView();

    // Shell change notifications
    void RegisterShellChangeNotify();
    void UnregisterShellChangeNotify();
    void OnShellChange(LONG eventId, LPITEMIDLIST pidl1, LPITEMIDLIST pidl2);

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
    LRESULT HandleLeftButtonDoubleClick(POINT pt);
    LRESULT HandleRightButtonDown(POINT pt);
    LRESULT HandleMouseWheel(int delta);
    LRESULT HandleKeyDown(WPARAM key);
    LRESULT HandleChar(WPARAM charCode);
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
    void RenderColumnHeader(HDC dc, const ColumnInfo& column, const RECT& rect, bool isSortColumn);
    void RenderSortArrow(HDC dc, const RECT& rect, bool ascending);
    void RenderThumbnailDirect2D(ID2D1RenderTarget* rt, const FileListItem& item, const RECT& rect);

    // Layout calculation
    void RecalculateLayout();
    void CalculateItemPositions();
    RECT GetItemRect(int index) const;
    SIZE GetItemSize() const;
    int GetItemsPerRow() const;
    int GetVisibleItemCount() const;
    int GetHeaderHeight() const { return m_viewMode == FileListViewMode::Details ? m_headerHeight : 0; }
    int HitTestColumn(POINT pt) const;
    RECT GetColumnRect(int columnIndex) const;

    // Scrolling
    void UpdateScrollbars();
    void DoScroll(int dx, int dy);

    // Sorting helpers
    static bool CompareItems(const FileListItem& a, const FileListItem& b, ColumnType column, bool ascending);
    static int CompareFileTime(const FILETIME& a, const FILETIME& b);
    static std::wstring FormatFileSize(ULONGLONG size);
    static std::wstring FormatFileTime(const FILETIME& ft);

    // Filter helpers
    void ApplyFilter();
    bool MatchesFilter(const FileListItem& item) const;
    void SelectFirstMatch();

    // Initialization
    void InitializeDefaultColumns();

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
    bool InvokeContextMenu(const std::vector<int>& itemIndices, POINT pt);

    // File operations
    void OpenItem(int index);
    void OpenSelectedItems();
    void DeleteSelectedItems();
    void RenameItem(int index);
    void BeginInPlaceRename(int index);
    void EndInPlaceRename(bool commit);
    void CopySelectedItems();
    void CutSelectedItems();
    void PasteItems();

    // Drag and drop
    void BeginDrag(int index);
    bool IsDragging() const { return m_isDragging; }

    // Clipboard operations
    bool HasClipboardData() const;

private:
    HWND m_hwnd = nullptr;
    HINSTANCE m_hInstance = nullptr;

    // View state
    std::vector<FileListItem> m_items;
    std::vector<FileListItem> m_filteredItems;  // Filtered view of m_items
    FileListViewMode m_viewMode = FileListViewMode::Details;
    int m_scrollY = 0;
    int m_scrollX = 0;
    int m_hoveredIndex = -1;
    int m_lastSelectedIndex = -1;
    POINT m_lastMousePos{};

    // Sorting and columns
    SortState m_sortState;
    std::vector<ColumnInfo> m_columns;
    int m_headerHeight = 25;
    int m_hoveredColumnIndex = -1;

    // Search and filter
    std::wstring m_filterText;
    bool m_hasFilter = false;
    std::wstring m_typeAheadText;
    UINT_PTR m_typeAheadTimer = 0;
    bool m_isTypeAhead = false;

    // Interaction state
    bool m_isDragging = false;
    int m_dragStartIndex = -1;
    POINT m_dragStartPos{};
    bool m_isRenaming = false;
    int m_renameIndex = -1;
    HWND m_renameEdit = nullptr;
    std::wstring m_renameOriginalText;

    // Shell integration
    IShellView* m_shellView = nullptr;
    IShellFolder* m_shellFolder = nullptr;
    IShellFolderView* m_shellFolderView = nullptr;
    ULONG m_shellChangeNotifyId = 0;
    std::wstring m_currentFolderPath;

    // Custom rendering
    bool (*m_backgroundPaintCallback)(HDC, HWND, const RECT&, void*) = nullptr;
    void* m_backgroundPaintContext = nullptr;
    bool (*m_itemPaintCallback)(HDC, const FileListItem&, const RECT&, void*) = nullptr;
    void* m_itemPaintContext = nullptr;

    // Theme integration
    shelltabs::ExplorerGlowCoordinator* m_coordinator = nullptr;
    const shelltabs::SurfaceColorDescriptor* m_colorDescriptor = nullptr;

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
    // Window creation hooks
    static HWND WINAPI CreateWindowExW_Hook(
        DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
        DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
        HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam);

    // Window search hooks
    static HWND WINAPI FindWindowW_Hook(
        LPCWSTR lpClassName, LPCWSTR lpWindowName);

    static HWND WINAPI FindWindowExW_Hook(
        HWND hWndParent, HWND hWndChildAfter, LPCWSTR lpClassName, LPCWSTR lpWindowName);

    static bool IsDirectUIClassName(LPCWSTR className);
    static HWND CreateReplacementWindow(DWORD dwExStyle, DWORD dwStyle,
                                       int X, int Y, int nWidth, int nHeight,
                                       HWND hWndParent, HINSTANCE hInstance);

    static bool s_enabled;
    static bool s_minHookAcquired;
    static void* s_originalCreateWindowExW;
    static void* s_originalFindWindowW;
    static void* s_originalFindWindowExW;
    static std::unordered_map<HWND, CustomFileListView*> s_instances;
};

} // namespace ShellTabs
