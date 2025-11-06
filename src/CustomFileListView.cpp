#include "CustomFileListView.h"
#include "ExplorerGlowSurfaces.h"
#include "MinHook.h"
#include "Logging.h"
#include "DirectUIReplacementIntegration.h"
#include <algorithm>
#include <mutex>
#include <Shlwapi.h>
#include <ShlObj.h>
#include <ShObjIdl.h>
#include <commoncontrols.h>
#include <wrl/client.h>
#include <windowsx.h>

#pragma comment(lib, "Shlwapi.lib")

namespace ShellTabs {

using shelltabs::LogMessage;
using shelltabs::LogLevel;

namespace {
    std::mutex g_instanceMutex;
}

// ============================================================================
// CustomFileListView Implementation
// ============================================================================

CustomFileListView::CustomFileListView() {
    InitializeDefaultColumns();
}

CustomFileListView::~CustomFileListView() {
    ReleaseD2DResources();

    if (m_font) {
        DeleteObject(m_font);
    }

    // Release shell interfaces
    if (m_shellFolderView) m_shellFolderView->Release();
    if (m_shellFolder) m_shellFolder->Release();
    if (m_shellView) m_shellView->Release();

    // Clean up items
    for (auto& item : m_items) {
        if (item.icon) DestroyIcon(item.icon);
        if (item.thumbnail) item.thumbnail->Release();
        if (item.pidl) CoTaskMemFree(item.pidl);
    }
}

bool CustomFileListView::RegisterWindowClass(HINSTANCE hInstance) {
    WNDCLASSEXW wcex = {};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wcex.lpfnWndProc = WindowProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = sizeof(CustomFileListView*);
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = nullptr;  // We handle all painting
    wcex.lpszClassName = WINDOW_CLASS_NAME;

    return RegisterClassExW(&wcex) != 0;
}

HWND CustomFileListView::Create(DWORD dwExStyle, DWORD dwStyle, int x, int y,
                                 int width, int height, HWND hWndParent,
                                 HINSTANCE hInstance) {
    m_hInstance = hInstance;

    m_hwnd = CreateWindowExW(
        dwExStyle,
        WINDOW_CLASS_NAME,
        L"",
        dwStyle | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL,
        x, y, width, height,
        hWndParent,
        nullptr,
        hInstance,
        this  // Pass 'this' pointer via lpParam
    );

    if (m_hwnd) {
        // Notify integration layer that view was created
        shelltabs::DirectUIReplacementIntegration::NotifyViewCreated(this, m_hwnd);
    }

    return m_hwnd;
}

LRESULT CALLBACK CustomFileListView::WindowProc(HWND hwnd, UINT msg,
                                                 WPARAM wParam, LPARAM lParam) {
    CustomFileListView* pThis = nullptr;

    if (msg == WM_NCCREATE) {
        auto* createStruct = reinterpret_cast<LPCREATESTRUCT>(lParam);
        pThis = static_cast<CustomFileListView*>(createStruct->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
    } else {
        pThis = reinterpret_cast<CustomFileListView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (pThis) {
        switch (msg) {
            case WM_CREATE:
                return pThis->HandleCreate(hwnd, reinterpret_cast<LPCREATESTRUCT>(lParam));
            case WM_PAINT:
                return pThis->HandlePaint();
            case WM_SIZE:
                return pThis->HandleSize(LOWORD(lParam), HIWORD(lParam));
            case WM_MOUSEMOVE:
                return pThis->HandleMouseMove({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
            case WM_LBUTTONDOWN:
                return pThis->HandleLeftButtonDown({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
            case WM_LBUTTONUP:
                return pThis->HandleLeftButtonUp({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
            case WM_LBUTTONDBLCLK:
                return pThis->HandleLeftButtonDoubleClick({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
            case WM_RBUTTONDOWN:
                return pThis->HandleRightButtonDown({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
            case WM_MOUSEWHEEL:
                return pThis->HandleMouseWheel(GET_WHEEL_DELTA_WPARAM(wParam));
            case WM_KEYDOWN:
                return pThis->HandleKeyDown(wParam);
            case WM_CHAR:
                return pThis->HandleChar(wParam);
            case WM_ERASEBKGND:
                return 1;  // We handle all painting
            case WM_SHELL_NOTIFY:
                {
                    // Handle shell change notification
                    LONG eventId;
                    LPITEMIDLIST* ppidl;
                    HANDLE hNotifyLock = SHChangeNotification_Lock(
                        reinterpret_cast<HANDLE>(wParam),
                        static_cast<DWORD>(lParam),
                        &ppidl,
                        &eventId
                    );

                    if (hNotifyLock) {
                        pThis->OnShellChange(eventId, ppidl[0], ppidl[1]);
                        SHChangeNotification_Unlock(hNotifyLock);
                    }
                    return 0;
                }
            case WM_TIMER:
                if (wParam == pThis->m_typeAheadTimer) {
                    // Type-ahead timeout - end type-ahead mode
                    pThis->EndTypeAhead();
                    return 0;
                }
                break;
            case WM_DESTROY:
                return pThis->HandleDestroy();
        }
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

LRESULT CustomFileListView::HandleCreate(HWND hwnd, LPCREATESTRUCT createStruct) {
    UNREFERENCED_PARAMETER(createStruct);
    m_hwnd = hwnd;

    // Initialize Direct2D resources
    InitializeD2DResources();

    // Create default font
    NONCLIENTMETRICSW ncm = { sizeof(ncm) };
    if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0)) {
        m_font = CreateFontIndirectW(&ncm.lfMessageFont);
    }

    return 0;
}

LRESULT CustomFileListView::HandlePaint() {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(m_hwnd, &ps);

    if (hdc) {
        RECT clientRect;
        GetClientRect(m_hwnd, &clientRect);

        // Create memory DC for double buffering
        HDC memDC = CreateCompatibleDC(hdc);
        HBITMAP memBitmap = CreateCompatibleBitmap(hdc,
            clientRect.right - clientRect.left,
            clientRect.bottom - clientRect.top);
        HBITMAP oldBitmap = (HBITMAP)SelectObject(memDC, memBitmap);

        // Render to memory DC
        RenderBackground(memDC, clientRect);
        RenderItems(memDC, clientRect);

        // Blit to screen
        BitBlt(hdc, 0, 0,
            clientRect.right - clientRect.left,
            clientRect.bottom - clientRect.top,
            memDC, 0, 0, SRCCOPY);

        // Cleanup
        SelectObject(memDC, oldBitmap);
        DeleteObject(memBitmap);
        DeleteDC(memDC);

        EndPaint(m_hwnd, &ps);
    }

    return 0;
}

LRESULT CustomFileListView::HandleSize(int width, int height) {
    // Resize render target
    if (m_renderTarget) {
        m_renderTarget->Resize(D2D1::SizeU(width, height));
    }

    InvalidateLayout();
    UpdateScrollbars();

    return 0;
}

LRESULT CustomFileListView::HandleMouseMove(POINT pt) {
    int hitIndex = HitTest(pt);

    if (hitIndex != m_hoveredIndex) {
        int oldHovered = m_hoveredIndex;
        m_hoveredIndex = hitIndex;

        // Invalidate old and new hovered items
        if (oldHovered >= 0 && oldHovered < static_cast<int>(m_items.size())) {
            m_items[oldHovered].isHovered = false;
            InvalidateRect(m_hwnd, &m_items[oldHovered].bounds, FALSE);
        }
        if (m_hoveredIndex >= 0 && m_hoveredIndex < static_cast<int>(m_items.size())) {
            m_items[m_hoveredIndex].isHovered = true;
            InvalidateRect(m_hwnd, &m_items[m_hoveredIndex].bounds, FALSE);
        }
    }

    m_lastMousePos = pt;
    return 0;
}

LRESULT CustomFileListView::HandleLeftButtonDown(POINT pt) {
    SetFocus(m_hwnd);

    // Check for column header click in Details view
    if (m_viewMode == FileListViewMode::Details && pt.y < m_headerHeight) {
        int columnIndex = HitTestColumn(pt);
        if (columnIndex >= 0 && columnIndex < static_cast<int>(m_columns.size())) {
            ToggleSort(m_columns[columnIndex].type);
            return 0;
        }
    }

    int hitIndex = HitTest(pt);
    if (hitIndex >= 0) {
        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        UpdateSelection(hitIndex, ctrl, shift);
    } else {
        // Clicked empty space - deselect all
        DeselectAll();
    }

    return 0;
}

LRESULT CustomFileListView::HandleLeftButtonUp(POINT pt) {
    UNREFERENCED_PARAMETER(pt);
    if (m_isDragging) {
        // End drag operation
        m_isDragging = false;
        ReleaseCapture();
    }
    return 0;
}

LRESULT CustomFileListView::HandleLeftButtonDoubleClick(POINT pt) {
    int hitIndex = HitTest(pt);
    if (hitIndex >= 0) {
        OpenItem(hitIndex);
    }
    return 0;
}

LRESULT CustomFileListView::HandleRightButtonDown(POINT pt) {
    int hitIndex = HitTest(pt);

    // If right-clicked item is not selected, select only that item
    if (hitIndex >= 0 && !m_items[hitIndex].isSelected) {
        UpdateSelection(hitIndex, false, false);
    }

    // Show context menu
    ShowContextMenu(pt, hitIndex);

    return 0;
}

LRESULT CustomFileListView::HandleMouseWheel(int delta) {
    // Scroll by 3 items per wheel notch
    SIZE itemSize = GetItemSize();
    int scrollAmount = -delta / WHEEL_DELTA * 3 * (itemSize.cy + m_itemSpacing);

    DoScroll(0, scrollAmount);

    return 0;
}

LRESULT CustomFileListView::HandleKeyDown(WPARAM key) {
    bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    // Handle special keys first
    switch (key) {
        case VK_RETURN:
            // Open selected items
            if (!ctrl && !shift) {
                OpenSelectedItems();
                return 0;
            }
            break;

        case VK_DELETE:
            // Delete selected items
            if (!ctrl && !shift) {
                DeleteSelectedItems();
                return 0;
            }
            break;

        case VK_F2:
            // Rename selected item
            if (!ctrl && !shift && m_lastSelectedIndex >= 0) {
                BeginInPlaceRename(m_lastSelectedIndex);
                return 0;
            }
            break;

        case VK_ESCAPE:
            // Cancel rename if active
            if (m_isRenaming) {
                EndInPlaceRename(false);
                return 0;
            }
            // Otherwise deselect all
            DeselectAll();
            return 0;

        case 'A':
            // Select all (Ctrl+A)
            if (ctrl && !shift) {
                for (size_t i = 0; i < m_items.size(); ++i) {
                    m_items[i].isSelected = true;
                }
                InvalidateRect(m_hwnd, nullptr, FALSE);
                return 0;
            }
            break;

        case 'C':
            // Copy (Ctrl+C)
            if (ctrl && !shift) {
                CopySelectedItems();
                return 0;
            }
            break;

        case 'X':
            // Cut (Ctrl+X)
            if (ctrl && !shift) {
                CutSelectedItems();
                return 0;
            }
            break;

        case 'V':
            // Paste (Ctrl+V)
            if (ctrl && !shift) {
                PasteItems();
                return 0;
            }
            break;
    }

    // Handle navigation keys
    int selectedIndex = m_lastSelectedIndex;
    if (selectedIndex < 0 && !m_items.empty()) {
        selectedIndex = 0;
    }

    int newIndex = selectedIndex;
    int itemsPerRow = GetItemsPerRow();

    switch (key) {
        case VK_UP:
            newIndex = std::max(0, selectedIndex - itemsPerRow);
            break;
        case VK_DOWN:
            newIndex = std::min(static_cast<int>(m_items.size()) - 1,
                              selectedIndex + itemsPerRow);
            break;
        case VK_LEFT:
            newIndex = std::max(0, selectedIndex - 1);
            break;
        case VK_RIGHT:
            newIndex = std::min(static_cast<int>(m_items.size()) - 1,
                              selectedIndex + 1);
            break;
        case VK_HOME:
            newIndex = 0;
            break;
        case VK_END:
            newIndex = static_cast<int>(m_items.size()) - 1;
            break;
        case VK_PRIOR:  // Page Up
            newIndex = std::max(0, selectedIndex - GetVisibleItemCount());
            break;
        case VK_NEXT:  // Page Down
            newIndex = std::min(static_cast<int>(m_items.size()) - 1,
                              selectedIndex + GetVisibleItemCount());
            break;
        default:
            return 0;
    }

    if (newIndex != selectedIndex) {
        UpdateSelection(newIndex, ctrl, shift);
        EnsureVisible(newIndex);
    }

    return 0;
}

LRESULT CustomFileListView::HandleChar(WPARAM charCode) {
    // Ignore characters during rename
    if (m_isRenaming) {
        return 0;
    }

    // Ignore control characters except backspace
    if (charCode < 32 && charCode != VK_BACK) {
        return 0;
    }

    // Type-ahead find
    if (charCode == VK_BACK) {
        // Backspace - remove last character
        if (!m_typeAheadText.empty()) {
            m_typeAheadText.pop_back();
            if (m_typeAheadText.empty()) {
                EndTypeAhead();
            } else {
                SelectFirstMatch();
            }
        }
    } else {
        // Add character to search
        if (!m_isTypeAhead) {
            BeginTypeAhead();
        }
        AddTypeAheadChar(static_cast<wchar_t>(charCode));
    }

    return 0;
}

LRESULT CustomFileListView::HandleDestroy() {
    UnregisterShellChangeNotify();
    DirectUIReplacementHook::UnregisterInstance(m_hwnd);
    return 0;
}

void CustomFileListView::InitializeD2DResources() {
    // Create D2D factory
    D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &m_d2dFactory);

    // Create DWrite factory
    DWriteCreateFactory(
        DWRITE_FACTORY_TYPE_SHARED,
        __uuidof(IDWriteFactory),
        reinterpret_cast<IUnknown**>(&m_dwriteFactory)
    );

    // Create WIC factory
    CoCreateInstance(
        CLSID_WICImagingFactory,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&m_wicFactory)
    );

    // Create render target (will be created in WM_SIZE)
    if (m_d2dFactory && m_hwnd) {
        RECT rc;
        GetClientRect(m_hwnd, &rc);

        m_d2dFactory->CreateHwndRenderTarget(
            D2D1::RenderTargetProperties(),
            D2D1::HwndRenderTargetProperties(
                m_hwnd,
                D2D1::SizeU(rc.right - rc.left, rc.bottom - rc.top)
            ),
            &m_renderTarget
        );
    }

    // Create text format
    if (m_dwriteFactory) {
        m_dwriteFactory->CreateTextFormat(
            L"Segoe UI",
            nullptr,
            DWRITE_FONT_WEIGHT_NORMAL,
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            9.0f,
            L"en-us",
            &m_textFormat
        );
    }
}

void CustomFileListView::ReleaseD2DResources() {
    if (m_textFormat) m_textFormat->Release();
    if (m_renderTarget) m_renderTarget->Release();
    if (m_wicFactory) m_wicFactory->Release();
    if (m_dwriteFactory) m_dwriteFactory->Release();
    if (m_d2dFactory) m_d2dFactory->Release();
}

void CustomFileListView::RenderBackground(HDC dc, const RECT& clientRect) {
    // Use custom callback if provided
    if (m_backgroundPaintCallback && m_backgroundPaintCallback(dc, m_hwnd, clientRect, m_backgroundPaintContext)) {
        return;
    }

    // Use glow coordinator colors if available
    if (m_colorDescriptor && m_colorDescriptor->fillOverride) {
        // Paint gradient or solid color based on descriptor
        const auto& colors = m_colorDescriptor->fillColors;
        if (colors.gradient && colors.start != colors.end) {
            TRIVERTEX vertices[2] = {
                { clientRect.left, clientRect.top,
                  static_cast<COLOR16>(GetRValue(colors.start) << 8),
                  static_cast<COLOR16>(GetGValue(colors.start) << 8),
                  static_cast<COLOR16>(GetBValue(colors.start) << 8),
                  static_cast<COLOR16>(0xFF00) },
                { clientRect.right, clientRect.bottom,
                  static_cast<COLOR16>(GetRValue(colors.end) << 8),
                  static_cast<COLOR16>(GetGValue(colors.end) << 8),
                  static_cast<COLOR16>(GetBValue(colors.end) << 8),
                  static_cast<COLOR16>(0xFF00) }
            };
            GRADIENT_RECT gradRect = { 0, 1 };
            GdiGradientFill(dc, vertices, 2, &gradRect, 1, GRADIENT_FILL_RECT_V);
        } else {
            HBRUSH brush = CreateSolidBrush(colors.start);
            FillRect(dc, &clientRect, brush);
            DeleteObject(brush);
        }
        return;
    }

    // Default: white background
    HBRUSH brush = (HBRUSH)GetStockObject(WHITE_BRUSH);
    FillRect(dc, &clientRect, brush);
}

void CustomFileListView::RenderItems(HDC dc, const RECT& clientRect) {
    if (m_layoutDirty) {
        RecalculateLayout();
    }

    // Render column headers in details view
    if (m_viewMode == FileListViewMode::Details) {
        RenderColumnsHeader(dc, clientRect);
    }

    // Render each visible item
    for (auto& item : m_items) {
        // Check if item is visible in current scroll position
        RECT itemRect = item.bounds;
        OffsetRect(&itemRect, -m_scrollX, -m_scrollY);

        RECT intersectRect;
        if (IntersectRect(&intersectRect, &itemRect, &clientRect)) {
            RenderItem(dc, item, m_itemPaintCallback != nullptr);
        }
    }
}

void CustomFileListView::RenderItem(HDC dc, const FileListItem& item, bool useCustomCallback) {
    RECT itemRect = item.bounds;
    OffsetRect(&itemRect, -m_scrollX, -m_scrollY);

    // Use custom callback if provided
    if (useCustomCallback && m_itemPaintCallback) {
        if (m_itemPaintCallback(dc, item, itemRect, m_itemPaintContext)) {
            return;
        }
    }

    // Render selection background
    if (item.isSelected || item.isHovered) {
        RenderItemSelection(dc, item);
    }

    // Render icon/thumbnail
    if (m_viewMode == FileListViewMode::ExtraLargeIcons ||
        m_viewMode == FileListViewMode::LargeIcons) {
        RenderItemThumbnail(dc, item);
    } else {
        RenderItemIcon(dc, item);
    }

    // Render text
    RenderItemText(dc, item);
}

void CustomFileListView::RenderItemIcon(HDC dc, const FileListItem& item) {
    if (!item.icon) return;

    RECT iconRect = item.bounds;
    OffsetRect(&iconRect, -m_scrollX, -m_scrollY);

    // Center icon vertically
    int iconY = iconRect.top + (iconRect.bottom - iconRect.top - m_iconSize) / 2;
    DrawIconEx(dc, iconRect.left + 4, iconY, item.icon, m_iconSize, m_iconSize, 0, nullptr, DI_NORMAL);
}

void CustomFileListView::RenderItemText(HDC dc, const FileListItem& item) {
    RECT textRect = item.bounds;
    OffsetRect(&textRect, -m_scrollX, -m_scrollY);

    // Offset for icon
    textRect.left += m_iconSize + 8;

    // Set text colors
    SetBkMode(dc, TRANSPARENT);
    if (item.isSelected) {
        SetTextColor(dc, GetSysColor(COLOR_HIGHLIGHTTEXT));
    } else if (m_colorDescriptor && m_colorDescriptor->textColor != CLR_INVALID) {
        SetTextColor(dc, m_colorDescriptor->textColor);
    } else {
        SetTextColor(dc, GetSysColor(COLOR_WINDOWTEXT));
    }

    // Select font
    HFONT oldFont = (HFONT)SelectObject(dc, m_font);

    // Draw text
    DrawTextW(dc, item.displayName.c_str(), -1, &textRect,
             DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);

    SelectObject(dc, oldFont);
}

void CustomFileListView::RenderItemSelection(HDC dc, const FileListItem& item) {
    RECT selRect = item.bounds;
    OffsetRect(&selRect, -m_scrollX, -m_scrollY);

    COLORREF color;
    if (item.isSelected) {
        color = GetSysColor(COLOR_HIGHLIGHT);
    } else if (item.isHovered) {
        color = RGB(229, 243, 255);  // Light blue hover
    } else {
        return;
    }

    HBRUSH brush = CreateSolidBrush(color);
    FillRect(dc, &selRect, brush);
    DeleteObject(brush);
}

void CustomFileListView::RenderItemThumbnail(HDC dc, const FileListItem& item) {
    // Thumbnail rendering for large icon modes
    // Implementation would use WIC to render thumbnail
    RenderItemIcon(dc, item);  // Fallback to icon for now
}

void CustomFileListView::RenderColumnsHeader(HDC dc, const RECT& clientRect) {
    if (m_viewMode != FileListViewMode::Details) return;

    RECT headerRect = clientRect;
    headerRect.bottom = headerRect.top + m_headerHeight;

    int x = -m_scrollX;
    for (size_t i = 0; i < m_columns.size(); ++i) {
        if (!m_columns[i].visible) continue;

        RECT colRect = headerRect;
        colRect.left = x;
        colRect.right = x + m_columns[i].width;

        bool isSortColumn = (m_columns[i].type == m_sortState.column);
        RenderColumnHeader(dc, m_columns[i], colRect, isSortColumn);

        x += m_columns[i].width;
    }
}

void CustomFileListView::RecalculateLayout() {
    if (!m_layoutDirty) return;

    CalculateItemPositions();
    m_layoutDirty = false;
}

void CustomFileListView::CalculateItemPositions() {
    if (m_items.empty()) return;

    RECT clientRect;
    GetClientRect(m_hwnd, &clientRect);

    SIZE itemSize = GetItemSize();
    int itemsPerRow = GetItemsPerRow();

    int x = 0;
    int y = 0;
    int col = 0;

    for (size_t i = 0; i < m_items.size(); ++i) {
        m_items[i].bounds.left = x;
        m_items[i].bounds.top = y;
        m_items[i].bounds.right = x + itemSize.cx;
        m_items[i].bounds.bottom = y + itemSize.cy;

        col++;
        if (col >= itemsPerRow) {
            col = 0;
            x = 0;
            y += itemSize.cy + m_itemSpacing;
        } else {
            x += itemSize.cx + m_itemSpacing;
        }
    }

    UpdateScrollbars();
}

RECT CustomFileListView::GetItemRect(int index) const {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return {};
    }
    return m_items[index].bounds;
}

SIZE CustomFileListView::GetItemSize() const {
    SIZE size = {};

    switch (m_viewMode) {
        case FileListViewMode::ExtraLargeIcons:
            size.cx = 256;
            size.cy = 256 + 40;
            break;
        case FileListViewMode::LargeIcons:
            size.cx = 96;
            size.cy = 96 + 40;
            break;
        case FileListViewMode::MediumIcons:
            size.cx = 48;
            size.cy = 48 + 40;
            break;
        case FileListViewMode::SmallIcons:
        case FileListViewMode::List:
            size.cx = 200;
            size.cy = 20;
            break;
        case FileListViewMode::Details:
            size.cx = 0;  // Full width
            size.cy = 20;
            RECT clientRect;
            if (GetClientRect(m_hwnd, &clientRect)) {
                size.cx = clientRect.right - clientRect.left;
            }
            break;
        case FileListViewMode::Tiles:
        case FileListViewMode::Content:
            size.cx = 250;
            size.cy = 48;
            break;
    }

    return size;
}

int CustomFileListView::GetItemsPerRow() const {
    if (m_viewMode == FileListViewMode::Details ||
        m_viewMode == FileListViewMode::List) {
        return 1;
    }

    RECT clientRect;
    GetClientRect(m_hwnd, &clientRect);
    int clientWidth = clientRect.right - clientRect.left;

    SIZE itemSize = GetItemSize();
    int itemsPerRow = std::max(1, static_cast<int>(clientWidth / (itemSize.cx + m_itemSpacing)));

    return itemsPerRow;
}

int CustomFileListView::GetVisibleItemCount() const {
    RECT clientRect;
    GetClientRect(m_hwnd, &clientRect);
    int clientHeight = clientRect.bottom - clientRect.top;

    SIZE itemSize = GetItemSize();
    int rowsVisible = clientHeight / (itemSize.cy + m_itemSpacing);
    int itemsPerRow = GetItemsPerRow();

    return rowsVisible * itemsPerRow;
}

void CustomFileListView::UpdateScrollbars() {
    RECT clientRect;
    GetClientRect(m_hwnd, &clientRect);
    int clientWidth = clientRect.right - clientRect.left;
    int clientHeight = clientRect.bottom - clientRect.top;

    // Calculate total content size
    int totalHeight = 0;
    int totalWidth = 0;

    if (!m_items.empty()) {
        const auto& lastItem = m_items.back();
        totalHeight = lastItem.bounds.bottom;
        totalWidth = lastItem.bounds.right;
    }

    // Setup vertical scrollbar
    SCROLLINFO si = {};
    si.cbSize = sizeof(SCROLLINFO);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = totalHeight;
    si.nPage = clientHeight;
    si.nPos = m_scrollY;
    SetScrollInfo(m_hwnd, SB_VERT, &si, TRUE);

    // Setup horizontal scrollbar
    si.nMax = totalWidth;
    si.nPage = clientWidth;
    si.nPos = m_scrollX;
    SetScrollInfo(m_hwnd, SB_HORZ, &si, TRUE);
}

void CustomFileListView::DoScroll(int dx, int dy) {
    m_scrollX += dx;
    m_scrollY += dy;

    // Clamp scroll positions
    m_scrollX = std::max(0, m_scrollX);
    m_scrollY = std::max(0, m_scrollY);

    UpdateScrollbars();
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

int CustomFileListView::HitTest(POINT pt) const {
    POINT scrolledPt = pt;
    scrolledPt.x += m_scrollX;
    scrolledPt.y += m_scrollY;

    for (size_t i = 0; i < m_items.size(); ++i) {
        if (PtInRect(&m_items[i].bounds, scrolledPt)) {
            return static_cast<int>(i);
        }
    }

    return -1;
}

void CustomFileListView::UpdateSelection(int index, bool ctrl, bool shift) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    if (shift && m_lastSelectedIndex >= 0) {
        // Range selection
        int start = std::min(m_lastSelectedIndex, index);
        int end = std::max(m_lastSelectedIndex, index);

        if (!ctrl) {
            DeselectAll();
        }

        for (int i = start; i <= end; ++i) {
            if (i >= 0 && i < static_cast<int>(m_items.size())) {
                m_items[i].isSelected = true;
                InvalidateRect(m_hwnd, &m_items[i].bounds, FALSE);
            }
        }
    } else if (ctrl) {
        // Toggle selection
        m_items[index].isSelected = !m_items[index].isSelected;
        InvalidateRect(m_hwnd, &m_items[index].bounds, FALSE);
    } else {
        // Single selection
        DeselectAll();
        m_items[index].isSelected = true;
        InvalidateRect(m_hwnd, &m_items[index].bounds, FALSE);
    }

    m_lastSelectedIndex = index;
    NotifySelectionChanged();
}

void CustomFileListView::DeselectAll() {
    for (auto& item : m_items) {
        if (item.isSelected) {
            item.isSelected = false;
            InvalidateRect(m_hwnd, &item.bounds, FALSE);
        }
    }
}

std::vector<int> CustomFileListView::GetSelectedIndices() const {
    std::vector<int> indices;
    for (size_t i = 0; i < m_items.size(); ++i) {
        if (m_items[i].isSelected) {
            indices.push_back(static_cast<int>(i));
        }
    }
    return indices;
}

void CustomFileListView::NotifySelectionChanged() {
    // Notify parent window of selection change
    NMHDR nmhdr = {};
    nmhdr.hwndFrom = m_hwnd;
    nmhdr.code = LVN_ITEMCHANGED;
    SendMessage(GetParent(m_hwnd), WM_NOTIFY, 0, reinterpret_cast<LPARAM>(&nmhdr));
}

void CustomFileListView::AddItem(const FileListItem& item) {
    m_items.push_back(item);
    InvalidateLayout();
}

void CustomFileListView::RemoveItem(int index) {
    if (index >= 0 && index < static_cast<int>(m_items.size())) {
        auto& item = m_items[index];
        if (item.icon) DestroyIcon(item.icon);
        if (item.thumbnail) item.thumbnail->Release();
        if (item.pidl) CoTaskMemFree(item.pidl);

        m_items.erase(m_items.begin() + index);
        InvalidateLayout();
    }
}

void CustomFileListView::UpdateItem(int index, const FileListItem& item) {
    if (index >= 0 && index < static_cast<int>(m_items.size())) {
        m_items[index] = item;
        InvalidateRect(m_hwnd, &m_items[index].bounds, FALSE);
    }
}

void CustomFileListView::ClearItems() {
    for (auto& item : m_items) {
        if (item.icon) DestroyIcon(item.icon);
        if (item.thumbnail) item.thumbnail->Release();
        if (item.pidl) CoTaskMemFree(item.pidl);
    }
    m_items.clear();
    InvalidateLayout();
}

void CustomFileListView::RefreshItems() {
    ClearItems();
    QueryShellViewItems();
}

void CustomFileListView::SetViewMode(FileListViewMode mode) {
    if (m_viewMode != mode) {
        m_viewMode = mode;
        InvalidateLayout();
    }
}

void CustomFileListView::SelectItem(int index, bool addToSelection) {
    UpdateSelection(index, addToSelection, false);
}

void CustomFileListView::ScrollTo(int index) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    RECT itemRect = GetItemRect(index);
    m_scrollY = itemRect.top;
    UpdateScrollbars();
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

void CustomFileListView::EnsureVisible(int index) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    RECT clientRect;
    GetClientRect(m_hwnd, &clientRect);
    RECT itemRect = GetItemRect(index);

    // Check if item is above visible area
    if (itemRect.top < m_scrollY) {
        m_scrollY = itemRect.top;
    }
    // Check if item is below visible area
    else if (itemRect.bottom > m_scrollY + (clientRect.bottom - clientRect.top)) {
        m_scrollY = itemRect.bottom - (clientRect.bottom - clientRect.top);
    }

    UpdateScrollbars();
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

void CustomFileListView::InvalidateLayout() {
    m_layoutDirty = true;
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

FileListItem* CustomFileListView::GetItem(int index) {
    if (index >= 0 && index < static_cast<int>(m_items.size())) {
        return &m_items[index];
    }
    return nullptr;
}

void CustomFileListView::SetBackgroundPaintCallback(
    bool (*callback)(HDC dc, HWND hwnd, const RECT& rect, void* context),
    void* context) {
    m_backgroundPaintCallback = callback;
    m_backgroundPaintContext = context;
}

void CustomFileListView::SetItemPaintCallback(
    bool (*callback)(HDC dc, const FileListItem& item, const RECT& rect, void* context),
    void* context) {
    m_itemPaintCallback = callback;
    m_itemPaintContext = context;
}

void CustomFileListView::SetGlowCoordinator(shelltabs::ExplorerGlowCoordinator* coordinator) {
    m_coordinator = coordinator;
}

void CustomFileListView::SetColorDescriptor(const shelltabs::SurfaceColorDescriptor* descriptor) {
    m_colorDescriptor = descriptor;
}

bool CustomFileListView::AttachToShellView(IShellView* shellView) {
    if (!shellView) return false;

    // Release old references
    if (m_shellView) m_shellView->Release();
    if (m_shellFolderView) m_shellFolderView->Release();
    if (m_shellFolder) m_shellFolder->Release();

    m_shellView = shellView;
    m_shellView->AddRef();

    // Query for IShellFolderView interface
    HRESULT hr = shellView->QueryInterface(IID_PPV_ARGS(&m_shellFolderView));
    if (FAILED(hr)) {
        // Try to get it through IFolderView
        Microsoft::WRL::ComPtr<IFolderView> folderView;
        if (SUCCEEDED(shellView->QueryInterface(IID_PPV_ARGS(&folderView)))) {
            folderView->QueryInterface(IID_PPV_ARGS(&m_shellFolderView));
        }
    }

    // Get the IShellFolder interface
    Microsoft::WRL::ComPtr<IPersistFolder2> persistFolder;
    hr = shellView->QueryInterface(IID_PPV_ARGS(&persistFolder));
    if (SUCCEEDED(hr) && persistFolder) {
        LPITEMIDLIST pidlFolder = nullptr;
        if (SUCCEEDED(persistFolder->GetCurFolder(&pidlFolder)) && pidlFolder) {
            SHGetDesktopFolder(&m_shellFolder);

            // If not desktop, bind to the folder
            if (pidlFolder->mkid.cb != 0) {
                IShellFolder* desktop = m_shellFolder;
                m_shellFolder = nullptr;
                desktop->BindToObject(pidlFolder, nullptr, IID_PPV_ARGS(&m_shellFolder));
                desktop->Release();
            }

            CoTaskMemFree(pidlFolder);
        }
    }

    // If still don't have folder, try IServiceProvider
    if (!m_shellFolder) {
        Microsoft::WRL::ComPtr<IServiceProvider> serviceProvider;
        if (SUCCEEDED(shellView->QueryInterface(IID_PPV_ARGS(&serviceProvider)))) {
            serviceProvider->QueryService(SID_SFolderView, IID_PPV_ARGS(&m_shellFolder));
        }
    }

    // Sync items from shell view
    if (m_shellFolder) {
        // Sync view mode from Explorer
        SyncViewModeFromShellView();

        // Register for shell change notifications
        RegisterShellChangeNotify();

        // Load items
        RefreshItems();

        return true;
    }

    return false;
}

void CustomFileListView::SyncWithShellView() {
    // Synchronize our item list with the shell view
    RefreshItems();
}

void CustomFileListView::QueryShellViewItems() {
    if (!m_shellFolder) {
        LogMessage(LogLevel::Warning, L"QueryShellViewItems: No shell folder available");
        return;
    }

    // Enumerate items from the shell folder
    Microsoft::WRL::ComPtr<IEnumIDList> enumIDList;
    HRESULT hr = m_shellFolder->EnumObjects(
        m_hwnd,
        SHCONTF_FOLDERS | SHCONTF_NONFOLDERS | SHCONTF_INCLUDEHIDDEN,
        &enumIDList
    );

    if (FAILED(hr) || !enumIDList) {
        LogMessage(LogLevel::Warning, L"QueryShellViewItems: Failed to enumerate objects (hr=0x%08X)", hr);
        return;
    }

    // Clear existing items first
    ClearItems();

    LPITEMIDLIST pidl = nullptr;
    ULONG fetched = 0;
    int itemIndex = 0;

    while (enumIDList->Next(1, &pidl, &fetched) == S_OK && pidl) {
        FileListItem item;
        item.pidl = pidl;  // Store the PIDL
        item.itemIndex = itemIndex++;

        // Get display name
        STRRET strret;
        if (SUCCEEDED(m_shellFolder->GetDisplayNameOf(pidl, SHGDN_NORMAL, &strret))) {
            wchar_t displayName[MAX_PATH] = {};
            StrRetToBufW(&strret, pidl, displayName, MAX_PATH);
            item.displayName = displayName;
        }

        // Get full path for file system items
        STRRET pathStrret;
        if (SUCCEEDED(m_shellFolder->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &pathStrret))) {
            wchar_t fullPath[MAX_PATH] = {};
            StrRetToBufW(&pathStrret, pidl, fullPath, MAX_PATH);
            item.fullPath = fullPath;
        }

        // Check if it's a folder
        SFGAOF attributes = SFGAO_FOLDER;
        if (SUCCEEDED(m_shellFolder->GetAttributesOf(1, (LPCITEMIDLIST*)&pidl, &attributes))) {
            item.isFolder = (attributes & SFGAO_FOLDER) != 0;
        }

        // Get file information (size, date, attributes) for file system items
        if (!item.fullPath.empty()) {
            WIN32_FIND_DATAW findData;
            HANDLE hFind = FindFirstFileW(item.fullPath.c_str(), &findData);
            if (hFind != INVALID_HANDLE_VALUE) {
                item.dateModified = findData.ftLastWriteTime;
                item.dateCreated = findData.ftCreationTime;
                item.attributes = findData.dwFileAttributes;

                if (!item.isFolder) {
                    ULARGE_INTEGER fileSize;
                    fileSize.LowPart = findData.nFileSizeLow;
                    fileSize.HighPart = findData.nFileSizeHigh;
                    item.fileSize = fileSize.QuadPart;
                }

                FindClose(hFind);
            }
        }

        // Get file type
        if (item.isFolder) {
            item.fileType = L"File folder";
        } else if (!item.fullPath.empty()) {
            // Get file extension
            const wchar_t* ext = wcsrchr(item.fullPath.c_str(), L'.');
            if (ext && ext[1]) {
                item.fileType = ext + 1;  // Skip the dot
                // Convert to uppercase
                std::transform(item.fileType.begin(), item.fileType.end(),
                             item.fileType.begin(), ::towupper);
                item.fileType += L" File";
            } else {
                item.fileType = L"File";
            }
        }

        // Get icon (will be loaded on demand)
        item.icon = GetShellIcon(pidl, m_viewMode == FileListViewMode::ExtraLargeIcons ||
                                       m_viewMode == FileListViewMode::LargeIcons);

        // Add the item
        m_items.push_back(item);

        // Don't free pidl here - we store it in the item
        pidl = nullptr;
    }

    LogMessage(LogLevel::Info, L"QueryShellViewItems: Loaded %zu items", m_items.size());

    // Sort items by default sort state
    SortItems();

    // Trigger layout recalculation
    InvalidateLayout();
}

HICON CustomFileListView::GetShellIcon(LPITEMIDLIST pidl, bool large) const {
    if (!pidl || !m_shellFolder) return nullptr;

    // Try to get icon using IExtractIcon for better quality
    Microsoft::WRL::ComPtr<IExtractIconW> extractIcon;
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        1,
        (LPCITEMIDLIST*)&pidl,
        IID_IExtractIconW,
        nullptr,
        (void**)&extractIcon
    );

    if (SUCCEEDED(hr) && extractIcon) {
        wchar_t iconPath[MAX_PATH] = {};
        int iconIndex = 0;
        UINT flags = 0;

        hr = extractIcon->GetIconLocation(
            large ? GIL_FORSHELL : GIL_FORSHELL,
            iconPath,
            MAX_PATH,
            &iconIndex,
            &flags
        );

        if (SUCCEEDED(hr)) {
            HICON hIconLarge = nullptr;
            HICON hIconSmall = nullptr;

            hr = extractIcon->Extract(iconPath, iconIndex, &hIconLarge, &hIconSmall,
                                     MAKELONG(large ? 32 : 16, large ? 32 : 16));

            if (SUCCEEDED(hr)) {
                return large ? hIconLarge : hIconSmall;
            }
        }
    }

    // Fallback to SHGetFileInfo
    // Try to get the full path for file system items
    STRRET strret;
    if (SUCCEEDED(m_shellFolder->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &strret))) {
        wchar_t path[MAX_PATH] = {};
        StrRetToBufW(&strret, pidl, path, MAX_PATH);

        SHFILEINFOW sfi = {};
        if (SHGetFileInfoW(path, 0, &sfi, sizeof(sfi),
                          SHGFI_ICON | (large ? SHGFI_LARGEICON : SHGFI_SMALLICON))) {
            return sfi.hIcon;
        }
    }

    // Last resort: use PIDL directly with SHGetFileInfo
    SHFILEINFOW sfi = {};
    SHGetFileInfoW(reinterpret_cast<LPCWSTR>(pidl), 0, &sfi, sizeof(sfi),
                   SHGFI_PIDL | SHGFI_ICON | (large ? SHGFI_LARGEICON : SHGFI_SMALLICON));

    return sfi.hIcon;
}

IWICBitmapSource* CustomFileListView::GetShellThumbnail(LPITEMIDLIST pidl) const {
    if (!pidl || !m_shellFolder || !m_wicFactory) return nullptr;

    // Get IShellItem from PIDL
    Microsoft::WRL::ComPtr<IShellItem> shellItem;

    // Get full path
    STRRET strret;
    if (FAILED(m_shellFolder->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &strret))) {
        return nullptr;
    }

    wchar_t path[MAX_PATH] = {};
    StrRetToBufW(&strret, pidl, path, MAX_PATH);

    HRESULT hr = SHCreateItemFromParsingName(path, nullptr, IID_PPV_ARGS(&shellItem));
    if (FAILED(hr) || !shellItem) {
        return nullptr;
    }

    // Get IShellItemImageFactory
    Microsoft::WRL::ComPtr<IShellItemImageFactory> imageFactory;
    hr = shellItem.As(&imageFactory);
    if (FAILED(hr) || !imageFactory) {
        return nullptr;
    }

    // Request thumbnail (256x256 for extra large icons)
    SIZE thumbnailSize = { 256, 256 };
    if (m_viewMode == FileListViewMode::LargeIcons) {
        thumbnailSize.cx = thumbnailSize.cy = 96;
    } else if (m_viewMode == FileListViewMode::MediumIcons) {
        thumbnailSize.cx = thumbnailSize.cy = 48;
    }

    HBITMAP hBitmap = nullptr;
    hr = imageFactory->GetImage(thumbnailSize, SIIGBF_THUMBNAILONLY, &hBitmap);

    if (FAILED(hr) || !hBitmap) {
        return nullptr;
    }

    // Convert HBITMAP to WIC bitmap
    IWICBitmap* wicBitmap = nullptr;
    hr = m_wicFactory->CreateBitmapFromHBITMAP(hBitmap, nullptr, WICBitmapUseAlpha, &wicBitmap);

    DeleteObject(hBitmap);

    return wicBitmap;
}

// ============================================================================
// DirectUIReplacementHook Implementation
//
// This hook system ensures that our custom DirectUIHWND implementation is used
// everywhere in explorer.exe by hooking multiple Windows API functions:
//
// 1. CreateWindowExW - Extended window creation (primary hook)
// 2. CreateWindowW - Standard window creation (fallback)
// 3. FindWindowW - Window search by class name
// 4. FindWindowExW - Extended window search
//
// When explorer.exe tries to create or find a DirectUIHWND window, we intercept
// the call and substitute our ShellTabsFileListView custom implementation instead.
// ============================================================================

bool DirectUIReplacementHook::s_enabled = false;
void* DirectUIReplacementHook::s_originalCreateWindowExW = nullptr;
void* DirectUIReplacementHook::s_originalCreateWindowW = nullptr;
void* DirectUIReplacementHook::s_originalFindWindowW = nullptr;
void* DirectUIReplacementHook::s_originalFindWindowExW = nullptr;
std::unordered_map<HWND, CustomFileListView*> DirectUIReplacementHook::s_instances;

bool DirectUIReplacementHook::Initialize() {
    if (s_enabled) return true;

    if (MH_Initialize() != MH_OK) {
        LogMessage(LogLevel::Error, L"DirectUIReplacementHook: Failed to initialize MinHook");
        return false;
    }

    // Hook CreateWindowExW
    if (MH_CreateHook(&CreateWindowExW, &CreateWindowExW_Hook,
                      &s_originalCreateWindowExW) != MH_OK) {
        LogMessage(LogLevel::Error, L"DirectUIReplacementHook: Failed to create CreateWindowExW hook");
        MH_Uninitialize();
        return false;
    }

    if (MH_EnableHook(&CreateWindowExW) != MH_OK) {
        LogMessage(LogLevel::Error, L"DirectUIReplacementHook: Failed to enable CreateWindowExW hook");
        MH_Uninitialize();
        return false;
    }

    // Hook CreateWindowW
    if (MH_CreateHook(&CreateWindowW, &CreateWindowW_Hook,
                      &s_originalCreateWindowW) != MH_OK) {
        LogMessage(LogLevel::Warning, L"DirectUIReplacementHook: Failed to create CreateWindowW hook");
        // Continue anyway - this is not critical
    } else {
        if (MH_EnableHook(&CreateWindowW) != MH_OK) {
            LogMessage(LogLevel::Warning, L"DirectUIReplacementHook: Failed to enable CreateWindowW hook");
        } else {
            LogMessage(LogLevel::Info, L"DirectUIReplacementHook: CreateWindowW hooked successfully");
        }
    }

    // Hook FindWindowW
    if (MH_CreateHook(&FindWindowW, &FindWindowW_Hook,
                      &s_originalFindWindowW) != MH_OK) {
        LogMessage(LogLevel::Warning, L"DirectUIReplacementHook: Failed to create FindWindowW hook");
    } else {
        if (MH_EnableHook(&FindWindowW) != MH_OK) {
            LogMessage(LogLevel::Warning, L"DirectUIReplacementHook: Failed to enable FindWindowW hook");
        } else {
            LogMessage(LogLevel::Info, L"DirectUIReplacementHook: FindWindowW hooked successfully");
        }
    }

    // Hook FindWindowExW
    if (MH_CreateHook(&FindWindowExW, &FindWindowExW_Hook,
                      &s_originalFindWindowExW) != MH_OK) {
        LogMessage(LogLevel::Warning, L"DirectUIReplacementHook: Failed to create FindWindowExW hook");
    } else {
        if (MH_EnableHook(&FindWindowExW) != MH_OK) {
            LogMessage(LogLevel::Warning, L"DirectUIReplacementHook: Failed to enable FindWindowExW hook");
        } else {
            LogMessage(LogLevel::Info, L"DirectUIReplacementHook: FindWindowExW hooked successfully");
        }
    }

    s_enabled = true;
    LogMessage(LogLevel::Info, L"DirectUIReplacementHook: All hooks initialized successfully");
    return true;
}

void DirectUIReplacementHook::Shutdown() {
    if (!s_enabled) return;

    // Disable and remove all hooks
    if (s_originalCreateWindowExW) {
        MH_DisableHook(&CreateWindowExW);
        MH_RemoveHook(&CreateWindowExW);
    }

    if (s_originalCreateWindowW) {
        MH_DisableHook(&CreateWindowW);
        MH_RemoveHook(&CreateWindowW);
    }

    if (s_originalFindWindowW) {
        MH_DisableHook(&FindWindowW);
        MH_RemoveHook(&FindWindowW);
    }

    if (s_originalFindWindowExW) {
        MH_DisableHook(&FindWindowExW);
        MH_RemoveHook(&FindWindowExW);
    }

    MH_Uninitialize();

    s_enabled = false;
    LogMessage(LogLevel::Info, L"DirectUIReplacementHook: All hooks shut down");
}

void DirectUIReplacementHook::RegisterInstance(HWND hwnd, CustomFileListView* view) {
    std::lock_guard<std::mutex> lock(g_instanceMutex);
    s_instances[hwnd] = view;
}

void DirectUIReplacementHook::UnregisterInstance(HWND hwnd) {
    std::lock_guard<std::mutex> lock(g_instanceMutex);
    s_instances.erase(hwnd);
}

CustomFileListView* DirectUIReplacementHook::GetInstance(HWND hwnd) {
    std::lock_guard<std::mutex> lock(g_instanceMutex);
    auto it = s_instances.find(hwnd);
    return (it != s_instances.end()) ? it->second : nullptr;
}

HWND WINAPI DirectUIReplacementHook::CreateWindowExW_Hook(
    DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
    DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam) {

    // Check if this is a DirectUIHWND window
    if (IsDirectUIClassName(lpClassName)) {
        LogMessage(LogLevel::Info, L"DirectUIReplacementHook: Intercepted CreateWindowExW for DirectUIHWND");
        return CreateReplacementWindow(dwExStyle, dwStyle, X, Y,
                                      nWidth, nHeight, hWndParent, hInstance);
    }

    // Call original CreateWindowExW
    auto original = reinterpret_cast<decltype(&CreateWindowExW)>(s_originalCreateWindowExW);
    return original(dwExStyle, lpClassName, lpWindowName, dwStyle,
                   X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
}

HWND WINAPI DirectUIReplacementHook::CreateWindowW_Hook(
    LPCWSTR lpClassName, LPCWSTR lpWindowName,
    DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam) {

    // Check if this is a DirectUIHWND window
    if (IsDirectUIClassName(lpClassName)) {
        LogMessage(LogLevel::Info, L"DirectUIReplacementHook: Intercepted CreateWindowW for DirectUIHWND");
        return CreateReplacementWindow(0, dwStyle, X, Y,
                                      nWidth, nHeight, hWndParent, hInstance);
    }

    // Call original CreateWindowW
    auto original = reinterpret_cast<decltype(&CreateWindowW)>(s_originalCreateWindowW);
    return original(lpClassName, lpWindowName, dwStyle,
                   X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
}

HWND WINAPI DirectUIReplacementHook::FindWindowW_Hook(
    LPCWSTR lpClassName, LPCWSTR lpWindowName) {

    // If searching for DirectUIHWND by class name, search for our replacement too
    if (lpClassName && !lpWindowName && IsDirectUIClassName(lpClassName)) {
        // First try to find our custom window
        HWND customWindow = ::FindWindowW(L"ShellTabsFileListView", nullptr);
        if (customWindow) {
            LogMessage(LogLevel::Verbose, L"DirectUIReplacementHook: FindWindowW returning custom window");
            return customWindow;
        }
    }

    // Call original FindWindowW
    auto original = reinterpret_cast<decltype(&FindWindowW)>(s_originalFindWindowW);
    return original(lpClassName, lpWindowName);
}

HWND WINAPI DirectUIReplacementHook::FindWindowExW_Hook(
    HWND hWndParent, HWND hWndChildAfter, LPCWSTR lpClassName, LPCWSTR lpWindowName) {

    // If searching for DirectUIHWND by class name, search for our replacement too
    if (lpClassName && IsDirectUIClassName(lpClassName)) {
        // First try to find our custom window
        HWND customWindow = ::FindWindowExW(hWndParent, hWndChildAfter,
                                           L"ShellTabsFileListView", lpWindowName);
        if (customWindow) {
            LogMessage(LogLevel::Verbose, L"DirectUIReplacementHook: FindWindowExW returning custom window");
            return customWindow;
        }
    }

    // Call original FindWindowExW
    auto original = reinterpret_cast<decltype(&FindWindowExW)>(s_originalFindWindowExW);
    return original(hWndParent, hWndChildAfter, lpClassName, lpWindowName);
}

bool DirectUIReplacementHook::IsDirectUIClassName(LPCWSTR className) {
    if (!className || reinterpret_cast<ULONG_PTR>(className) <= 0xFFFF) {
        return false;  // Atom or invalid
    }

    return wcscmp(className, L"DirectUIHWND") == 0;
}

HWND DirectUIReplacementHook::CreateReplacementWindow(
    DWORD dwExStyle, DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HINSTANCE hInstance) {

    // Create our custom file list view
    auto* customView = new CustomFileListView();

    HWND hwnd = customView->Create(dwExStyle, dwStyle, X, Y, nWidth, nHeight,
                                   hWndParent, hInstance);

    if (hwnd) {
        RegisterInstance(hwnd, customView);
    } else {
        delete customView;
    }

    return hwnd;
}

void CustomFileListView::SyncViewModeFromShellView() {
    if (!m_shellView) return;

    // Try to get IFolderView2 for view mode info
    Microsoft::WRL::ComPtr<IFolderView2> folderView2;
    if (SUCCEEDED(m_shellView->QueryInterface(IID_PPV_ARGS(&folderView2)))) {
        FOLDERVIEWMODE fvm;
        if (SUCCEEDED(folderView2->GetViewModeAndIconSize(&fvm, nullptr))) {
            // Map Explorer's view mode to our view mode
            switch (fvm) {
                case FVM_ICON:
                    m_viewMode = FileListViewMode::ExtraLargeIcons;
                    m_iconSize = 256;
                    break;
                case FVM_SMALLICON:
                    m_viewMode = FileListViewMode::SmallIcons;
                    m_iconSize = 16;
                    break;
                case FVM_LIST:
                    m_viewMode = FileListViewMode::List;
                    m_iconSize = 16;
                    break;
                case FVM_DETAILS:
                    m_viewMode = FileListViewMode::Details;
                    m_iconSize = 16;
                    break;
                case FVM_THUMBNAIL:
                    m_viewMode = FileListViewMode::LargeIcons;
                    m_iconSize = 96;
                    break;
                case FVM_TILE:
                    m_viewMode = FileListViewMode::Tiles;
                    m_iconSize = 48;
                    break;
                case FVM_CONTENT:
                    m_viewMode = FileListViewMode::Content;
                    m_iconSize = 48;
                    break;
                default:
                    m_viewMode = FileListViewMode::Details;
                    m_iconSize = 16;
                    break;
            }

            LogMessage(LogLevel::Info, L"Synced view mode from Explorer: mode=%d, iconSize=%d",
                      static_cast<int>(m_viewMode), m_iconSize);

            InvalidateLayout();
        }
    }
}

void CustomFileListView::RegisterShellChangeNotify() {
    if (m_shellChangeNotifyId != 0) {
        return;  // Already registered
    }

    if (m_currentFolderPath.empty()) {
        // Try to get current folder path
        if (m_shellFolder) {
            Microsoft::WRL::ComPtr<IPersistFolder2> persistFolder;
            if (SUCCEEDED(m_shellFolder->QueryInterface(IID_PPV_ARGS(&persistFolder)))) {
                LPITEMIDLIST pidl = nullptr;
                if (SUCCEEDED(persistFolder->GetCurFolder(&pidl)) && pidl) {
                    wchar_t path[MAX_PATH] = {};
                    if (SHGetPathFromIDListW(pidl, path)) {
                        m_currentFolderPath = path;
                    }
                    CoTaskMemFree(pidl);
                }
            }
        }
    }

    if (m_currentFolderPath.empty()) {
        return;  // Can't register without a path
    }

    // Register for shell change notifications
    SHChangeNotifyEntry entry = {};
    entry.pidl = nullptr;
    entry.fRecursive = FALSE;

    // Convert path to PIDL
    LPITEMIDLIST pidl = nullptr;
    if (SUCCEEDED(SHParseDisplayName(m_currentFolderPath.c_str(), nullptr, &pidl, 0, nullptr))) {
        entry.pidl = pidl;

        // Register for all file system changes
        LONG events = SHCNE_CREATE | SHCNE_DELETE | SHCNE_RENAMEITEM |
                     SHCNE_RENAMEFOLDER | SHCNE_MKDIR | SHCNE_RMDIR |
                     SHCNE_UPDATEITEM | SHCNE_UPDATEDIR;

        m_shellChangeNotifyId = SHChangeNotifyRegister(
            m_hwnd,
            SHCNRF_ShellLevel | SHCNRF_NewDelivery,
            events,
            WM_SHELL_NOTIFY,
            1,
            &entry
        );

        CoTaskMemFree(pidl);

        if (m_shellChangeNotifyId != 0) {
            LogMessage(LogLevel::Info, L"Registered shell change notifications for %s (id=%lu)",
                      m_currentFolderPath.c_str(), m_shellChangeNotifyId);
        } else {
            LogMessage(LogLevel::Warning, L"Failed to register shell change notifications");
        }
    }
}

void CustomFileListView::UnregisterShellChangeNotify() {
    if (m_shellChangeNotifyId != 0) {
        SHChangeNotifyDeregister(m_shellChangeNotifyId);
        LogMessage(LogLevel::Info, L"Unregistered shell change notifications (id=%lu)",
                  m_shellChangeNotifyId);
        m_shellChangeNotifyId = 0;
    }
}

void CustomFileListView::OnShellChange(LONG eventId, LPITEMIDLIST pidl1, LPITEMIDLIST pidl2) {
    UNREFERENCED_PARAMETER(pidl1);
    UNREFERENCED_PARAMETER(pidl2);
    LogMessage(LogLevel::Verbose, L"Shell change notification: event=0x%08X", eventId);

    // Handle various shell change events
    switch (eventId) {
        case SHCNE_CREATE:
        case SHCNE_MKDIR:
            // New item created - refresh to add it
            LogMessage(LogLevel::Info, L"Shell notification: Item created");
            RefreshItems();
            break;

        case SHCNE_DELETE:
        case SHCNE_RMDIR:
            // Item deleted - refresh to remove it
            LogMessage(LogLevel::Info, L"Shell notification: Item deleted");
            RefreshItems();
            break;

        case SHCNE_RENAMEITEM:
        case SHCNE_RENAMEFOLDER:
            // Item renamed - refresh to update name
            LogMessage(LogLevel::Info, L"Shell notification: Item renamed");
            RefreshItems();
            break;

        case SHCNE_UPDATEITEM:
        case SHCNE_UPDATEDIR:
            // Item updated - refresh to update properties
            LogMessage(LogLevel::Verbose, L"Shell notification: Item updated");
            RefreshItems();
            break;

        default:
            break;
    }
}

// ============================================================================
// Context Menu and File Operations
// ============================================================================

void CustomFileListView::ShowContextMenu(POINT pt, int itemIndex) {
    // Get selected items
    std::vector<int> selectedIndices = GetSelectedIndices();

    // If clicking on an item, ensure it's selected
    if (itemIndex >= 0) {
        bool itemSelected = false;
        for (int idx : selectedIndices) {
            if (idx == itemIndex) {
                itemSelected = true;
                break;
            }
        }

        if (!itemSelected) {
            selectedIndices.clear();
            selectedIndices.push_back(itemIndex);
        }
    }

    // Convert client to screen coordinates
    POINT screenPt = pt;
    ClientToScreen(m_hwnd, &screenPt);

    // Invoke context menu
    InvokeContextMenu(selectedIndices, screenPt);
}

bool CustomFileListView::InvokeContextMenu(const std::vector<int>& itemIndices, POINT pt) {
    if (!m_shellFolder) return false;

    // Build array of PIDLs
    std::vector<LPCITEMIDLIST> pidls;
    for (int idx : itemIndices) {
        if (idx >= 0 && idx < static_cast<int>(m_items.size())) {
            pidls.push_back(m_items[idx].pidl);
        }
    }

    if (pidls.empty()) {
        // Show background context menu
        Microsoft::WRL::ComPtr<IContextMenu> contextMenu;
        HRESULT hr = m_shellFolder->CreateViewObject(m_hwnd, IID_PPV_ARGS(&contextMenu));
        if (FAILED(hr) || !contextMenu) {
            return false;
        }

        HMENU hMenu = CreatePopupMenu();
        if (!hMenu) return false;

        hr = contextMenu->QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);
        if (FAILED(hr)) {
            DestroyMenu(hMenu);
            return false;
        }

        int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, m_hwnd, nullptr);
        if (cmd > 0) {
            CMINVOKECOMMANDINFO ici = {};
            ici.cbSize = sizeof(ici);
            ici.hwnd = m_hwnd;
            ici.lpVerb = MAKEINTRESOURCEA(cmd - 1);
            ici.nShow = SW_SHOWNORMAL;

            contextMenu->InvokeCommand(&ici);
        }

        DestroyMenu(hMenu);
        return true;
    }

    // Get context menu for items
    Microsoft::WRL::ComPtr<IContextMenu> contextMenu;
    UINT reserved = 0;
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        static_cast<UINT>(pidls.size()),
        pidls.data(),
        IID_IContextMenu,
        &reserved,
        reinterpret_cast<void**>(contextMenu.GetAddressOf())
    );

    if (FAILED(hr) || !contextMenu) {
        LogMessage(LogLevel::Warning, L"Failed to get IContextMenu (hr=0x%08X)", hr);
        return false;
    }

    // Create and populate menu
    HMENU hMenu = CreatePopupMenu();
    if (!hMenu) return false;

    hr = contextMenu->QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL | CMF_EXPLORE);
    if (FAILED(hr)) {
        DestroyMenu(hMenu);
        LogMessage(LogLevel::Warning, L"QueryContextMenu failed (hr=0x%08X)", hr);
        return false;
    }

    // Show menu and get command
    int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, m_hwnd, nullptr);

    if (cmd > 0) {
        CMINVOKECOMMANDINFO ici = {};
        ici.cbSize = sizeof(ici);
        ici.hwnd = m_hwnd;
        ici.lpVerb = MAKEINTRESOURCEA(cmd - 1);
        ici.nShow = SW_SHOWNORMAL;

        hr = contextMenu->InvokeCommand(&ici);
        if (FAILED(hr)) {
            LogMessage(LogLevel::Warning, L"InvokeCommand failed (hr=0x%08X)", hr);
        }
    }

    DestroyMenu(hMenu);
    return true;
}

void CustomFileListView::OpenItem(int index) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    const auto& item = m_items[index];

    if (!m_shellFolder || !item.pidl) {
        LogMessage(LogLevel::Warning, L"OpenItem: No shell folder or PIDL");
        return;
    }

    // Execute the default verb
    Microsoft::WRL::ComPtr<IContextMenu> contextMenu;
    LPCITEMIDLIST pidl = item.pidl;
    UINT reserved = 0;

    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        1,
        &pidl,
        IID_IContextMenu,
        &reserved,
        reinterpret_cast<void**>(contextMenu.GetAddressOf())
    );

    if (FAILED(hr) || !contextMenu) {
        LogMessage(LogLevel::Warning, L"OpenItem: Failed to get IContextMenu (hr=0x%08X)", hr);
        return;
    }

    // Invoke default command
    CMINVOKECOMMANDINFO ici = {};
    ici.cbSize = sizeof(ici);
    ici.hwnd = m_hwnd;
    ici.lpVerb = nullptr;  // Default verb
    ici.nShow = SW_SHOWNORMAL;

    hr = contextMenu->InvokeCommand(&ici);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Warning, L"OpenItem: InvokeCommand failed (hr=0x%08X)", hr);
    } else {
        LogMessage(LogLevel::Info, L"Opened item: %s", item.displayName.c_str());
    }
}

void CustomFileListView::OpenSelectedItems() {
    auto selected = GetSelectedIndices();
    for (int idx : selected) {
        OpenItem(idx);
    }
}

void CustomFileListView::DeleteSelectedItems() {
    auto selected = GetSelectedIndices();
    if (selected.empty()) return;

    // Build PIDL array
    std::vector<LPCITEMIDLIST> pidls;
    for (int idx : selected) {
        if (idx >= 0 && idx < static_cast<int>(m_items.size())) {
            pidls.push_back(m_items[idx].pidl);
        }
    }

    if (pidls.empty() || !m_shellFolder) return;

    // Get IContextMenu and invoke "delete" command
    Microsoft::WRL::ComPtr<IContextMenu> contextMenu;
    UINT reserved = 0;
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        static_cast<UINT>(pidls.size()),
        pidls.data(),
        IID_IContextMenu,
        &reserved,
        reinterpret_cast<void**>(contextMenu.GetAddressOf())
    );

    if (SUCCEEDED(hr) && contextMenu) {
        CMINVOKECOMMANDINFO ici = {};
        ici.cbSize = sizeof(ici);
        ici.hwnd = m_hwnd;
        ici.lpVerb = "delete";  // Delete verb
        ici.nShow = SW_SHOWNORMAL;

        hr = contextMenu->InvokeCommand(&ici);
        if (FAILED(hr)) {
            LogMessage(LogLevel::Warning, L"DeleteSelectedItems: InvokeCommand failed (hr=0x%08X)", hr);
        } else {
            LogMessage(LogLevel::Info, L"Deleted %zu items", selected.size());
        }
    }
}

void CustomFileListView::RenameItem(int index) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    BeginInPlaceRename(index);
}

void CustomFileListView::BeginInPlaceRename(int index) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    if (m_isRenaming) {
        EndInPlaceRename(false);
    }

    m_renameIndex = index;
    m_isRenaming = true;

    const auto& item = m_items[index];
    m_renameOriginalText = item.displayName;

    // Get item rect for edit control
    RECT itemRect = item.bounds;
    OffsetRect(&itemRect, -m_scrollX, -m_scrollY);

    // Adjust for icon
    itemRect.left += m_iconSize + 8;

    // Create edit control
    m_renameEdit = CreateWindowExW(
        0,
        L"EDIT",
        item.displayName.c_str(),
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        itemRect.left, itemRect.top,
        itemRect.right - itemRect.left, itemRect.bottom - itemRect.top,
        m_hwnd,
        nullptr,
        m_hInstance,
        nullptr
    );

    if (m_renameEdit) {
        // Set font
        SendMessageW(m_renameEdit, WM_SETFONT, reinterpret_cast<WPARAM>(m_font), TRUE);

        // Select all text
        SendMessageW(m_renameEdit, EM_SETSEL, 0, -1);

        // Focus
        SetFocus(m_renameEdit);

        LogMessage(LogLevel::Info, L"Begin rename: %s", item.displayName.c_str());
    }
}

void CustomFileListView::EndInPlaceRename(bool commit) {
    if (!m_isRenaming || !m_renameEdit) {
        return;
    }

    if (commit && m_renameIndex >= 0 && m_renameIndex < static_cast<int>(m_items.size())) {
        // Get new name from edit control
        wchar_t newName[MAX_PATH] = {};
        GetWindowTextW(m_renameEdit, newName, MAX_PATH);

        if (wcscmp(newName, m_renameOriginalText.c_str()) != 0) {
            // Name changed - perform rename via shell
            const auto& item = m_items[m_renameIndex];

            if (m_shellFolder && item.pidl) {
                // Use IShellFolder::SetNameOf to rename
                LPITEMIDLIST pidlNew = nullptr;
                HRESULT hr = m_shellFolder->SetNameOf(
                    m_hwnd,
                    item.pidl,
                    newName,
                    SHGDN_INFOLDER,
                    &pidlNew
                );

                if (SUCCEEDED(hr)) {
                    LogMessage(LogLevel::Info, L"Renamed: %s -> %s",
                              m_renameOriginalText.c_str(), newName);

                    if (pidlNew) {
                        CoTaskMemFree(pidlNew);
                    }

                    // Refresh to show new name
                    RefreshItems();
                } else {
                    LogMessage(LogLevel::Warning, L"Rename failed (hr=0x%08X)", hr);
                }
            }
        }
    }

    // Cleanup
    DestroyWindow(m_renameEdit);
    m_renameEdit = nullptr;
    m_isRenaming = false;
    m_renameIndex = -1;
    m_renameOriginalText.clear();

    // Restore focus
    SetFocus(m_hwnd);
}

void CustomFileListView::CopySelectedItems() {
    auto selected = GetSelectedIndices();
    if (selected.empty() || !m_shellFolder) return;

    // Build PIDL array
    std::vector<LPCITEMIDLIST> pidls;
    for (int idx : selected) {
        if (idx >= 0 && idx < static_cast<int>(m_items.size())) {
            pidls.push_back(m_items[idx].pidl);
        }
    }

    if (pidls.empty()) return;

    // Get IDataObject
    Microsoft::WRL::ComPtr<IDataObject> dataObject;
    UINT reserved = 0;
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        static_cast<UINT>(pidls.size()),
        pidls.data(),
        IID_IDataObject,
        &reserved,
        reinterpret_cast<void**>(dataObject.GetAddressOf())
    );

    if (SUCCEEDED(hr) && dataObject) {
        // Set to clipboard
        OleSetClipboard(dataObject.Get());
        LogMessage(LogLevel::Info, L"Copied %zu items to clipboard", selected.size());
    } else {
        LogMessage(LogLevel::Warning, L"CopySelectedItems: Failed to get IDataObject (hr=0x%08X)", hr);
    }
}

void CustomFileListView::CutSelectedItems() {
    // For cut, we use the same as copy but set a "preferred drop effect"
    CopySelectedItems();

    // TODO: Mark items with cut visual indication
    LogMessage(LogLevel::Info, L"Cut items (copy with move effect)");
}

void CustomFileListView::PasteItems() {
    // Get data from clipboard
    Microsoft::WRL::ComPtr<IDataObject> dataObject;
    HRESULT hr = OleGetClipboard(&dataObject);

    if (FAILED(hr) || !dataObject) {
        LogMessage(LogLevel::Warning, L"PasteItems: No clipboard data");
        return;
    }

    if (!m_shellFolder) return;

    // Get IDropTarget from current folder
    Microsoft::WRL::ComPtr<IDropTarget> dropTarget;
    hr = m_shellFolder->CreateViewObject(m_hwnd, IID_PPV_ARGS(&dropTarget));

    if (SUCCEEDED(hr) && dropTarget) {
        POINTL pt = { 0, 0 };
        DWORD effect = DROPEFFECT_COPY;

        hr = dropTarget->DragEnter(dataObject.Get(), MK_LBUTTON, pt, &effect);
        if (SUCCEEDED(hr)) {
            hr = dropTarget->Drop(dataObject.Get(), MK_LBUTTON, pt, &effect);
            if (SUCCEEDED(hr)) {
                LogMessage(LogLevel::Info, L"Pasted items");
            } else {
                LogMessage(LogLevel::Warning, L"PasteItems: Drop failed (hr=0x%08X)", hr);
            }
        }
    } else {
        LogMessage(LogLevel::Warning, L"PasteItems: Failed to get IDropTarget (hr=0x%08X)", hr);
    }
}

void CustomFileListView::BeginDrag(int index) {
    if (index < 0 || index >= static_cast<int>(m_items.size())) {
        return;
    }

    m_isDragging = true;
    m_dragStartIndex = index;
    SetCapture(m_hwnd);

    LogMessage(LogLevel::Info, L"Begin drag: %s", m_items[index].displayName.c_str());

    // TODO: Implement full drag-drop with IDropSource
}

bool CustomFileListView::HasClipboardData() const {
    Microsoft::WRL::ComPtr<IDataObject> dataObject;
    return SUCCEEDED(OleGetClipboard(&dataObject)) && dataObject;
}

// ============================================================================
// Phase 4: Advanced Features - Sorting, Columns, Filter, Type-Ahead
// ============================================================================

void CustomFileListView::InitializeDefaultColumns() {
    m_columns.clear();

    m_columns.push_back({ColumnType::Name, L"Name", 300, true});
    m_columns.push_back({ColumnType::DateModified, L"Date modified", 150, true});
    m_columns.push_back({ColumnType::Type, L"Type", 120, true});
    m_columns.push_back({ColumnType::Size, L"Size", 100, true});
    m_columns.push_back({ColumnType::DateCreated, L"Date created", 150, false});
    m_columns.push_back({ColumnType::Attributes, L"Attributes", 100, false});
}

void CustomFileListView::SortBy(ColumnType column, bool ascending) {
    m_sortState.column = column;
    m_sortState.ascending = ascending;
    SortItems();
}

void CustomFileListView::ToggleSort(ColumnType column) {
    if (m_sortState.column == column) {
        m_sortState.ascending = !m_sortState.ascending;
    } else {
        m_sortState.column = column;
        m_sortState.ascending = true;
    }
    SortItems();
}

void CustomFileListView::SortItems() {
    // Sort the items vector
    std::sort(m_items.begin(), m_items.end(),
             [this](const FileListItem& a, const FileListItem& b) {
                 return CompareItems(a, b, m_sortState.column, m_sortState.ascending);
             });

    // Update indices
    for (size_t i = 0; i < m_items.size(); ++i) {
        m_items[i].itemIndex = static_cast<int>(i);
    }

    InvalidateLayout();
    LogMessage(LogLevel::Info, L"Sorted by column %d (%s)",
              static_cast<int>(m_sortState.column),
              m_sortState.ascending ? L"ascending" : L"descending");
}

bool CustomFileListView::CompareItems(const FileListItem& a, const FileListItem& b,
                                      ColumnType column, bool ascending) {
    // Folders always come first
    if (a.isFolder != b.isFolder) {
        return a.isFolder;
    }

    int result = 0;

    switch (column) {
        case ColumnType::Name:
            result = _wcsicmp(a.displayName.c_str(), b.displayName.c_str());
            break;

        case ColumnType::Size:
            if (a.fileSize < b.fileSize) result = -1;
            else if (a.fileSize > b.fileSize) result = 1;
            break;

        case ColumnType::Type:
            result = _wcsicmp(a.fileType.c_str(), b.fileType.c_str());
            break;

        case ColumnType::DateModified:
            result = CompareFileTime(a.dateModified, b.dateModified);
            break;

        case ColumnType::DateCreated:
            result = CompareFileTime(a.dateCreated, b.dateCreated);
            break;

        case ColumnType::Attributes:
            if (a.attributes < b.attributes) result = -1;
            else if (a.attributes > b.attributes) result = 1;
            break;
    }

    return ascending ? (result < 0) : (result > 0);
}

int CustomFileListView::CompareFileTime(const FILETIME& a, const FILETIME& b) {
    ULARGE_INTEGER ua, ub;
    ua.LowPart = a.dwLowDateTime;
    ua.HighPart = a.dwHighDateTime;
    ub.LowPart = b.dwLowDateTime;
    ub.HighPart = b.dwHighDateTime;

    if (ua.QuadPart < ub.QuadPart) return -1;
    if (ua.QuadPart > ub.QuadPart) return 1;
    return 0;
}

std::wstring CustomFileListView::FormatFileSize(ULONGLONG size) {
    if (size == 0) return L"0 bytes";

    const wchar_t* units[] = {L"bytes", L"KB", L"MB", L"GB", L"TB"};
    int unitIndex = 0;
    double displaySize = static_cast<double>(size);

    while (displaySize >= 1024.0 && unitIndex < 4) {
        displaySize /= 1024.0;
        unitIndex++;
    }

    wchar_t buffer[64];
    if (unitIndex == 0) {
        swprintf_s(buffer, L"%llu %s", size, units[unitIndex]);
    } else {
        swprintf_s(buffer, L"%.2f %s", displaySize, units[unitIndex]);
    }

    return buffer;
}

std::wstring CustomFileListView::FormatFileTime(const FILETIME& ft) {
    SYSTEMTIME st, localSt;
    FileTimeToSystemTime(&ft, &st);
    SystemTimeToTzSpecificLocalTime(nullptr, &st, &localSt);

    wchar_t buffer[128];
    swprintf_s(buffer, L"%04d/%02d/%02d %02d:%02d",
              localSt.wYear, localSt.wMonth, localSt.wDay,
              localSt.wHour, localSt.wMinute);

    return buffer;
}

void CustomFileListView::SetColumns(const std::vector<ColumnInfo>& columns) {
    m_columns = columns;
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

int CustomFileListView::HitTestColumn(POINT pt) const {
    if (m_viewMode != FileListViewMode::Details) return -1;
    if (pt.y >= m_headerHeight) return -1;

    int x = -m_scrollX;
    for (size_t i = 0; i < m_columns.size(); ++i) {
        if (!m_columns[i].visible) continue;

        if (pt.x >= x && pt.x < x + m_columns[i].width) {
            return static_cast<int>(i);
        }
        x += m_columns[i].width;
    }

    return -1;
}

RECT CustomFileListView::GetColumnRect(int columnIndex) const {
    RECT rect = {0, 0, 0, m_headerHeight};

    if (columnIndex < 0 || columnIndex >= static_cast<int>(m_columns.size())) {
        return rect;
    }

    int x = -m_scrollX;
    for (int i = 0; i < columnIndex; ++i) {
        if (m_columns[i].visible) {
            x += m_columns[i].width;
        }
    }

    rect.left = x;
    rect.right = x + m_columns[columnIndex].width;

    return rect;
}

void CustomFileListView::RenderColumnHeader(HDC dc, const ColumnInfo& column,
                                            const RECT& rect, bool isSortColumn) {
    // Draw background
    COLORREF bgColor = RGB(240, 240, 240);
    if (isSortColumn) {
        bgColor = RGB(220, 230, 240);  // Slightly highlighted
    }

    HBRUSH brush = CreateSolidBrush(bgColor);
    FillRect(dc, &rect, brush);
    DeleteObject(brush);

    // Draw border
    HPEN pen = CreatePen(PS_SOLID, 1, RGB(180, 180, 180));
    HPEN oldPen = (HPEN)SelectObject(dc, pen);
    MoveToEx(dc, rect.right - 1, rect.top, nullptr);
    LineTo(dc, rect.right - 1, rect.bottom);
    MoveToEx(dc, rect.left, rect.bottom - 1, nullptr);
    LineTo(dc, rect.right, rect.bottom - 1);
    SelectObject(dc, oldPen);
    DeleteObject(pen);

    // Draw text
    RECT textRect = rect;
    textRect.left += 8;
    textRect.right -= 8;

    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(0, 0, 0));

    HFONT oldFont = (HFONT)SelectObject(dc, m_font);
    DrawTextW(dc, column.title.c_str(), -1, &textRect,
             DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    SelectObject(dc, oldFont);

    // Draw sort arrow if this is the sort column
    if (isSortColumn) {
        RECT arrowRect = rect;
        arrowRect.left = arrowRect.right - 20;
        RenderSortArrow(dc, arrowRect, m_sortState.ascending);
    }
}

void CustomFileListView::RenderSortArrow(HDC dc, const RECT& rect, bool ascending) {
    int centerX = (rect.left + rect.right) / 2;
    int centerY = (rect.top + rect.bottom) / 2;

    HPEN pen = CreatePen(PS_SOLID, 2, RGB(60, 60, 60));
    HPEN oldPen = (HPEN)SelectObject(dc, pen);

    if (ascending) {
        // Up arrow
        MoveToEx(dc, centerX, centerY - 4, nullptr);
        LineTo(dc, centerX - 4, centerY + 2);
        MoveToEx(dc, centerX, centerY - 4, nullptr);
        LineTo(dc, centerX + 4, centerY + 2);
    } else {
        // Down arrow
        MoveToEx(dc, centerX, centerY + 4, nullptr);
        LineTo(dc, centerX - 4, centerY - 2);
        MoveToEx(dc, centerX, centerY + 4, nullptr);
        LineTo(dc, centerX + 4, centerY - 2);
    }

    SelectObject(dc, oldPen);
    DeleteObject(pen);
}

void CustomFileListView::SetFilter(const std::wstring& filter) {
    m_filterText = filter;
    m_hasFilter = !filter.empty();
    ApplyFilter();
}

void CustomFileListView::ClearFilter() {
    m_filterText.clear();
    m_hasFilter = false;
    m_filteredItems.clear();
    InvalidateRect(m_hwnd, nullptr, FALSE);
}

void CustomFileListView::ApplyFilter() {
    if (!m_hasFilter) {
        m_filteredItems.clear();
        InvalidateRect(m_hwnd, nullptr, FALSE);
        return;
    }

    m_filteredItems.clear();
    for (const auto& item : m_items) {
        if (MatchesFilter(item)) {
            m_filteredItems.push_back(item);
        }
    }

    InvalidateLayout();
    LogMessage(LogLevel::Info, L"Filter applied: %zu of %zu items match",
              m_filteredItems.size(), m_items.size());
}

bool CustomFileListView::MatchesFilter(const FileListItem& item) const {
    if (!m_hasFilter) return true;

    // Simple case-insensitive substring match
    std::wstring lowerName = item.displayName;
    std::wstring lowerFilter = m_filterText;

    std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);
    std::transform(lowerFilter.begin(), lowerFilter.end(), lowerFilter.begin(), ::towlower);

    return lowerName.find(lowerFilter) != std::wstring::npos;
}

void CustomFileListView::BeginTypeAhead() {
    m_typeAheadText.clear();
    m_isTypeAhead = true;
}

void CustomFileListView::AddTypeAheadChar(wchar_t ch) {
    m_typeAheadText += ch;

    // Find and select first match
    SelectFirstMatch();

    // Reset timer
    if (m_typeAheadTimer) {
        KillTimer(m_hwnd, m_typeAheadTimer);
    }
    m_typeAheadTimer = SetTimer(m_hwnd, 1, 1000, nullptr);  // 1 second timeout

    LogMessage(LogLevel::Verbose, L"Type-ahead: '%s'", m_typeAheadText.c_str());
}

void CustomFileListView::EndTypeAhead() {
    m_typeAheadText.clear();
    m_isTypeAhead = false;

    if (m_typeAheadTimer) {
        KillTimer(m_hwnd, m_typeAheadTimer);
        m_typeAheadTimer = 0;
    }
}

void CustomFileListView::SelectFirstMatch() {
    if (m_typeAheadText.empty()) return;

    std::wstring lowerSearch = m_typeAheadText;
    std::transform(lowerSearch.begin(), lowerSearch.end(), lowerSearch.begin(), ::towlower);

    for (size_t i = 0; i < m_items.size(); ++i) {
        std::wstring lowerName = m_items[i].displayName;
        std::transform(lowerName.begin(), lowerName.end(), lowerName.begin(), ::towlower);

        if (lowerName.find(lowerSearch) == 0) {  // Starts with
            DeselectAll();
            SelectItem(static_cast<int>(i), false);
            EnsureVisible(static_cast<int>(i));
            break;
        }
    }
}

void CustomFileListView::RenderThumbnailDirect2D(ID2D1RenderTarget* rt,
                                                 const FileListItem& item,
                                                 const RECT& rect) {
    if (!item.thumbnail || !rt) return;

    // Convert WIC bitmap to D2D bitmap
    Microsoft::WRL::ComPtr<ID2D1Bitmap> d2dBitmap;
    HRESULT hr = rt->CreateBitmapFromWicBitmap(item.thumbnail, &d2dBitmap);

    if (SUCCEEDED(hr) && d2dBitmap) {
        D2D1_SIZE_F size = d2dBitmap->GetSize();
        D2D1_RECT_F destRect = D2D1::RectF(
            static_cast<FLOAT>(rect.left),
            static_cast<FLOAT>(rect.top),
            static_cast<FLOAT>(rect.right),
            static_cast<FLOAT>(rect.bottom)
        );

        rt->DrawBitmap(d2dBitmap.Get(), destRect, 1.0f,
                      D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
    }
}

} // namespace ShellTabs
