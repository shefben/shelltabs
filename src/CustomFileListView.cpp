#include "CustomFileListView.h"
#include "ExplorerGlowSurfaces.h"
#include "MinHook.h"
#include "Logging.h"
#include <algorithm>
#include <mutex>
#include <Shlwapi.h>
#include <ShlObj.h>
#include <ShObjIdl.h>
#include <commoncontrols.h>
#include <wrl/client.h>

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
            case WM_DESTROY:
                return pThis->HandleDestroy();
        }
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

LRESULT CustomFileListView::HandleCreate(HWND hwnd, LPCREATESTRUCT createStruct) {
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
    // Quick search by typing
    // For now, just ignore characters during rename
    if (m_isRenaming) {
        return 0;
    }

    // Could implement type-ahead find here
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
                  GetRValue(colors.start) << 8, GetGValue(colors.start) << 8, GetBValue(colors.start) << 8, 0xFF00 },
                { clientRect.right, clientRect.bottom,
                  GetRValue(colors.end) << 8, GetGValue(colors.end) << 8, GetBValue(colors.end) << 8, 0xFF00 }
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
    // Render column headers for details view
    // Would show Name, Size, Type, Date Modified columns
    RECT headerRect = clientRect;
    headerRect.bottom = headerRect.top + 25;

    HBRUSH brush = CreateSolidBrush(RGB(240, 240, 240));
    FillRect(dc, &headerRect, brush);
    DeleteObject(brush);
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
    int itemsPerRow = std::max(1, clientWidth / (itemSize.cx + m_itemSpacing));

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

void CustomFileListView::ShowContextMenu(POINT pt, int itemIndex) {
    // Would integrate with IContextMenu to show shell context menu
    // For now, just a placeholder
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

void CustomFileListView::SetGlowCoordinator(ExplorerGlowCoordinator* coordinator) {
    m_coordinator = coordinator;
}

void CustomFileListView::SetColorDescriptor(const SurfaceColorDescriptor* descriptor) {
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

        // Get file information (size, date) for file system items
        if (!item.fullPath.empty() && !item.isFolder) {
            WIN32_FIND_DATAW findData;
            HANDLE hFind = FindFirstFileW(item.fullPath.c_str(), &findData);
            if (hFind != INVALID_HANDLE_VALUE) {
                item.dateModified = findData.ftLastWriteTime;
                ULARGE_INTEGER fileSize;
                fileSize.LowPart = findData.nFileSizeLow;
                fileSize.HighPart = findData.nFileSizeHigh;
                item.fileSize = fileSize.QuadPart;
                FindClose(hFind);
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
    // Need to get the full PIDL (desktop-relative)
    LPITEMIDLIST fullPidl = nullptr;

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
// ============================================================================

bool DirectUIReplacementHook::s_enabled = false;
void* DirectUIReplacementHook::s_originalCreateWindowExW = nullptr;
std::unordered_map<HWND, CustomFileListView*> DirectUIReplacementHook::s_instances;

bool DirectUIReplacementHook::Initialize() {
    if (s_enabled) return true;

    if (MH_Initialize() != MH_OK) {
        return false;
    }

    // Hook CreateWindowExW
    if (MH_CreateHook(&CreateWindowExW, &CreateWindowExW_Hook,
                      &s_originalCreateWindowExW) != MH_OK) {
        return false;
    }

    if (MH_EnableHook(&CreateWindowExW) != MH_OK) {
        return false;
    }

    s_enabled = true;
    return true;
}

void DirectUIReplacementHook::Shutdown() {
    if (!s_enabled) return;

    MH_DisableHook(&CreateWindowExW);
    MH_RemoveHook(&CreateWindowExW);
    MH_Uninitialize();

    s_enabled = false;
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
        return CreateReplacementWindow(dwExStyle, dwStyle, X, Y,
                                      nWidth, nHeight, hWndParent, hInstance);
    }

    // Call original CreateWindowExW
    auto original = reinterpret_cast<decltype(&CreateWindowExW)>(s_originalCreateWindowExW);
    return original(dwExStyle, lpClassName, lpWindowName, dwStyle,
                   X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
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
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        static_cast<UINT>(pidls.size()),
        pidls.data(),
        IID_IContextMenu,
        nullptr,
        reinterpret_cast<void**>(&contextMenu)
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

    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        1,
        &pidl,
        IID_IContextMenu,
        nullptr,
        reinterpret_cast<void**>(&contextMenu)
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
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        static_cast<UINT>(pidls.size()),
        pidls.data(),
        IID_IContextMenu,
        nullptr,
        reinterpret_cast<void**>(&contextMenu)
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
    HRESULT hr = m_shellFolder->GetUIObjectOf(
        m_hwnd,
        static_cast<UINT>(pidls.size()),
        pidls.data(),
        IID_IDataObject,
        nullptr,
        reinterpret_cast<void**>(&dataObject)
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

} // namespace ShellTabs
