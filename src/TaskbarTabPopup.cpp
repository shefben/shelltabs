#include "TaskbarTabPopup.h"

#include <CommCtrl.h>
#include <windowsx.h>

#include <algorithm>

#include "Logging.h"
#include "Module.h"
#include "TabBand.h"
#include "TabBandWindow.h"

namespace shelltabs {
namespace {

constexpr wchar_t kPopupClassName[] = L"ShellTabsTaskbarPopup";
constexpr int kMaxVisibleItems = 10;
constexpr int kItemHeight = 28;

}  // namespace

TaskbarTabPopup::TaskbarTabPopup(TabBand* owner) : m_owner(owner) {}

TaskbarTabPopup::~TaskbarTabPopup() { Destroy(); }

ATOM TaskbarTabPopup::EnsurePopupWindowClass() {
    static ATOM atom = 0;
    if (atom != 0) {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.style = CS_DROPSHADOW | CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = &TaskbarTabPopup::WindowProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = GetModuleHandleInstance();
    wc.hIcon = nullptr;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszMenuName = nullptr;
    wc.lpszClassName = kPopupClassName;
    wc.hIconSm = nullptr;

    atom = RegisterClassExW(&wc);
    if (atom == 0 && GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
        atom = 1;  // sentinel for already registered
    }

    return atom;
}

void TaskbarTabPopup::EnsureWindow(HWND ownerWindow) {
    if (m_hwnd) {
        if (ownerWindow) {
            SetWindowLongPtrW(m_hwnd, GWLP_HWNDPARENT, reinterpret_cast<LONG_PTR>(ownerWindow));
        }
        return;
    }

    if (!EnsurePopupWindowClass()) {
        LogMessage(LogLevel::Warning, L"TaskbarTabPopup::EnsureWindow failed to register window class");
        return;
    }

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, kPopupClassName, L"", WS_POPUP | WS_BORDER,
                                CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, ownerWindow, nullptr,
                                GetModuleHandleInstance(), this);
    if (!hwnd) {
        LogLastError(L"CreateWindowEx(TaskbarTabPopup)", GetLastError());
        return;
    }

    m_hwnd = hwnd;
    ShowWindow(m_hwnd, SW_HIDE);
}

void TaskbarTabPopup::InitializeWindow(HWND hwnd) {
    if (m_windowInitialized) {
        return;
    }

    HINSTANCE instance = GetModuleHandleInstance();
    m_listView = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
                                 WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS |
                                     LVS_NOCOLUMNHEADER | WS_VSCROLL,
                                 0, 0, 0, 0, hwnd, nullptr, instance, nullptr);
    if (!m_listView) {
        LogLastError(L"CreateWindowEx(ListView)", GetLastError());
        return;
    }

    ListView_SetExtendedListViewStyle(m_listView,
                                      LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP);

    LVCOLUMNW column{};
    column.mask = LVCF_WIDTH;
    column.cx = m_lastColumnWidth;
    ListView_InsertColumn(m_listView, 0, &column);

    m_windowInitialized = true;
}

void TaskbarTabPopup::Populate(TabBandWindow* tabWindow) {
    m_items.clear();

    if (!m_listView) {
        return;
    }

    ListView_DeleteAllItems(m_listView);

    if (m_imageList) {
        ImageList_Destroy(m_imageList);
        m_imageList = nullptr;
    }
    ListView_SetImageList(m_listView, nullptr, LVSIL_SMALL);

    if (!tabWindow) {
        return;
    }

    const auto& data = tabWindow->GetTabData();
    m_items.reserve(data.size());
    for (const auto& item : data) {
        if (item.type != TabViewItemType::kTab) {
            continue;
        }
        m_items.push_back(item);
    }

    if (m_items.empty()) {
        return;
    }

    m_imageList = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, static_cast<int>(m_items.size()), 4);
    if (m_imageList) {
        ListView_SetImageList(m_listView, m_imageList, LVSIL_SMALL);
    }

    int selectedIndex = -1;
    for (size_t i = 0; i < m_items.size(); ++i) {
        LVITEMW entry{};
        entry.mask = LVIF_TEXT | LVIF_PARAM;
        entry.iItem = static_cast<int>(i);
        entry.pszText = const_cast<wchar_t*>(m_items[i].name.c_str());
        entry.lParam = static_cast<LPARAM>(i);

        if (m_imageList && tabWindow) {
            HICON icon = tabWindow->GetTaskbarIcon(m_items[i], true);
            if (icon) {
                int imageIndex = ImageList_AddIcon(m_imageList, icon);
                DestroyIcon(icon);
                if (imageIndex >= 0) {
                    entry.mask |= LVIF_IMAGE;
                    entry.iImage = imageIndex;
                }
            }
        }

        if (m_items[i].selected && selectedIndex < 0) {
            selectedIndex = static_cast<int>(i);
        }

        ListView_InsertItem(m_listView, &entry);
    }

    ListView_SetColumnWidth(m_listView, 0, LVSCW_AUTOSIZE_USEHEADER);
    int columnWidth = ListView_GetColumnWidth(m_listView, 0);
    if (columnWidth > 0) {
        m_lastColumnWidth = columnWidth;
    }

    if (selectedIndex >= 0) {
        ListView_SetItemState(m_listView, selectedIndex, LVIS_SELECTED | LVIS_FOCUSED,
                              LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(m_listView, selectedIndex, FALSE);
    } else {
        ListView_SetItemState(m_listView, 0, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }
}

void TaskbarTabPopup::Show(const POINT& anchor, HWND ownerWindow, TabBandWindow* tabWindow) {
    EnsureWindow(ownerWindow);
    if (!m_hwnd) {
        return;
    }

    InitializeWindow(m_hwnd);
    Populate(tabWindow);

    if (m_items.empty()) {
        Hide();
        return;
    }

    int width = std::clamp(m_lastColumnWidth + 32, 220, 480);
    const int visibleCount = static_cast<int>(std::min<size_t>(m_items.size(), kMaxVisibleItems));
    int height = std::max(kItemHeight, visibleCount * kItemHeight + 4);

    RECT clientRect{0, 0, width, height};
    const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(m_hwnd, GWL_STYLE));
    const DWORD exStyle = static_cast<DWORD>(GetWindowLongPtrW(m_hwnd, GWL_EXSTYLE));
    AdjustWindowRectEx(&clientRect, style, FALSE, exStyle);
    const int totalWidth = clientRect.right - clientRect.left;
    const int totalHeight = clientRect.bottom - clientRect.top;

    RECT workArea{};
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &workArea, 0);

    int x = anchor.x - (totalWidth / 2);
    int y = anchor.y - totalHeight - 12;
    if (x + totalWidth > workArea.right) {
        x = workArea.right - totalWidth;
    }
    if (x < workArea.left) {
        x = workArea.left;
    }
    if (y < workArea.top) {
        y = anchor.y + 12;
        if (y + totalHeight > workArea.bottom) {
            y = workArea.bottom - totalHeight;
        }
    }

    MoveWindow(m_hwnd, x, y, totalWidth, totalHeight, FALSE);
    if (m_listView) {
        MoveWindow(m_listView, 0, 0, width, height, TRUE);
    }

    ShowWindow(m_hwnd, SW_SHOWNORMAL);
    SetForegroundWindow(m_hwnd);
    if (m_listView) {
        SetFocus(m_listView);
    }
    m_visible = true;
}

void TaskbarTabPopup::Hide() {
    if (!m_hwnd || !m_visible) {
        return;
    }
    m_visible = false;
    ShowWindow(m_hwnd, SW_HIDE);
}

void TaskbarTabPopup::Destroy() {
    Hide();
    if (m_hwnd && IsWindow(m_hwnd)) {
        DestroyWindow(m_hwnd);
    }
    m_hwnd = nullptr;
    m_listView = nullptr;
    if (m_imageList) {
        ImageList_Destroy(m_imageList);
        m_imageList = nullptr;
    }
    m_items.clear();
    m_windowInitialized = false;
}

void TaskbarTabPopup::ActivateIndex(int index) {
    if (index < 0 || static_cast<size_t>(index) >= m_items.size()) {
        return;
    }

    Hide();
    if (m_owner) {
        m_owner->OnTabSelected(m_items[index].location);
    }
}

void TaskbarTabPopup::HandleNotify(NMHDR* header) {
    if (!header || header->hwndFrom != m_listView) {
        return;
    }

    switch (header->code) {
        case LVN_ITEMACTIVATE: {
            auto* activate = reinterpret_cast<NMITEMACTIVATE*>(header);
            ActivateIndex(activate ? activate->iItem : -1);
            break;
        }
        case LVN_KEYDOWN: {
            auto* key = reinterpret_cast<NMLVKEYDOWN*>(header);
            if (!key) {
                break;
            }
            if (key->wVKey == VK_RETURN) {
                int index = ListView_GetNextItem(m_listView, -1, LVNI_SELECTED);
                ActivateIndex(index);
            } else if (key->wVKey == VK_ESCAPE) {
                Hide();
            }
            break;
        }
        case NM_KILLFOCUS: {
            HideInternal();
            break;
        }
        default:
            break;
    }
}

void TaskbarTabPopup::HideInternal() { Hide(); }

LRESULT CALLBACK TaskbarTabPopup::WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* self = reinterpret_cast<TaskbarTabPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE) {
        auto* create = reinterpret_cast<LPCREATESTRUCTW>(lParam);
        self = create ? static_cast<TaskbarTabPopup*>(create->lpCreateParams) : nullptr;
        if (self) {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            self->m_hwnd = hwnd;
        }
    }

    if (!self) {
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
        case WM_CREATE:
            self->InitializeWindow(hwnd);
            return 0;
        case WM_DESTROY:
            if (self->m_imageList) {
                ImageList_Destroy(self->m_imageList);
                self->m_imageList = nullptr;
            }
            self->m_listView = nullptr;
            self->m_hwnd = nullptr;
            self->m_windowInitialized = false;
            return 0;
        case WM_ACTIVATE:
            if (LOWORD(wParam) == WA_INACTIVE) {
                self->HideInternal();
            }
            return 0;
        case WM_KILLFOCUS:
            self->HideInternal();
            return 0;
        case WM_CLOSE:
            self->Hide();
            return 0;
        case WM_NOTIFY:
            self->HandleNotify(reinterpret_cast<NMHDR*>(lParam));
            return 0;
        case WM_SIZE:
            if (self->m_listView) {
                MoveWindow(self->m_listView, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
            }
            return 0;
        default:
            break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

}  // namespace shelltabs
