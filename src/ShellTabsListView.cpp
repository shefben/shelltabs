#include "ShellTabsListView.h"

#include <CommCtrl.h>
#include <ShlObj_core.h>
#include <Shlwapi.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <strsafe.h>

#include <algorithm>
#include <limits>
#include <mutex>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "BreadcrumbGradient.h"
#include "ExplorerThemeUtils.h"
#include "Module.h"
#include "OptionsStore.h"

namespace shelltabs {
namespace {
constexpr wchar_t kListViewHostClassName[] = L"ShellTabs.ListViewHost";

#ifdef SVSI_NOSINGLESELECT
// Available starting with newer Windows 10 SDKs; prevents Explorer from collapsing the
// selection to a single item when we mirror state changes back to the shell view.
constexpr DWORD kSelectMultiFlag = SVSI_SELECT | SVSI_NOSINGLESELECT;
#else
// Older SDKs (for example, 10.0.19041) do not expose SVSI_NOSINGLESELECT. Fall back to
// SVSI_SELECT so the project still builds; the list view already owns the authoritative
// selection state and will keep other items highlighted locally.
constexpr DWORD kSelectMultiFlag = SVSI_SELECT;
#endif

std::mutex g_windowRegistryMutex;
std::unordered_map<HWND, ShellTabsListView*> g_windowRegistry;
std::mutex g_listViewRegistryMutex;
std::unordered_map<HWND, ShellTabsListView*> g_listViewRegistry;

#ifndef ListView_FindItemW
int ListView_FindItemW(HWND hwnd, int start, const LVFINDINFOW* findInfo) {
    return static_cast<int>(SendMessageW(
        hwnd, LVM_FINDITEMW, static_cast<WPARAM>(start),
        reinterpret_cast<LPARAM>(findInfo)));
}
#endif

#ifndef ListView_InsertColumnW
int ListView_InsertColumnW(HWND hwnd, int column, const LVCOLUMNW* info) {
    return static_cast<int>(SendMessageW(
        hwnd, LVM_INSERTCOLUMNW, static_cast<WPARAM>(column),
        reinterpret_cast<LPARAM>(info)));
}
#endif

bool CopyTextToBuffer(std::wstring_view text, wchar_t* buffer, int bufferChars) {
    if (!buffer || bufferChars <= 0) {
        return false;
    }
    if (text.empty()) {
        buffer[0] = L'\0';
        return true;
    }
    HRESULT hr = StringCchCopyNW(buffer, static_cast<size_t>(bufferChars), text.data(), static_cast<size_t>(bufferChars - 1));
    if (FAILED(hr)) {
        return false;
    }
    buffer[bufferChars - 1] = L'\0';
    return true;
}

}  // namespace

ShellTabsListView::ShellTabsListView() = default;
ShellTabsListView::~ShellTabsListView() { DestroyListView(); }

ATOM ShellTabsListView::EnsureWindowClass() {
    static ATOM atom = 0;
    if (atom != 0) {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = &ShellTabsListView::WindowProc;
    wc.hInstance = GetModuleHandleInstance();
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszClassName = kListViewHostClassName;

    atom = RegisterClassExW(&wc);
    if (atom == 0 && GetLastError() == ERROR_CLASS_ALREADY_EXISTS) {
        atom = 1;  // Sentinel indicating the class already exists.
    }
    return atom;
}

bool ShellTabsListView::Initialize(HWND parent,
                                   IFolderView2* folderView,
                                   HighlightResolver resolver,
                                   BackgroundResolver backgroundResolver,
                                   AccentColorResolver accentResolver,
                                   bool useAccentColors) {
    if (!parent || !folderView) {
        return false;
    }

    if (!EnsureWindowClass()) {
        return false;
    }

    m_folderView = folderView;
    m_highlightResolver = std::move(resolver);
    m_backgroundResolver = std::move(backgroundResolver);
    m_accentResolver = std::move(accentResolver);
    m_useAccentColors = useAccentColors;

    if (!CreateHostWindow(parent)) {
        return false;
    }

    if (!EnsureListView()) {
        return false;
    }

    m_cache.clear();

    int count = 0;
    if (SUCCEEDED(m_folderView->ItemCount(SVGIO_ALLVIEW, &count))) {
        m_itemCount = count;
        if (m_listView) {
            ListView_SetItemCountEx(m_listView, count, LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
        }
    } else {
        m_itemCount = 0;
    }

    return true;
}

bool ShellTabsListView::CreateHostWindow(HWND parent) {
    if (m_window && IsWindow(m_window)) {
        if (parent && GetParent(m_window) != parent) {
            SetParent(m_window, parent);
        }
        return true;
    }

    HWND window = CreateWindowExW(0, kListViewHostClassName, L"", WS_CHILD | WS_VISIBLE,
                                  0, 0, 0, 0, parent, nullptr, GetModuleHandleInstance(), this);
    if (!window) {
        return false;
    }

    m_window = window;
    return true;
}

bool ShellTabsListView::EnsureListView() {
    if (m_listView && IsWindow(m_listView)) {
        return true;
    }

    if (!m_window || !IsWindow(m_window)) {
        return false;
    }

    DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_OWNERDATA | LVS_SHAREIMAGELISTS |
                  LVS_SHOWSELALWAYS | LVS_SINGLESEL;
    DWORD exStyle = WS_EX_CLIENTEDGE;

    HWND listView = CreateWindowExW(exStyle, WC_LISTVIEWW, L"", style, 0, 0, 0, 0, m_window, nullptr,
                                    GetModuleHandleInstance(), nullptr);
    if (!listView) {
        return false;
    }

    ListView_SetExtendedListViewStyleEx(listView, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER,
                                        LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);

    SHFILEINFOW shfi{};
    HIMAGELIST smallIcons = reinterpret_cast<HIMAGELIST>(SHGetFileInfoW(L"", 0, &shfi, sizeof(shfi),
                                                                         SHGFI_SYSICONINDEX | SHGFI_SMALLICON));
    if (smallIcons) {
        ListView_SetImageList(listView, smallIcons, LVSIL_SMALL);
    }

    HIMAGELIST largeIcons = reinterpret_cast<HIMAGELIST>(SHGetFileInfoW(L"", 0, &shfi, sizeof(shfi),
                                                                         SHGFI_SYSICONINDEX | SHGFI_LARGEICON));
    if (largeIcons) {
        ListView_SetImageList(listView, largeIcons, LVSIL_NORMAL);
    }

    LVCOLUMNW column{};
    column.mask = LVCF_WIDTH | LVCF_TEXT;
    column.cx = 280;
    column.pszText = const_cast<LPWSTR>(L"Name");
    ListView_InsertColumnW(listView, 0, &column);

    if (!SetWindowSubclass(listView, &ShellTabsListView::ListViewSubclassProc, 0,
                           reinterpret_cast<DWORD_PTR>(this))) {
        DestroyWindow(listView);
        return false;
    }

    {
        std::scoped_lock lock(g_listViewRegistryMutex);
        g_listViewRegistry[listView] = this;
    }

    m_listView = listView;
    SubscribeListViewForHighlights(m_listView);
    return true;
}

void ShellTabsListView::DestroyListView() {
    // Cache the window handle to avoid use-after-free
    HWND cachedListView = m_listView;

    // Erase from registry BEFORE destroying to prevent other threads from finding it
    if (cachedListView) {
        std::scoped_lock lock(g_listViewRegistryMutex);
        g_listViewRegistry.erase(cachedListView);
    }

    if (cachedListView && IsWindow(cachedListView)) {
        UnsubscribeListViewForHighlights(cachedListView);
        RemoveWindowSubclass(cachedListView, &ShellTabsListView::ListViewSubclassProc, 0);
        DestroyWindow(cachedListView);
    }

    ResetBackgroundSurface();
    ResetAccentResources();
    m_listView = nullptr;
}

void ShellTabsListView::OnSize(int width, int height) {
    if (m_listView && IsWindow(m_listView)) {
        MoveWindow(m_listView, 0, 0, width, height, TRUE);
        ResetBackgroundSurface();
        InvalidateRect(m_listView, nullptr, FALSE);
    }
}

LRESULT ShellTabsListView::OnNotify(const NMHDR* header) {
    if (!header || header->hwndFrom != m_listView) {
        return 0;
    }

    switch (header->code) {
        case NM_CUSTOMDRAW: {
            auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(const_cast<NMHDR*>(header));
            LRESULT result = CDRF_DODEFAULT;
            HandleCustomDraw(draw, &result);
            return result;
        }
        case LVN_GETDISPINFOW: {
            auto* info = reinterpret_cast<NMLVDISPINFOW*>(const_cast<NMHDR*>(header));
            HandleGetDispInfo(info);
            break;
        }
        case LVN_ODCACHEHINT: {
            const auto* hint = reinterpret_cast<const NMLVCACHEHINT*>(header);
            HandleCacheHint(hint);
            break;
        }
        case LVN_ITEMCHANGING: {
            auto* change = reinterpret_cast<NMLISTVIEW*>(const_cast<NMHDR*>(header));
            HandleItemChanging(change);
            break;
        }
        case LVN_ITEMCHANGED: {
            auto* change = reinterpret_cast<NMLISTVIEW*>(const_cast<NMHDR*>(header));
            HandleItemChanged(change);
            break;
        }
        default:
            break;
    }

    return 0;
}

bool ShellTabsListView::HandleGetDispInfo(NMLVDISPINFOW* info) {
    if (!info) {
        return false;
    }

    LVITEMW& item = info->item;
    if (item.iItem < 0) {
        return false;
    }

    CachedItem* cached = EnsureCachedItem(item.iItem);
    if (!cached || !cached->pidl) {
        return false;
    }

    if ((item.mask & LVIF_PARAM) != 0) {
        item.lParam = reinterpret_cast<LPARAM>(cached->pidl.get());
    }

    if ((item.mask & LVIF_TEXT) != 0 && item.pszText) {
        if (cached->displayName.empty()) {
            cached->displayName = GetDisplayName(cached->pidl.get());
        }
        CopyTextToBuffer(cached->displayName, item.pszText, item.cchTextMax);
    }

    if ((item.mask & LVIF_IMAGE) != 0) {
        if (cached->imageIndex < 0) {
            cached->imageIndex = ResolveIconIndex(cached->pidl.get());
        }
        item.iImage = cached->imageIndex;
    }

    if ((item.mask & LVIF_STATE) != 0) {
        item.stateMask = LVIS_SELECTED | LVIS_FOCUSED | LVIS_DROPHILITED;
        if (m_listView && IsWindow(m_listView)) {
            item.state = ListView_GetItemState(m_listView, item.iItem, item.stateMask);
        } else {
            item.state = 0;
        }
    }

    return true;
}

void ShellTabsListView::HandleCacheHint(const NMLVCACHEHINT* hint) {
    if (!hint) {
        return;
    }

    if (m_itemCount <= 0) {
        m_cache.clear();
        return;
    }

    int from = std::max(0, hint->iFrom);
    int to = std::min(hint->iTo, m_itemCount - 1);
    if (from > to) {
        std::swap(from, to);
    }

    for (int index = from; index <= to; ++index) {
        EnsureCachedItem(index);
    }

    constexpr int kCacheMargin = 128;
    int keepFrom = std::max(0, from - kCacheMargin);
    int keepTo = std::min(m_itemCount - 1, to + kCacheMargin);
    PruneCache(keepFrom, keepTo);
}

void ShellTabsListView::HandleViewRangeChanged() {
    if (!m_listView || !IsWindow(m_listView) || m_itemCount <= 0) {
        return;
    }

    const int topIndex = static_cast<int>(SendMessageW(m_listView, LVM_GETTOPINDEX, 0, 0));
    if (topIndex < 0) {
        return;
    }

    int countPerPage = static_cast<int>(SendMessageW(m_listView, LVM_GETCOUNTPERPAGE, 0, 0));
    if (countPerPage <= 0) {
        countPerPage = 1;
    }

    NMLVCACHEHINT hint{};
    hint.iFrom = topIndex;
    hint.iTo = std::min(m_itemCount - 1, topIndex + countPerPage * 2);
    HandleCacheHint(&hint);
}

void ShellTabsListView::HandleItemChanging(NMLISTVIEW* change) {
    if (!change || change->iItem < 0) {
        return;
    }

    CachedItem* cached = EnsureCachedItem(change->iItem);
    if (!cached || !cached->pidl) {
        return;
    }

    change->lParam = reinterpret_cast<LPARAM>(cached->pidl.get());
}

void ShellTabsListView::HandleItemChanged(NMLISTVIEW* change) {
    if (!change || m_suppressSelectionNotifications) {
        return;
    }

    if ((change->uChanged & LVIF_STATE) == 0) {
        return;
    }

    constexpr UINT kRelevantMask = LVIS_SELECTED | LVIS_FOCUSED;
    const UINT oldState = change->uOldState & kRelevantMask;
    const UINT newState = change->uNewState & kRelevantMask;
    const UINT delta = oldState ^ newState;
    if (delta == 0 || !m_folderView) {
        return;
    }

    DWORD flags = 0;
    if ((delta & LVIS_SELECTED) != 0) {
        if ((newState & LVIS_SELECTED) != 0) {
            flags |= kSelectMultiFlag;
        } else {
            flags |= SVSI_DESELECT;
        }
    }

    if ((delta & LVIS_FOCUSED) != 0 && (newState & LVIS_FOCUSED) != 0) {
        flags |= SVSI_FOCUSED;
    }

    if (flags != 0) {
        m_folderView->SelectItem(change->iItem, flags);
    }
}

ShellTabsListView::CachedItem* ShellTabsListView::EnsureCachedItem(int index) {
    if (index < 0) {
        return nullptr;
    }

    auto it = m_cache.find(index);
    if (it != m_cache.end()) {
        return &it->second;
    }

    if (!m_folderView) {
        return nullptr;
    }

    Microsoft::WRL::ComPtr<IUnknown> viewItem;
    HRESULT hr = m_folderView->GetItem(index, IID_PPV_ARGS(&viewItem));
    if (FAILED(hr) || !viewItem) {
        return nullptr;
    }

    PIDLIST_ABSOLUTE pidl = nullptr;
    hr = SHGetIDListFromObject(viewItem.Get(), &pidl);
    if (FAILED(hr) || !pidl) {
        return nullptr;
    }

    CachedItem cached{};
    cached.pidl = UniquePidl(pidl);
    cached.displayName = GetDisplayName(cached.pidl.get());
    cached.imageIndex = -1;

    auto [insertedIt, inserted] = m_cache.emplace(index, std::move(cached));
    if (!inserted) {
        return nullptr;
    }
    return &insertedIt->second;
}

void ShellTabsListView::PruneCache(int keepFrom, int keepTo) {
    if (keepFrom > keepTo) {
        m_cache.clear();
        return;
    }

    for (auto it = m_cache.begin(); it != m_cache.end();) {
        if (it->first < keepFrom || it->first > keepTo) {
            it = m_cache.erase(it);
        } else {
            ++it;
        }
    }
}

int ShellTabsListView::ResolveIconIndex(PCIDLIST_ABSOLUTE pidl) const {
    if (!pidl) {
        return -1;
    }

    SHFILEINFOW info{};
    if (SHGetFileInfoW(reinterpret_cast<PCWSTR>(pidl), 0, &info, sizeof(info),
                        SHGFI_PIDL | SHGFI_SYSICONINDEX | SHGFI_SMALLICON)) {
        return info.iIcon;
    }
    return -1;
}

void ShellTabsListView::HandleInvalidationTargets(const PaneHighlightInvalidationTargets& targets) const {
    if (!m_listView || !IsWindow(m_listView)) {
        return;
    }

    if (targets.invalidateAll) {
        const int count = ListView_GetItemCount(m_listView);
        if (count > 0) {
            ListView_RedrawItems(m_listView, 0, count - 1);
        }
        InvalidateRect(m_listView, nullptr, FALSE);
        return;
    }

    if (targets.items.empty()) {
        return;
    }

    int minIndex = std::numeric_limits<int>::max();
    int maxIndex = std::numeric_limits<int>::min();
    RECT invalidRect{};
    bool hasRect = false;

    for (const auto& target : targets.items) {
        if (!target.pidl) {
            continue;
        }

        LVFINDINFOW find{};
        find.flags = LVFI_PARAM;
        find.lParam = reinterpret_cast<LPARAM>(target.pidl);

        int index = ListView_FindItemW(m_listView, -1, &find);
        if (index < 0) {
            continue;
        }

        minIndex = std::min(minIndex, index);
        maxIndex = std::max(maxIndex, index);

        RECT itemRect{};
        if (ListView_GetItemRect(m_listView, index, &itemRect, LVIR_BOUNDS)) {
            if (hasRect) {
                UnionRect(&invalidRect, &invalidRect, &itemRect);
            } else {
                invalidRect = itemRect;
                hasRect = true;
            }
        }
    }

    if (minIndex == std::numeric_limits<int>::max() || maxIndex == std::numeric_limits<int>::min()) {
        return;
    }

    ListView_RedrawItems(m_listView, minIndex, maxIndex);
    if (hasRect) {
        InvalidateRect(m_listView, &invalidRect, FALSE);
    } else {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
}

bool ShellTabsListView::TryResolveHighlight(int index, PaneHighlight* highlight) {
    if (!highlight || !m_highlightResolver) {
        return false;
    }

    CachedItem* cached = EnsureCachedItem(index);
    if (!cached || !cached->pidl) {
        return false;
    }

    return m_highlightResolver(cached->pidl.get(), highlight);
}

void ShellTabsListView::SetBackgroundResolver(BackgroundResolver resolver) {
    m_backgroundResolver = std::move(resolver);
    ResetBackgroundSurface();
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
}

void ShellTabsListView::SetAccentColorResolver(AccentColorResolver resolver) {
    m_accentResolver = std::move(resolver);
    ResetAccentResources();
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
}

void ShellTabsListView::SetUseAccentColors(bool enabled) {
    if (m_useAccentColors == enabled) {
        return;
    }
    m_useAccentColors = enabled;
    ResetAccentResources();
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, FALSE);
    }
}

void ShellTabsListView::SetCustomDrawObserver(CustomDrawObserver observer) {
    m_customDrawObserver = std::move(observer);
}

bool ShellTabsListView::HitTest(const POINT& clientPoint, HitTestResult* result) {
    if (!m_listView || !IsWindow(m_listView)) {
        return false;
    }

    LVHITTESTINFO hit{};
    hit.pt = clientPoint;
    const int index = ListView_SubItemHitTest(m_listView, &hit);
    if (index < 0) {
        return false;
    }

    if (result) {
        result->index = index;
        result->flags = hit.flags;
        result->pidl.reset();
        if (CachedItem* cached = EnsureCachedItem(index); cached && cached->pidl) {
            result->pidl = ClonePidl(cached->pidl.get());
        }
    }

    return true;
}

bool ShellTabsListView::SelectExclusive(int index) {
    if (!m_listView || !IsWindow(m_listView) || index < 0 || index >= m_itemCount) {
        return false;
    }

    m_suppressSelectionNotifications = true;
    ListView_SetItemState(m_listView, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetItemState(m_listView, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    m_suppressSelectionNotifications = false;

    ListView_EnsureVisible(m_listView, index, FALSE);

    if (m_folderView) {
        DWORD flags = SVSI_SELECT | SVSI_FOCUSED | SVSI_SELECTIONMARK | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE;
        m_folderView->SelectItem(index, flags);
    }

    return true;
}

bool ShellTabsListView::ToggleSelection(int index) {
    if (!m_listView || !IsWindow(m_listView) || index < 0 || index >= m_itemCount) {
        return false;
    }

    const UINT current = ListView_GetItemState(m_listView, index, LVIS_SELECTED);
    const bool selected = (current & LVIS_SELECTED) != 0;

    m_suppressSelectionNotifications = true;
    if (selected) {
        ListView_SetItemState(m_listView, index, 0, LVIS_SELECTED);
    } else {
        ListView_SetItemState(m_listView, index, LVIS_SELECTED, LVIS_SELECTED);
    }
    m_suppressSelectionNotifications = false;

    if (m_folderView) {
        DWORD flags = selected ? SVSI_DESELECT : kSelectMultiFlag;
        m_folderView->SelectItem(index, flags);
    }

    return true;
}

bool ShellTabsListView::FocusItem(int index, bool ensureVisible) {
    if (!m_listView || !IsWindow(m_listView) || index < 0 || index >= m_itemCount) {
        return false;
    }

    m_suppressSelectionNotifications = true;
    ListView_SetItemState(m_listView, -1, 0, LVIS_FOCUSED);
    ListView_SetItemState(m_listView, index, LVIS_FOCUSED, LVIS_FOCUSED);
    m_suppressSelectionNotifications = false;

    if (ensureVisible) {
        ListView_EnsureVisible(m_listView, index, FALSE);
    }

    if (m_folderView) {
        DWORD flags = SVSI_FOCUSED;
        if (ensureVisible) {
            flags |= SVSI_ENSUREVISIBLE;
        }
        m_folderView->SelectItem(index, flags);
    }

    return true;
}

bool ShellTabsListView::EnsureVisible(int index) const {
    if (!m_listView || !IsWindow(m_listView) || index < 0 || index >= m_itemCount) {
        return false;
    }
    return ListView_EnsureVisible(m_listView, index, FALSE) != FALSE;
}

int ShellTabsListView::GetNextSelectedIndex(int start) const {
    if (!m_listView || !IsWindow(m_listView)) {
        return -1;
    }
    return ListView_GetNextItem(m_listView, start, LVNI_SELECTED);
}

UINT ShellTabsListView::GetItemState(int index, UINT mask) const {
    if (!m_listView || !IsWindow(m_listView)) {
        return 0;
    }
    return ListView_GetItemState(m_listView, index, mask);
}

std::vector<ShellTabsListView::SelectionItem> ShellTabsListView::GetSelectionSnapshot() const {
    std::vector<SelectionItem> result;
    if (!m_listView || !IsWindow(m_listView)) {
        return result;
    }

    const int focusedIndex = ListView_GetNextItem(m_listView, -1, LVNI_FOCUSED);

    int index = -1;
    while ((index = ListView_GetNextItem(m_listView, index, LVNI_SELECTED)) != -1) {
        SelectionItem item{};
        item.index = index;
        item.focused = (index == focusedIndex);
        if (auto* cached = const_cast<ShellTabsListView*>(this)->EnsureCachedItem(index); cached && cached->pidl) {
            item.pidl = ClonePidl(cached->pidl.get());
        }
        result.emplace_back(std::move(item));
    }

    if (focusedIndex >= 0) {
        const bool hasFocused = std::any_of(result.begin(), result.end(), [&](const SelectionItem& entry) {
            return entry.index == focusedIndex;
        });
        if (!hasFocused) {
            SelectionItem focused{};
            focused.index = focusedIndex;
            focused.focused = true;
            if (auto* cached = const_cast<ShellTabsListView*>(this)->EnsureCachedItem(focusedIndex); cached && cached->pidl) {
                focused.pidl = ClonePidl(cached->pidl.get());
            }
            result.emplace_back(std::move(focused));
        }
    }

    return result;
}

bool ShellTabsListView::TryGetFocusedItem(SelectionItem* item) const {
    if (!item) {
        return false;
    }

    item->index = -1;
    item->pidl.reset();
    item->focused = false;

    if (!m_listView || !IsWindow(m_listView)) {
        return false;
    }

    const int focusedIndex = ListView_GetNextItem(m_listView, -1, LVNI_FOCUSED);
    if (focusedIndex < 0) {
        return false;
    }

    SelectionItem focused{};
    focused.index = focusedIndex;
    focused.focused = true;
    if (auto* cached = const_cast<ShellTabsListView*>(this)->EnsureCachedItem(focusedIndex); cached && cached->pidl) {
        focused.pidl = ClonePidl(cached->pidl.get());
    }

    *item = std::move(focused);
    return true;
}

bool ShellTabsListView::HandleCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result) {
    if (!draw || !result) {
        return false;
    }

    const DWORD stage = draw->nmcd.dwDrawStage;
    if (m_customDrawObserver) {
        m_customDrawObserver(stage);
    }
    switch (stage) {
        case CDDS_PREPAINT: {
            *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYSUBITEMDRAW | CDRF_NOTIFYPOSTPAINT;
            return true;
        }
        case CDDS_ITEMPREPAINT:
        case CDDS_ITEMPREPAINT | CDDS_SUBITEM: {
            const bool isSubItemStage = (stage & CDDS_SUBITEM) != 0;
            if (isSubItemStage && draw->iSubItem != 0) {
                *result = CDRF_DODEFAULT;
                return true;
            }

            const int index = static_cast<int>(draw->nmcd.dwItemSpec);
            bool handled = false;

            if ((draw->nmcd.uItemState & CDIS_SELECTED) != 0) {
                AccentResources* resources = nullptr;
                if (EnsureAccentResources(&resources) && resources) {
                    FillRect(draw->nmcd.hdc, &draw->nmcd.rc, resources->backgroundBrush);
                    draw->clrText = resources->textColor;
                    draw->clrTextBk = resources->accentColor;
                    *result = CDRF_NEWFONT;
                    handled = true;
                }
            }

            if (!handled && index >= 0) {
                PaneHighlight highlight{};
                if (TryResolveHighlight(index, &highlight)) {
                    bool applied = false;
                    if (highlight.hasTextColor) {
                        draw->clrText = highlight.textColor;
                        applied = true;
                    }
                    if (highlight.hasBackgroundColor) {
                        draw->clrTextBk = highlight.backgroundColor;
                        applied = true;
                    }
                    if (applied) {
                        *result = CDRF_NEWFONT;
                        handled = true;
                    }
                }
            }

            if (!handled) {
                *result = CDRF_DODEFAULT;
            } else if (!isSubItemStage) {
                *result |= CDRF_NOTIFYSUBITEMDRAW;
            }

            // Request postpaint notification if gradient font is enabled
            const ShellTabsOptions& options = OptionsStore::Instance().Get();
            if (options.enableFileGradientFont && (draw->nmcd.uItemState & CDIS_SELECTED) == 0) {
                *result |= CDRF_NOTIFYPOSTPAINT;
            }

            return true;
        }
        case CDDS_ITEMPOSTPAINT: {
            const int index = static_cast<int>(draw->nmcd.dwItemSpec);
            const ShellTabsOptions& options = OptionsStore::Instance().Get();

            // Render gradient text if enabled and item is not selected
            if (options.enableFileGradientFont && (draw->nmcd.uItemState & CDIS_SELECTED) == 0) {
                wchar_t textBuffer[MAX_PATH];
                LVITEMW item{};
                item.mask = LVIF_TEXT;
                item.iItem = index;
                item.iSubItem = 0;
                item.pszText = textBuffer;
                item.cchTextMax = MAX_PATH;

                if (ListView_GetItem(m_listView, &item) && item.pszText && item.pszText[0] != L'\0') {
                    const std::wstring text(item.pszText);

                    // Resolve gradient palette
                    BreadcrumbGradientConfig gradientConfig{};
                    gradientConfig.enabled = true;
                    gradientConfig.brightness = options.breadcrumbFontBrightness;
                    gradientConfig.useCustomFontColors = options.useCustomBreadcrumbFontColors;
                    gradientConfig.useCustomGradientColors = options.useCustomBreadcrumbGradientColors;
                    gradientConfig.fontGradientStartColor = options.breadcrumbFontGradientStartColor;
                    gradientConfig.fontGradientEndColor = options.breadcrumbFontGradientEndColor;
                    gradientConfig.gradientStartColor = options.breadcrumbGradientStartColor;
                    gradientConfig.gradientEndColor = options.breadcrumbGradientEndColor;

                    const BreadcrumbGradientPalette palette = ResolveBreadcrumbGradientPalette(gradientConfig);

                    // Get the font
                    HFONT font = reinterpret_cast<HFONT>(SendMessageW(m_listView, WM_GETFONT, 0, 0));
                    if (!font) {
                        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                    }
                    HFONT oldFont = nullptr;
                    if (font) {
                        oldFont = static_cast<HFONT>(SelectObject(draw->nmcd.hdc, font));
                    }

                    const int oldBkMode = SetBkMode(draw->nmcd.hdc, TRANSPARENT);

                    // Calculate text rectangle with proper padding
                    RECT textRect = draw->nmcd.rc;
                    textRect.left += 2;  // Small padding from left edge

                    // Measure the total text width to calculate gradient bounds
                    SIZE totalSize{};
                    GetTextExtentPoint32W(draw->nmcd.hdc, text.c_str(), static_cast<int>(text.size()), &totalSize);

                    const double gradientWidth = std::max(1.0, static_cast<double>(totalSize.cx));
                    double currentX = static_cast<double>(textRect.left);

                    // Draw each character with gradient color
                    for (size_t i = 0; i < text.size(); ++i) {
                        SIZE charSize{};
                        if (!GetTextExtentPoint32W(draw->nmcd.hdc, &text[i], 1, &charSize)) {
                            continue;
                        }

                        const double charCenterX = currentX + static_cast<double>(charSize.cx) * 0.5;
                        const double position = std::clamp((charCenterX - static_cast<double>(textRect.left)) / gradientWidth, 0.0, 1.0);
                        const COLORREF color = EvaluateBreadcrumbGradientColor(palette, position);

                        SetTextColor(draw->nmcd.hdc, color);

                        RECT charRect = textRect;
                        charRect.left = static_cast<LONG>(currentX);
                        charRect.right = charRect.left + charSize.cx;

                        DrawTextW(draw->nmcd.hdc, &text[i], 1, &charRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

                        currentX += static_cast<double>(charSize.cx);
                    }

                    SetBkMode(draw->nmcd.hdc, oldBkMode);

                    if (oldFont) {
                        SelectObject(draw->nmcd.hdc, oldFont);
                    }
                }
            }

            // Draw focus rectangle for selected items
            if ((draw->nmcd.uItemState & CDIS_SELECTED) != 0 && (draw->nmcd.uItemState & CDIS_FOCUS) != 0) {
                AccentResources* resources = nullptr;
                if (EnsureAccentResources(&resources) && resources && resources->focusPen) {
                    RECT focusRect = draw->nmcd.rc;
                    if (focusRect.right > focusRect.left && focusRect.bottom > focusRect.top) {
                        InflateRect(&focusRect, -1, -1);
                        HPEN oldPen = static_cast<HPEN>(SelectObject(draw->nmcd.hdc, resources->focusPen));
                        HBRUSH oldBrush = static_cast<HBRUSH>(
                            SelectObject(draw->nmcd.hdc, GetStockObject(HOLLOW_BRUSH)));
                        Rectangle(draw->nmcd.hdc, focusRect.left, focusRect.top, focusRect.right, focusRect.bottom);
                        SelectObject(draw->nmcd.hdc, oldBrush);
                        SelectObject(draw->nmcd.hdc, oldPen);
                    }
                }
            }
            *result = CDRF_DODEFAULT;
            return true;
        }
        default:
            break;
    }

    return false;
}

bool ShellTabsListView::PaintBackground(HDC dc) {
    if (!dc || !m_listView || !IsWindow(m_listView) || !m_backgroundResolver) {
        return false;
    }

    RECT client{};
    if (!GetClientRect(m_listView, &client) || client.right <= client.left || client.bottom <= client.top) {
        return false;
    }

    const BackgroundSource source = m_backgroundResolver();
    if (!source.bitmap) {
        ResetBackgroundSurface();
        return false;
    }

    if (!EnsureBackgroundSurface(client, source)) {
        return false;
    }

    Gdiplus::Graphics graphics(dc);
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        return false;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);

    const LONG width = client.right - client.left;
    const LONG height = client.bottom - client.top;
    const int widthInt = static_cast<int>(width);
    const int heightInt = static_cast<int>(height);
    const Gdiplus::Rect dest(client.left, client.top, widthInt, heightInt);
    const Gdiplus::Status status = graphics.DrawImage(m_backgroundSurface.bitmap.get(), dest, 0, 0, widthInt,
                                                     heightInt, Gdiplus::UnitPixel);
    return status == Gdiplus::Ok;
}

bool ShellTabsListView::EnsureBackgroundSurface(const RECT& client, const BackgroundSource& source) {
    if (!source.bitmap) {
        return false;
    }

    const LONG width = client.right - client.left;
    const LONG height = client.bottom - client.top;
    if (width <= 0 || height <= 0) {
        return false;
    }

    const int widthInt = static_cast<int>(width);
    const int heightInt = static_cast<int>(height);

    if (m_backgroundSurface.bitmap && m_backgroundSurface.size.cx == widthInt &&
        m_backgroundSurface.size.cy == heightInt && m_backgroundSurface.cacheKey == source.cacheKey) {
        return true;
    }

    const UINT srcWidth = source.bitmap->GetWidth();
    const UINT srcHeight = source.bitmap->GetHeight();
    if (srcWidth == 0 || srcHeight == 0) {
        ResetBackgroundSurface();
        return false;
    }

    if (srcWidth > static_cast<UINT>(std::numeric_limits<INT>::max()) ||
        srcHeight > static_cast<UINT>(std::numeric_limits<INT>::max())) {
        ResetBackgroundSurface();
        return false;
    }

    auto surface = std::make_unique<Gdiplus::Bitmap>(widthInt, heightInt, PixelFormat32bppARGB);
    if (!surface || surface->GetLastStatus() != Gdiplus::Ok) {
        ResetBackgroundSurface();
        return false;
    }

    Gdiplus::Graphics graphics(surface.get());
    if (graphics.GetLastStatus() != Gdiplus::Ok) {
        ResetBackgroundSurface();
        return false;
    }

    graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    graphics.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);

    const Gdiplus::Status status = graphics.DrawImage(source.bitmap,
                                                      Gdiplus::Rect(0, 0, widthInt, heightInt),
                                                      0,
                                                      0,
                                                      static_cast<INT>(srcWidth),
                                                      static_cast<INT>(srcHeight),
                                                      Gdiplus::UnitPixel);
    if (status != Gdiplus::Ok) {
        ResetBackgroundSurface();
        return false;
    }

    m_backgroundSurface.bitmap = std::move(surface);
    m_backgroundSurface.size = {widthInt, heightInt};
    m_backgroundSurface.cacheKey = source.cacheKey;
    return true;
}

void ShellTabsListView::ResetBackgroundSurface() {
    m_backgroundSurface.bitmap.reset();
    m_backgroundSurface.size = {0, 0};
    m_backgroundSurface.cacheKey.clear();
}

void ShellTabsListView::ResetAccentResources() {
    if (m_accentResources.backgroundBrush) {
        DeleteObject(m_accentResources.backgroundBrush);
        m_accentResources.backgroundBrush = nullptr;
    }
    if (m_accentResources.focusPen) {
        DeleteObject(m_accentResources.focusPen);
        m_accentResources.focusPen = nullptr;
    }
    m_accentResources.accentColor = 0;
    m_accentResources.textColor = 0;
}

bool ShellTabsListView::EnsureAccentResources(AccentResources** resources) {
    if (!ShouldUseAccentColors()) {
        ResetAccentResources();
        return false;
    }

    COLORREF accent = 0;
    COLORREF text = 0;
    if (!m_accentResolver || !m_accentResolver(&accent, &text)) {
        ResetAccentResources();
        return false;
    }

    const bool needsBrush = !m_accentResources.backgroundBrush || m_accentResources.accentColor != accent;
    const bool needsPen = !m_accentResources.focusPen || m_accentResources.textColor != text;

    if (needsBrush) {
        if (m_accentResources.backgroundBrush) {
            DeleteObject(m_accentResources.backgroundBrush);
            m_accentResources.backgroundBrush = nullptr;
        }
        m_accentResources.backgroundBrush = CreateSolidBrush(accent);
        if (!m_accentResources.backgroundBrush) {
            ResetAccentResources();
            return false;
        }
    }

    if (needsPen) {
        if (m_accentResources.focusPen) {
            DeleteObject(m_accentResources.focusPen);
            m_accentResources.focusPen = nullptr;
        }
        m_accentResources.focusPen = CreatePen(PS_SOLID, 1, text);
        if (!m_accentResources.focusPen) {
            m_accentResources.focusPen = nullptr;
        }
    }

    m_accentResources.accentColor = accent;
    m_accentResources.textColor = text;

    if (resources) {
        *resources = &m_accentResources;
    }
    return m_accentResources.backgroundBrush != nullptr;
}

bool ShellTabsListView::ShouldUseAccentColors() const {
    if (!m_useAccentColors || !m_accentResolver) {
        return false;
    }
    if (IsSystemHighContrastActive()) {
        return false;
    }
    return true;
}

void ShellTabsListView::HandleListViewThemeChanged() {
    ResetBackgroundSurface();
    ResetAccentResources();
    if (m_listView && IsWindow(m_listView)) {
        InvalidateRect(m_listView, nullptr, TRUE);
    }
}

ShellTabsListView* ShellTabsListView::FromWindow(HWND hwnd) {
    std::scoped_lock lock(g_windowRegistryMutex);
    auto it = g_windowRegistry.find(hwnd);
    if (it != g_windowRegistry.end()) {
        return it->second;
    }
    return nullptr;
}

bool ShellTabsListView::IsShellTabsListView(HWND hwnd) {
    return FromListView(hwnd) != nullptr;
}

ShellTabsListView* ShellTabsListView::FromListView(HWND hwnd) {
    std::scoped_lock lock(g_listViewRegistryMutex);
    auto it = g_listViewRegistry.find(hwnd);
    if (it != g_listViewRegistry.end()) {
        return it->second;
    }
    return nullptr;
}

LRESULT CALLBACK ShellTabsListView::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    ShellTabsListView* instance = nullptr;

    if (message == WM_NCCREATE) {
        const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        instance = static_cast<ShellTabsListView*>(create->lpCreateParams);
        if (!instance) {
            return FALSE;
        }
        instance->m_window = hwnd;
        {
            std::scoped_lock lock(g_windowRegistryMutex);
            g_windowRegistry[hwnd] = instance;
        }
        return TRUE;
    }

    instance = FromWindow(hwnd);
    if (!instance) {
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    switch (message) {
        case WM_SIZE: {
            instance->OnSize(LOWORD(lParam), HIWORD(lParam));
            return 0;
        }
        case WM_NOTIFY: {
            return instance->OnNotify(reinterpret_cast<const NMHDR*>(lParam));
        }
        case WM_DESTROY: {
            instance->DestroyListView();
            {
                std::scoped_lock lock(g_windowRegistryMutex);
                g_windowRegistry.erase(hwnd);
            }
            instance->m_window = nullptr;
            break;
        }
        default:
            break;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT CALLBACK ShellTabsListView::ListViewSubclassProc(HWND hwnd,
                                                         UINT message,
                                                         WPARAM wParam,
                                                         LPARAM lParam,
                                                         UINT_PTR subclassId,
                                                         DWORD_PTR referenceData) {
    auto* instance = reinterpret_cast<ShellTabsListView*>(referenceData);
    if (!instance) {
        instance = FromListView(hwnd);
    }
    switch (message) {
        case WM_ERASEBKGND: {
            if (instance && instance->PaintBackground(reinterpret_cast<HDC>(wParam))) {
                return 1;
            }
            break;
        }
        case WM_PAINT: {
            if (instance) {
                if (wParam) {
                    instance->PaintBackground(reinterpret_cast<HDC>(wParam));
                } else {
                    HDC dc = GetDCEx(hwnd, nullptr, DCX_CACHE | DCX_CLIPCHILDREN | DCX_CLIPSIBLINGS | DCX_WINDOW);
                    if (dc) {
                        instance->PaintBackground(dc);
                        ReleaseDC(hwnd, dc);
                    }
                }
            }
            break;
        }
        case WM_PRINTCLIENT: {
            if (instance) {
                instance->PaintBackground(reinterpret_cast<HDC>(wParam));
            }
            break;
        }
        case WM_THEMECHANGED:
        case WM_SETTINGCHANGE:
        case WM_DWMCOLORIZATIONCOLORCHANGED: {
            if (instance) {
                instance->HandleListViewThemeChanged();
            }
            break;
        }
        case WM_SIZE: {
            if (instance) {
                instance->ResetBackgroundSurface();
            }
            break;
        }
        case WM_KEYDOWN:
        case WM_MOUSEWHEEL:
        case WM_VSCROLL: {
            LRESULT result = DefSubclassProc(hwnd, message, wParam, lParam);
            if (instance) {
                instance->HandleViewRangeChanged();
            }
            return result;
        }
        case WM_NCDESTROY: {
            if (instance) {
                UnsubscribeListViewForHighlights(hwnd);
            }
            {
                std::scoped_lock lock(g_listViewRegistryMutex);
                g_listViewRegistry.erase(hwnd);
            }
            RemoveWindowSubclass(hwnd, &ShellTabsListView::ListViewSubclassProc, subclassId);
            break;
        }
        default:
            break;
    }

    return DefSubclassProc(hwnd, message, wParam, lParam);
}

}  // namespace shelltabs

