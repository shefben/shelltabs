#include "ShellTabsListView.h"

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

#include "ExplorerThemeUtils.h"
#include "Module.h"

namespace shelltabs {
namespace {
constexpr wchar_t kListViewHostClassName[] = L"ShellTabs.ListViewHost";

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

    if (!CreateWindow(parent)) {
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

bool ShellTabsListView::CreateWindow(HWND parent) {
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
    if (m_listView && IsWindow(m_listView)) {
        UnsubscribeListViewForHighlights(m_listView);
        RemoveWindowSubclass(m_listView, &ShellTabsListView::ListViewSubclassProc, 0);
        DestroyWindow(m_listView);
    }
    {
        std::scoped_lock lock(g_listViewRegistryMutex);
        g_listViewRegistry.erase(m_listView);
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
            auto* change = reinterpret_cast<NMLVITEMCHANGE*>(const_cast<NMHDR*>(header));
            HandleItemChanging(change);
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

void ShellTabsListView::HandleItemChanging(NMLVITEMCHANGE* change) {
    if (!change || change->iItem < 0) {
        return;
    }

    CachedItem* cached = EnsureCachedItem(change->iItem);
    if (!cached || !cached->pidl) {
        return;
    }

    change->lParam = reinterpret_cast<LPARAM>(cached->pidl.get());
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

bool ShellTabsListView::HandleCustomDraw(NMLVCUSTOMDRAW* draw, LRESULT* result) {
    if (!draw || !result) {
        return false;
    }

    const DWORD stage = draw->nmcd.dwDrawStage;
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
            return true;
        }
        case CDDS_ITEMPOSTPAINT: {
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

