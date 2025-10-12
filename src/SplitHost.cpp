#include "SplitHost.h"

#include <CommCtrl.h>
#include <ShlObj.h>
#include <windowsx.h>
#include <algorithm>
#include <array>
#include <atomic>
#include <cwchar>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <map>
#include <memory>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "FileColorOverrides.h"
#include "Utilities.h"

namespace shelltabs {
namespace {

constexpr wchar_t kSplitHostClass[] = L"ShellTabs.SplitHost";
constexpr int kSplitterWidth = 6;
constexpr int kAddressBarHeight = 28;
constexpr COLORREF kUniqueColor = RGB(138, 43, 226);   // purple accent
constexpr COLORREF kDifferenceColor = RGB(220, 0, 0);  // red for mismatched files
constexpr UINT WM_COMPARE_COMPLETE = WM_APP + 0x452;

struct CaseInsensitiveLess {
    bool operator()(const std::wstring& left, const std::wstring& right) const {
        return _wcsicmp(left.c_str(), right.c_str()) < 0;
    }
};

using EntryMap = std::map<std::wstring, std::filesystem::directory_entry, CaseInsensitiveLess>;

EntryMap CollectEntries(const std::filesystem::path& folder, std::stop_token stop, bool* hadError) {
    EntryMap entries;
    if (hadError) {
        *hadError = false;
    }

    std::error_code ec;
    for (std::filesystem::directory_iterator it(folder, std::filesystem::directory_options::skip_permission_denied, ec);
         !ec && it != std::filesystem::directory_iterator(); it.increment(ec)) {
        if (stop.stop_requested()) {
            break;
        }
        const std::wstring name = it->path().filename().wstring();
        entries.emplace(name, *it);
    }

    if (ec && hadError) {
        *hadError = true;
    }

    return entries;
}

bool FilesDifferent(const std::filesystem::path& left, const std::filesystem::path& right, std::stop_token stop) {
    std::error_code ecLeft;
    std::error_code ecRight;
    const auto leftSize = std::filesystem::file_size(left, ecLeft);
    const auto rightSize = std::filesystem::file_size(right, ecRight);
    if (ecLeft || ecRight || leftSize != rightSize) {
        return true;
    }

    std::ifstream leftStream(left, std::ios::binary);
    std::ifstream rightStream(right, std::ios::binary);
    if (!leftStream || !rightStream) {
        return true;
    }

    std::array<char, 64 * 1024> leftBuffer{};
    std::array<char, 64 * 1024> rightBuffer{};
    auto remaining = leftSize;
    while (remaining > 0) {
        if (stop.stop_requested()) {
            return false;
        }
        const auto chunk = static_cast<std::streamsize>(std::min<uintmax_t>(remaining, leftBuffer.size()));
        leftStream.read(leftBuffer.data(), chunk);
        rightStream.read(rightBuffer.data(), chunk);
        if (leftStream.gcount() != chunk || rightStream.gcount() != chunk) {
            return true;
        }
        if (std::memcmp(leftBuffer.data(), rightBuffer.data(), static_cast<size_t>(chunk)) != 0) {
            return true;
        }
        remaining -= static_cast<uintmax_t>(chunk);
    }
    return false;
}

bool CompareRecursive(const std::filesystem::path& left, const std::filesystem::path& right,
                      const std::filesystem::path& relative, SplitHost::CompareResult& result, std::stop_token stop) {
    if (stop.stop_requested()) {
        return false;
    }

    bool leftError = false;
    bool rightError = false;
    EntryMap leftEntries = CollectEntries(left, stop, &leftError);
    EntryMap rightEntries = CollectEntries(right, stop, &rightError);

    bool hasDifference = leftError || rightError;
    if (leftError) {
        result.folderDiffLeft.push_back(left.wstring());
    }
    if (rightError) {
        result.folderDiffRight.push_back(right.wstring());
    }

    const auto makeAbsolute = [](const std::filesystem::path& root, const std::filesystem::path& rel) {
        return (root / rel).wstring();
    };

    for (auto& [name, leftEntry] : leftEntries) {
        if (stop.stop_requested()) {
            return hasDifference;
        }

        const auto itRight = rightEntries.find(name);
        const std::filesystem::path rel = relative / leftEntry.path().filename();
        if (itRight == rightEntries.end()) {
            hasDifference = true;
            result.leftOnly.push_back(makeAbsolute(left, rel));
            continue;
        }

        const auto rightEntry = itRight->second;
        rightEntries.erase(itRight);

        std::error_code leftTypeError;
        std::error_code rightTypeError;
        const bool leftIsDir = leftEntry.is_directory(leftTypeError);
        const bool rightIsDir = rightEntry.is_directory(rightTypeError);

        if (leftTypeError || rightTypeError) {
            hasDifference = true;
            result.folderDiffLeft.push_back(makeAbsolute(left, rel));
            result.folderDiffRight.push_back(makeAbsolute(right, rel));
            continue;
        }

        if (leftIsDir && rightIsDir) {
            if (CompareRecursive(leftEntry.path(), rightEntry.path(), rel, result, stop)) {
                hasDifference = true;
                result.folderDiffLeft.push_back(makeAbsolute(left, rel));
                result.folderDiffRight.push_back(makeAbsolute(right, rel));
            }
            continue;
        }

        const bool leftFile = leftEntry.is_regular_file(leftTypeError);
        const bool rightFile = rightEntry.is_regular_file(rightTypeError);
        if (leftFile && rightFile) {
            if (FilesDifferent(leftEntry.path(), rightEntry.path(), stop)) {
                hasDifference = true;
                result.differingLeft.push_back(makeAbsolute(left, rel));
                result.differingRight.push_back(makeAbsolute(right, rel));
            }
            continue;
        }

        // Type mismatch (file vs directory) counts as a structural difference.
        hasDifference = true;
        result.folderDiffLeft.push_back(makeAbsolute(left, rel));
        result.folderDiffRight.push_back(makeAbsolute(right, rel));
    }

    for (auto& [name, rightEntry] : rightEntries) {
        if (stop.stop_requested()) {
            return hasDifference;
        }
        const std::filesystem::path rel = relative / rightEntry.path().filename();
        hasDifference = true;
        result.rightOnly.push_back(makeAbsolute(right, rel));
    }

    return hasDifference;
}

std::unique_ptr<SplitHost::CompareResult> CompareDirectories(const std::wstring& left, const std::wstring& right,
                                                             std::stop_token stop) {
    auto result = std::make_unique<SplitHost::CompareResult>();
    try {
        CompareRecursive(std::filesystem::path(left), std::filesystem::path(right), std::filesystem::path(), *result, stop);
    } catch (...) {
        result->folderDiffLeft.push_back(left);
        result->folderDiffRight.push_back(right);
    }
    return result;
}

}  // namespace

SplitHost::~SplitHost() {
    ResetComparison();
    if (m_leftAddress && IsWindow(m_leftAddress)) {
        RemoveWindowSubclass(m_leftAddress, AddressBarSubclassProc, 1);
    }
    if (m_rightAddress && IsWindow(m_rightAddress)) {
        RemoveWindowSubclass(m_rightAddress, AddressBarSubclassProc, 2);
    }
    FileColorOverrides::Instance().ClearEphemeral();
    m_left.Destroy();
    m_right.Destroy();
}

SplitHost* SplitHost::CreateAndAttach(HWND parent) {
    static std::atomic<bool> registered{false};
    bool expected = false;
    if (registered.compare_exchange_strong(expected, true)) {
        WNDCLASSW wc{};
        wc.lpfnWndProc = SplitHost::WndProc;
        wc.hInstance = static_cast<HINSTANCE>(GetModuleHandleW(nullptr));
        wc.lpszClassName = kSplitHostClass;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassW(&wc);
    }

    RECT rc{};
    GetClientRect(parent, &rc);
    HWND hwnd = CreateWindowExW(0, kSplitHostClass, L"", WS_CHILD | WS_VISIBLE, 0, 0, rc.right, rc.bottom, parent, nullptr,
                               static_cast<HINSTANCE>(GetModuleHandleW(nullptr)), nullptr);
    if (!hwnd) {
        return nullptr;
    }
    return FromHwnd(hwnd);
}

SplitHost* SplitHost::FromHwnd(HWND hwnd) {
    return reinterpret_cast<SplitHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

void SplitHost::DestroyIfExistsOn(HWND parent) {
    for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT)) {
        wchar_t className[128]{};
        GetClassNameW(child, className, ARRAYSIZE(className));
        if (lstrcmpiW(className, kSplitHostClass) == 0) {
            DestroyWindow(child);
            return;
        }
    }
}

LRESULT CALLBACK SplitHost::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_NCCREATE) {
        auto* self = new SplitHost();
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->m_hwnd = hwnd;
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    auto* self = FromHwnd(hwnd);
    if (!self) {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    switch (msg) {
    case WM_CREATE:
        return self->OnCreate(hwnd);
    case WM_SIZE:
        self->OnSize();
        break;
    case WM_LBUTTONDOWN:
        self->OnLButtonDown(GET_X_LPARAM(lp));
        break;
    case WM_MOUSEMOVE:
        self->OnMouseMove(GET_X_LPARAM(lp));
        break;
    case WM_LBUTTONUP:
        self->OnLButtonUp();
        break;
    case WM_COMPARE_COMPLETE: {
        auto token = static_cast<uint64_t>(wp);
        auto result = std::unique_ptr<CompareResult>(reinterpret_cast<CompareResult*>(lp));
        self->ApplyComparisonResult(token, std::move(result));
        return 0;
    }
    case WM_DESTROY:
        delete self;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        break;
    default:
        break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT SplitHost::OnCreate(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);

    m_leftAddress = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 0, 0,
                                   rc.right / 2, kAddressBarHeight, hwnd, nullptr, nullptr, nullptr);
    m_rightAddress = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, 0, 0,
                                    rc.right / 2, kAddressBarHeight, hwnd, nullptr, nullptr, nullptr);

    HFONT font = reinterpret_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    if (font) {
        SendMessageW(m_leftAddress, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
        SendMessageW(m_rightAddress, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    }

    m_leftAddressContext = {this, true};
    m_rightAddressContext = {this, false};
    SetWindowSubclass(m_leftAddress, AddressBarSubclassProc, 1, reinterpret_cast<DWORD_PTR>(&m_leftAddressContext));
    SetWindowSubclass(m_rightAddress, AddressBarSubclassProc, 2, reinterpret_cast<DWORD_PTR>(&m_rightAddressContext));

    if (FAILED(m_left.Create(hwnd, rc)) || FAILED(m_right.Create(hwnd, rc))) {
        return -1;
    }

    m_left.SetNavigationCallback([this](const std::wstring& path) { OnPaneNavigated(true, path); });
    m_right.SetNavigationCallback([this](const std::wstring& path) { OnPaneNavigated(false, path); });

    m_splitter = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_NOTIFY, 0, 0, kSplitterWidth, rc.bottom,
                                hwnd, nullptr, nullptr, nullptr);
    SetWindowPos(m_splitter, HWND_TOP, m_splitX, kAddressBarHeight, kSplitterWidth, rc.bottom - kAddressBarHeight,
                 SWP_NOACTIVATE);

    LayoutChildren();
    return 0;
}

void SplitHost::OnSize() { LayoutChildren(); }

void SplitHost::LayoutChildren() {
    RECT rc{};
    GetClientRect(m_hwnd, &rc);
    const int width = std::max(0, rc.right - rc.left);
    const int minPaneWidth = 150;
    const int usableWidth = std::max(width - kSplitterWidth, minPaneWidth * 2);
    const int maxSplit = usableWidth - minPaneWidth;
    if (m_splitX < minPaneWidth) {
        m_splitX = minPaneWidth;
    } else if (m_splitX > maxSplit) {
        m_splitX = maxSplit;
    }

    const int rightStart = m_splitX + kSplitterWidth;
    const int contentHeight = std::max(0, rc.bottom - kAddressBarHeight);

    MoveWindow(m_leftAddress, 0, 0, std::max(m_splitX, 0), kAddressBarHeight, TRUE);
    MoveWindow(m_rightAddress, rightStart, 0, std::max(width - rightStart, 0), kAddressBarHeight, TRUE);

    RECT leftRect{0, kAddressBarHeight, m_splitX, rc.bottom};
    RECT rightRect{rightStart, kAddressBarHeight, rc.right, rc.bottom};
    m_left.SetRect(leftRect);
    m_right.SetRect(rightRect);

    SetWindowPos(m_splitter, HWND_TOP, m_splitX, kAddressBarHeight, kSplitterWidth, contentHeight,
                 SWP_NOACTIVATE | SWP_SHOWWINDOW);
}

void SplitHost::OnLButtonDown(int x) {
    RECT splitterRect{};
    GetWindowRect(m_splitter, &splitterRect);
    POINT pt{x, 0};
    ClientToScreen(m_hwnd, &pt);
    if (pt.x >= splitterRect.left && pt.x <= splitterRect.right) {
        SetCapture(m_hwnd);
        m_drag = true;
    }
}

void SplitHost::OnMouseMove(int x) {
    if (!m_drag) {
        return;
    }
    m_splitX = x;
    LayoutChildren();
}

void SplitHost::OnLButtonUp() {
    if (!m_drag) {
        return;
    }
    m_drag = false;
    ReleaseCapture();
}

void SplitHost::SetFolders(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right) {
    bool cleared = false;
    if (left) {
        m_left.NavigateToPIDL(left);
    } else {
        m_leftPath.clear();
        UpdateAddressBar(true, L"");
        cleared = true;
    }
    if (right) {
        m_right.NavigateToPIDL(right);
    } else {
        m_rightPath.clear();
        UpdateAddressBar(false, L"");
        cleared = true;
    }
    if (cleared) {
        ScheduleComparison();
    }
}

void SplitHost::Swap() {
    RECT rc{};
    GetClientRect(m_hwnd, &rc);
    const int width = rc.right - rc.left;
    m_splitX = width - m_splitX - kSplitterWidth;
    LayoutChildren();
}

void SplitHost::OnPaneNavigated(bool isLeft, const std::wstring& path) {
    if (isLeft) {
        m_leftPath = path;
    } else {
        m_rightPath = path;
    }
    UpdateAddressBar(isLeft, path);
    ScheduleComparison();
}

void SplitHost::UpdateAddressBar(bool isLeft, const std::wstring& path) {
    HWND edit = isLeft ? m_leftAddress : m_rightAddress;
    if (!edit) {
        return;
    }
    const std::wstring display = path;
    SetWindowTextW(edit, display.c_str());
}

bool SplitHost::NavigateFromAddressBar(bool isLeft) {
    HWND edit = isLeft ? m_leftAddress : m_rightAddress;
    if (!edit) {
        return false;
    }
    const int length = GetWindowTextLengthW(edit);
    std::wstring buffer;
    buffer.resize(static_cast<size_t>(length > 0 ? length : 0));
    if (length > 0) {
        GetWindowTextW(edit, buffer.data(), length + 1);
        buffer.resize(wcslen(buffer.c_str()));
    } else {
        buffer.clear();
    }
    if (buffer.empty()) {
        MessageBeep(MB_ICONWARNING);
        return false;
    }

    HRESULT hr = isLeft ? m_left.NavigateToPath(buffer) : m_right.NavigateToPath(buffer);
    if (FAILED(hr)) {
        MessageBeep(MB_ICONERROR);
        return false;
    }
    return true;
}

void SplitHost::ScheduleComparison() {
    ResetComparison();

    FileColorOverrides::Instance().ClearEphemeral();
    InvalidatePanes();

    if (m_leftPath.empty() || m_rightPath.empty()) {
        return;
    }

    const uint64_t token = ++m_compareToken;
    m_latestToken = token;
    const std::wstring left = m_leftPath;
    const std::wstring right = m_rightPath;

    m_compareThread = std::jthread([this, left, right, token](std::stop_token stop) {
        auto result = CompareDirectories(left, right, stop);
        if (!result || stop.stop_requested()) {
            return;
        }
        PostMessageW(m_hwnd, WM_COMPARE_COMPLETE, static_cast<WPARAM>(token),
                     reinterpret_cast<LPARAM>(result.release()));
    });
}

void SplitHost::ResetComparison() {
    if (m_compareThread.joinable()) {
        m_compareThread.request_stop();
        m_compareThread.join();
        m_compareThread = std::jthread();
    }
}

void SplitHost::ApplyComparisonResult(uint64_t token, std::unique_ptr<CompareResult> result) {
    if (!result || token != m_latestToken.load()) {
        return;
    }

    FileColorOverrides& overrides = FileColorOverrides::Instance();
    overrides.ClearEphemeral();

    std::vector<std::wstring> purple;
    purple.reserve(result->leftOnly.size() + result->rightOnly.size() + result->folderDiffLeft.size() +
                   result->folderDiffRight.size());
    purple.insert(purple.end(), result->leftOnly.begin(), result->leftOnly.end());
    purple.insert(purple.end(), result->rightOnly.begin(), result->rightOnly.end());
    purple.insert(purple.end(), result->folderDiffLeft.begin(), result->folderDiffLeft.end());
    purple.insert(purple.end(), result->folderDiffRight.begin(), result->folderDiffRight.end());

    std::vector<std::wstring> red;
    red.reserve(result->differingLeft.size() + result->differingRight.size());
    red.insert(red.end(), result->differingLeft.begin(), result->differingLeft.end());
    red.insert(red.end(), result->differingRight.begin(), result->differingRight.end());

    if (!purple.empty()) {
        overrides.SetEphemeralColor(purple, kUniqueColor);
    }
    if (!red.empty()) {
        overrides.SetEphemeralColor(red, kDifferenceColor);
    }

    InvalidatePanes();
}

void SplitHost::InvalidatePanes() {
    m_left.InvalidateView();
    m_right.InvalidateView();
}

LRESULT CALLBACK SplitHost::AddressBarSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR, DWORD_PTR refData) {
    auto* context = reinterpret_cast<AddressBarContext*>(refData);
    if (!context || !context->host) {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg) {
    case WM_KEYDOWN:
        if (wp == VK_RETURN) {
            context->host->NavigateFromAddressBar(context->isLeft);
            return 0;
        }
        break;
    case WM_DESTROY:
    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, AddressBarSubclassProc, context->isLeft ? 1 : 2);
        break;
    default:
        break;
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

}  // namespace shelltabs
