#include "HttpShellFolder.h"

#include "ComUtils.h"
#include "Guids.h"
#include "HtmlDirectoryParser.h"
#include "Logging.h"
#include "NamespaceContextMenu.h"
#include "OptionsStore.h"

#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl_core.h>
#include <shlwapi.h>

#include <algorithm>
#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <cwchar>
#include <cwctype>
#include <iterator>
#include <memory>
#include <mutex>
#include <string_view>
#include <type_traits>
#include <thread>
#include <unordered_map>
#include <vector>

#include <commctrl.h>
#include <propvarutil.h>
#include <strsafe.h>

#ifdef _MSC_VER
#pragma comment(lib, "propsys.lib")
#endif

using Microsoft::WRL::ComPtr;

namespace shelltabs::http {

namespace {

#ifdef SHCONTF_ALLFOLDERS
constexpr SHCONTF kShcontfAllFolders = SHCONTF_ALLFOLDERS;
#else
constexpr SHCONTF kShcontfAllFolders = static_cast<SHCONTF>(0x00000080);
#endif

template <typename>
inline constexpr bool kDependentFalse = false;

template <typename SFVCreate>
void AssignFolderSettings(SFVCreate& create, const FOLDERSETTINGS& settings) {
    if constexpr (requires(SFVCreate& candidate) { candidate.pfs = &settings; }) {
        create.pfs = &settings;
    } else if constexpr (requires(SFVCreate& candidate) { candidate.psfs = &settings; }) {
        create.psfs = &settings;
    } else if constexpr (requires(SFVCreate& candidate) { candidate.pfolderSettings = &settings; }) {
        create.pfolderSettings = &settings;
    } else if constexpr (requires(SFVCreate& candidate) { candidate.pViewSettings = &settings; }) {
        create.pViewSettings = &settings;
    } else if constexpr (requires(SFVCreate& candidate) { candidate.pFolderSettings = &settings; }) {
        create.pFolderSettings = &settings;
    } else if constexpr (std::is_same_v<std::remove_cvref_t<SFVCreate>, SFV_CREATE>) {
        (void)settings;
    } else {
        static_assert(kDependentFalse<SFVCreate>, "SFV_CREATE is missing a folder settings member");
    }
}

std::wstring JoinSegments(const std::vector<std::wstring>& base, const std::vector<std::wstring>& extra) {
    std::wstring path;
    auto append = [&](const std::wstring& segment) {
        if (path.empty()) {
            path = L"/" + segment;
        } else {
            path.push_back(L'/');
            path += segment;
        }
    };
    for (const auto& segment : base) {
        append(segment);
    }
    for (const auto& segment : extra) {
        append(segment);
    }
    if (path.empty()) {
        path = L"/";
    }
    return path;
}

void SplitAndAppend(const std::wstring& path, std::vector<std::wstring>& out) {
    size_t start = 0;
    while (start < path.size()) {
        size_t sep = path.find(L'/', start);
        if (sep == std::wstring::npos) {
            out.push_back(path.substr(start));
            break;
        }
        if (sep > start) {
            out.push_back(path.substr(start, sep - start));
        }
        start = sep + 1;
    }
}

// ---------------------------------------------------------------------------
// Navigation tracker: Explorer reuses the site-root IShellFolder's DefView
// for subfolder navigation but polls GetCurFolder on the site-root to build
// absolute PIDLs for further navigation.  The site-root's absolutePidl_ never
// changes, so the intermediate path segments are lost.  This tracker is keyed
// by the canonical site URL and updated by EnumObjects whenever a folder for
// that site enumerates.  GetCurFolder on the site-root reads the tracker so
// Explorer gets the correct current-location PIDL.
// ---------------------------------------------------------------------------
struct ViewNavigationTracker {
    std::mutex mutex;
    UniquePidl absolutePidl;
    std::vector<std::wstring> pathSegments;
};

static std::mutex s_navTrackerMutex;
static std::unordered_map<std::wstring, std::shared_ptr<ViewNavigationTracker>> s_navTrackers;

static std::wstring MakeNavTrackerKey(const HttpUrlParts& parts) {
    std::wstring key = parts.useHttps ? L"https://" : L"http://";
    key += parts.host;
    key += L":";
    key += std::to_wstring(parts.port);
    key += parts.basePath;
    return key;
}

// Retrieve the current navigation state from the tracker for a given site.
// Returns true if the tracker was found and had valid state.
static bool GetNavTrackerState(const HttpUrlParts& parts,
                               std::vector<std::wstring>* outSegments,
                               UniquePidl* outPidl) {
    if (parts.host.empty()) return false;
    auto key = MakeNavTrackerKey(parts);
    std::shared_ptr<ViewNavigationTracker> tracker;
    {
        std::lock_guard lock(s_navTrackerMutex);
        auto it = s_navTrackers.find(key);
        if (it != s_navTrackers.end()) {
            tracker = it->second;
        }
    }
    if (!tracker) return false;
    std::lock_guard lock(tracker->mutex);
    if (outSegments) *outSegments = tracker->pathSegments;
    if (outPidl && tracker->absolutePidl) *outPidl = ClonePidl(tracker->absolutePidl.get());
    return true;
}

#ifndef SFVM_MERGEMENU
#define SFVM_MERGEMENU 1
#endif
#ifndef SFVM_INVOKECOMMAND
#define SFVM_INVOKECOMMAND 2
#endif
#ifndef SFVM_DBLCLK
#define SFVM_DBLCLK 19
#endif
#ifndef SFVM_GETNOTIFY
#define SFVM_GETNOTIFY 49
#endif

bool ContainsIgnoreCase(const std::wstring& haystack, const std::wstring& needle) {
    if (needle.empty()) return true;
    if (haystack.size() < needle.size()) return false;
    for (size_t i = 0; i <= haystack.size() - needle.size(); ++i) {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j) {
            if (towlower(haystack[i + j]) != towlower(needle[j])) {
                match = false;
                break;
            }
        }
        if (match) return true;
    }
    return false;
}

struct FilterDialogData {
    const wchar_t* prompt;
    const wchar_t* defaultValue;
    std::wstring result;
    bool accepted;
};

INT_PTR CALLBACK FilterDialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_INITDIALOG: {
            SetWindowLongPtrW(hwnd, DWLP_USER, lParam);
            auto* data = reinterpret_cast<FilterDialogData*>(lParam);
            if (data && data->defaultValue) {
                SetDlgItemTextW(hwnd, 101, data->defaultValue);
            }
            return TRUE;
        }
        case WM_COMMAND:
            if (LOWORD(wParam) == IDOK) {
                auto* data = reinterpret_cast<FilterDialogData*>(GetWindowLongPtrW(hwnd, DWLP_USER));
                if (data) {
                    wchar_t buf[512]{};
                    GetDlgItemTextW(hwnd, 101, buf, ARRAYSIZE(buf));
                    data->result = buf;
                    data->accepted = true;
                }
                EndDialog(hwnd, IDOK);
                return TRUE;
            }
            if (LOWORD(wParam) == IDCANCEL) {
                EndDialog(hwnd, IDCANCEL);
                return TRUE;
            }
            break;
    }
    return FALSE;
}

bool ShowFilterDialog(HWND parent, const wchar_t* defaultValue, std::wstring* result) {
    if (!result) return false;
    result->clear();

    FilterDialogData data{L"Filter items by name:", defaultValue, {}, false};

    alignas(4) BYTE buffer[2048]{};
    BYTE* ptr = buffer;
    auto WriteWord = [&](WORD v) { memcpy(ptr, &v, 2); ptr += 2; };
    auto WriteWStr = [&](const wchar_t* s) { size_t n = (wcslen(s) + 1) * 2; memcpy(ptr, s, n); ptr += n; };
    auto Align4 = [&]() { while ((reinterpret_cast<uintptr_t>(ptr) & 3) != 0) *ptr++ = 0; };

    auto* dlg = reinterpret_cast<DLGTEMPLATE*>(ptr);
    dlg->style = DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE | DS_SETFONT;
    dlg->cdit = 4; dlg->x = 0; dlg->y = 0; dlg->cx = 220; dlg->cy = 75;
    ptr += sizeof(DLGTEMPLATE);
    WriteWord(0); WriteWord(0);
    WriteWStr(L"Filter");
    WriteWord(8); WriteWStr(L"MS Shell Dlg");
    Align4();

    auto* item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
    item->style = WS_CHILD | WS_VISIBLE | SS_LEFT;
    item->x = 7; item->y = 7; item->cx = 206; item->cy = 12; item->id = 100;
    ptr += sizeof(DLGITEMTEMPLATE);
    WriteWord(0xFFFF); WriteWord(0x0082);
    WriteWStr(L"Filter items by name (leave empty to clear):");
    WriteWord(0); Align4();

    item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
    item->style = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL;
    item->dwExtendedStyle = WS_EX_CLIENTEDGE;
    item->x = 7; item->y = 22; item->cx = 206; item->cy = 14; item->id = 101;
    ptr += sizeof(DLGITEMTEMPLATE);
    WriteWord(0xFFFF); WriteWord(0x0081);
    WriteWStr(defaultValue ? defaultValue : L"");
    WriteWord(0); Align4();

    item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
    item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON;
    item->x = 108; item->y = 44; item->cx = 50; item->cy = 14; item->id = IDOK;
    ptr += sizeof(DLGITEMTEMPLATE);
    WriteWord(0xFFFF); WriteWord(0x0080);
    WriteWStr(L"OK"); WriteWord(0); Align4();

    item = reinterpret_cast<DLGITEMTEMPLATE*>(ptr);
    item->style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON;
    item->x = 163; item->y = 44; item->cx = 50; item->cy = 14; item->id = IDCANCEL;
    ptr += sizeof(DLGITEMTEMPLATE);
    WriteWord(0xFFFF); WriteWord(0x0080);
    WriteWStr(L"Cancel"); WriteWord(0);

    INT_PTR ret = DialogBoxIndirectParamW(nullptr, reinterpret_cast<DLGTEMPLATE*>(buffer),
                                          parent, FilterDialogProc, reinterpret_cast<LPARAM>(&data));
    if (ret == IDOK && data.accepted) {
        *result = std::move(data.result);
        return true;
    }
    return false;
}

bool EqualsIgnoreCase(std::wstring_view left, std::wstring_view right) {
    if (left.size() != right.size()) {
        return false;
    }
    for (size_t i = 0; i < left.size(); ++i) {
        if (towlower(left[i]) != towlower(right[i])) {
            return false;
        }
    }
    return true;
}

struct ColumnDefinition {
    PROPERTYKEY key;
    const wchar_t* name;
    int format = LVCFMT_LEFT;
    UINT width = 24;
    SHCOLSTATEF state = SHCOLSTATE_ONBYDEFAULT;
};

const ColumnDefinition kColumnDefinitions[] = {
    {PKEY_ItemNameDisplay, L"Name", LVCFMT_LEFT, 30, SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_STR},
    {PKEY_Size, L"Size", LVCFMT_RIGHT, 16, SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_INT | SHCOLSTATE_SECONDARYUI},
    {PKEY_DateModified, L"Date modified", LVCFMT_LEFT, 24, SHCOLSTATE_ONBYDEFAULT | SHCOLSTATE_TYPE_DATE},
};

constexpr size_t kColumnCount = std::size(kColumnDefinitions);

HRESULT AssignToStrRet(const std::wstring& value, STRRET* str) {
    if (!str) {
        return E_POINTER;
    }
    wchar_t* buffer = static_cast<wchar_t*>(CoTaskMemAlloc((value.size() + 1) * sizeof(wchar_t)));
    if (!buffer) {
        return E_OUTOFMEMORY;
    }
    HRESULT hr = StringCchCopyW(buffer, value.size() + 1, value.c_str());
    if (FAILED(hr)) {
        CoTaskMemFree(buffer);
        return hr;
    }
    str->uType = STRRET_WSTR;
    str->pOleStr = buffer;
    return S_OK;
}

ULONGLONG GetFileSizeFromFindData(const WIN32_FIND_DATAW& data) {
    return (static_cast<ULONGLONG>(data.nFileSizeHigh) << 32) | data.nFileSizeLow;
}

std::wstring FormatSizeString(ULONGLONG size) {
    if (size == 0) {
        return L"0 bytes";
    }
    wchar_t buffer[64];
    if (!StrFormatByteSizeW(static_cast<LONGLONG>(size), buffer, ARRAYSIZE(buffer))) {
        return {};
    }
    return buffer;
}

std::wstring FormatDateString(const FILETIME& fileTime) {
    if (fileTime.dwHighDateTime == 0 && fileTime.dwLowDateTime == 0) {
        return {};
    }
    FILETIME localTime;
    if (!FileTimeToLocalFileTime(&fileTime, &localTime)) {
        return {};
    }
    SYSTEMTIME systemTime;
    if (!FileTimeToSystemTime(&localTime, &systemTime)) {
        return {};
    }
    wchar_t dateBuffer[64];
    wchar_t timeBuffer[64];
    int dateLength = GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &systemTime, nullptr, dateBuffer,
                                     ARRAYSIZE(dateBuffer), nullptr);
    if (dateLength == 0) {
        return {};
    }
    int timeLength = GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, 0, &systemTime, nullptr, timeBuffer,
                                     ARRAYSIZE(timeBuffer));
    std::wstring formatted(dateBuffer);
    if (timeLength != 0) {
        formatted.push_back(L' ');
        formatted.append(timeBuffer);
    }
    return formatted;
}

ULONG MapFindDataToAttributes(const WIN32_FIND_DATAW& data) {
    ULONG attributes = SFGAO_STORAGE | SFGAO_READONLY | SFGAO_CANCOPY;
    const bool isDirectory = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    if (isDirectory) {
        attributes |= SFGAO_FOLDER | SFGAO_HASSUBFOLDER | SFGAO_BROWSABLE | SFGAO_FILESYSANCESTOR | SFGAO_FILESYSTEM;
    } else {
        attributes |= SFGAO_STREAM;
    }
    return attributes;
}

bool TryGetNameFromPidl(PCUIDLIST_RELATIVE pidl, std::wstring* name) {
    if (!pidl || !name) {
        return false;
    }
    if (pidl->mkid.cb == 0) {
        return false;
    }
    return TryGetComponentString(pidl->mkid, ComponentType::Name, name);
}

// Background enumeration for HTTP directories
class EnumerationState : public std::enable_shared_from_this<EnumerationState> {
public:
    EnumerationState(const HttpUrlParts& parts, std::vector<std::wstring> segments,
                     std::vector<std::uint8_t> absolute, SHCONTF flags, HWND owner,
                     std::wstring filter = {})
        : rootParts_(parts),
          pathSegments_(std::move(segments)),
          absolutePidlBytes_(std::move(absolute)),
          flags_(flags),
          ownerWindow_(owner),
          filterString_(std::move(filter)) {}

    ~EnumerationState() {
        Cancel();
        if (worker_.joinable()) {
            if (std::this_thread::get_id() == workerThreadId_) {
                worker_.detach();
            } else {
                worker_.join();
            }
        }
    }

    void Start() {
        worker_ = std::thread([self = shared_from_this()]() {
            self->workerThreadId_ = std::this_thread::get_id();
            self->WorkerProc();
        });
    }

    void Cancel() {
        cancelled_.store(true, std::memory_order_release);
        cv_.notify_all();
    }

    HRESULT GetItem(size_t index, std::vector<std::uint8_t>* bytes, bool* hasItem) {
        if (hasItem) {
            *hasItem = false;
        }
        std::unique_lock<std::mutex> lock(mutex_);
        cv_.wait(lock, [&]() { return index < items_.size() || workerFinished_ || cancelled_.load(std::memory_order_acquire); });
        if (index < items_.size()) {
            auto data = items_[index].bytes;
            lock.unlock();
            if (bytes) {
                *bytes = std::move(data);
            }
            if (hasItem) {
                *hasItem = true;
            }
            return S_OK;
        }
        if (cancelled_.load(std::memory_order_acquire)) {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        HRESULT result = result_;
        lock.unlock();
        if (FAILED(result)) {
            return result;
        }
        return S_FALSE;
    }

private:
    struct ItemBuffer {
        std::vector<std::uint8_t> bytes;
    };

    void WorkerProc() {
        LogMessage(LogLevel::Info, L"[HttpShellFolder] WorkerProc: started for host=%ls basePath=%ls segments=%zu",
                   rootParts_.host.c_str(), rootParts_.basePath.c_str(), pathSegments_.size());
        HRESULT hr = S_OK;
        bool coInitialized = false;
        HRESULT init = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (SUCCEEDED(init)) {
            coInitialized = true;
        }

        HttpConnectionOptions options;
        options.host = rootParts_.host;
        options.port = rootParts_.port;
        options.basePath = rootParts_.basePath;
        options.useHttps = rootParts_.useHttps;

        std::wstring remotePath = JoinSegments(pathSegments_, {});
        LogMessage(LogLevel::Info, L"[HttpShellFolder] WorkerProc: ListDirectory host=%ls basePath=%ls remotePath=%ls https=%d port=%d",
                   options.host.c_str(), options.basePath.c_str(), remotePath.c_str(),
                   options.useHttps ? 1 : 0, static_cast<int>(options.port));

        HttpClient client;
        std::vector<DirectoryEntry> entries;
        hr = client.ListDirectory(options, remotePath, &entries);

        LogMessage(LogLevel::Info, L"[HttpShellFolder] WorkerProc: ListDirectory returned hr=0x%08X entries=%zu",
                   hr, entries.size());

        if (SUCCEEDED(hr)) {
            size_t included = 0;
            for (const auto& entry : entries) {
                if (cancelled_.load(std::memory_order_acquire)) {
                    hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    break;
                }
                if (!ShouldInclude(entry)) {
                    continue;
                }
                ++included;
                WIN32_FIND_DATAW findData{};
                findData.dwFileAttributes = entry.isDirectory ? FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY
                                                              : FILE_ATTRIBUTE_ARCHIVE | FILE_ATTRIBUTE_READONLY;
                findData.ftCreationTime = entry.lastWriteTime;
                findData.ftLastAccessTime = entry.lastWriteTime;
                findData.ftLastWriteTime = entry.lastWriteTime;
                findData.nFileSizeHigh = static_cast<DWORD>(entry.size >> 32);
                findData.nFileSizeLow = static_cast<DWORD>(entry.size & 0xFFFFFFFFULL);
                StringCchCopyW(findData.cFileName, ARRAYSIZE(findData.cFileName), entry.name.c_str());

                PidlBuilder builder;
                ComponentDefinition nameComponent{ComponentType::Name, entry.name.c_str(), entry.name.size() * sizeof(wchar_t)};
                ComponentDefinition dataComponent{ComponentType::FindData, &findData, sizeof(findData)};
                ItemType type = entry.isDirectory ? ItemType::Directory : ItemType::File;
                hr = builder.Append(type, {nameComponent, dataComponent});
                if (FAILED(hr)) {
                    LogMessage(LogLevel::Error, L"[HttpShellFolder] WorkerProc: PidlBuilder::Append failed hr=0x%08X for '%ls'",
                               hr, entry.name.c_str());
                    break;
                }
                UniquePidl pidl = builder.Finalize();
                if (!pidl) {
                    hr = E_OUTOFMEMORY;
                    break;
                }
                UINT size = ILGetSize(pidl.get());
                ItemBuffer buffer;
                buffer.bytes.resize(size);
                std::memcpy(buffer.bytes.data(), pidl.get(), size);
                {
                    std::lock_guard<std::mutex> guard(mutex_);
                    items_.push_back(std::move(buffer));
                }
                cv_.notify_all();
            }
            LogMessage(LogLevel::Info, L"[HttpShellFolder] WorkerProc: produced %zu items from %zu entries",
                       included, entries.size());
        }

        if (cancelled_.load(std::memory_order_acquire) && SUCCEEDED(hr)) {
            hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        // Note: Do NOT call SHChangeNotify(SHCNE_UPDATEDIR) here.
        // The SFVM_GETNOTIFY handler subscribes to SHCNE_UPDATEDIR for our own
        // PIDL, so firing it would cause DefView to re-enumerate us in a loop.

        {
            std::lock_guard<std::mutex> guard(mutex_);
            workerFinished_ = true;
            result_ = hr;
        }
        cv_.notify_all();
        LogMessage(LogLevel::Info, L"[HttpShellFolder] WorkerProc: finished, final hr=0x%08X", hr);

        if (coInitialized) {
            CoUninitialize();
        }
    }

    bool ShouldInclude(const DirectoryEntry& entry) const {
        const bool isDirectory = entry.isDirectory;
        if (isDirectory) {
            if ((flags_ & (SHCONTF_FOLDERS | kShcontfAllFolders)) == 0) {
                return false;
            }
        } else {
            if ((flags_ & (SHCONTF_NONFOLDERS | SHCONTF_STORAGE)) == 0) {
                return false;
            }
        }
        if (!filterString_.empty() && !ContainsIgnoreCase(entry.name, filterString_)) {
            return false;
        }
        return true;
    }

    std::mutex mutex_;
    std::condition_variable cv_;
    std::vector<ItemBuffer> items_;
    std::atomic<bool> cancelled_{false};
    bool workerFinished_ = false;
    HRESULT result_ = S_OK;
    std::thread worker_;
    std::thread::id workerThreadId_;
    HttpUrlParts rootParts_{};
    std::vector<std::wstring> pathSegments_;
    std::vector<std::uint8_t> absolutePidlBytes_;
    SHCONTF flags_ = 0;
    HWND ownerWindow_ = nullptr;
    std::wstring filterString_;
};

class HttpEnumIDList : public IEnumIDList {
public:
    explicit HttpEnumIDList(std::shared_ptr<EnumerationState> state, size_t index)
        : refCount_(1), state_(std::move(state)), currentIndex_(index) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IEnumIDList) {
            *object = static_cast<IEnumIDList*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return ++refCount_; }

    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --refCount_;
        if (count == 0) {
            delete this;
        }
        return count;
    }

    IFACEMETHODIMP Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched) override {
        if (!rgelt) {
            return E_POINTER;
        }
        if (celt > 1 && !pceltFetched) {
            return E_INVALIDARG;
        }
        for (ULONG index = 0; index < celt; ++index) {
            rgelt[index] = nullptr;
        }
        ULONG fetched = 0;
        while (fetched < celt) {
            std::vector<std::uint8_t> bytes;
            bool hasItem = false;
            HRESULT hr = state_->GetItem(currentIndex_, &bytes, &hasItem);
            if (FAILED(hr)) {
                return hr;
            }
            if (!hasItem) {
                break;
            }
            auto* pidl = static_cast<PITEMID_CHILD>(CoTaskMemAlloc(bytes.size()));
            if (!pidl) {
                return E_OUTOFMEMORY;
            }
            std::memcpy(pidl, bytes.data(), bytes.size());
            rgelt[fetched] = pidl;
            ++fetched;
            ++currentIndex_;
        }
        if (pceltFetched) {
            *pceltFetched = fetched;
        }
        return fetched == celt ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG celt) override {
        for (ULONG index = 0; index < celt; ++index) {
            bool hasItem = false;
            HRESULT hr = state_->GetItem(currentIndex_, nullptr, &hasItem);
            if (FAILED(hr)) {
                return hr;
            }
            if (!hasItem) {
                return S_FALSE;
            }
            ++currentIndex_;
        }
        return S_OK;
    }

    IFACEMETHODIMP Reset() override {
        currentIndex_ = 0;
        return S_OK;
    }

    IFACEMETHODIMP Clone(IEnumIDList** ppenum) override {
        if (!ppenum) {
            return E_POINTER;
        }
        auto clone = new (std::nothrow) HttpEnumIDList(state_, currentIndex_);
        if (!clone) {
            return E_OUTOFMEMORY;
        }
        *ppenum = clone;
        return S_OK;
    }

private:
    std::atomic<ULONG> refCount_;
    std::shared_ptr<EnumerationState> state_;
    size_t currentIndex_ = 0;
};

// Static enum for namespace root: returns PIDLs for each configured WebFolderEntry.
class StaticEnumIDList : public IEnumIDList {
public:
    explicit StaticEnumIDList(std::vector<std::vector<std::uint8_t>> items)
        : refCount_(1), items_(std::move(items)) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IEnumIDList) {
            *object = static_cast<IEnumIDList*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return ++refCount_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --refCount_;
        if (count == 0) delete this;
        return count;
    }

    IFACEMETHODIMP Next(ULONG celt, PITEMID_CHILD* rgelt, ULONG* pceltFetched) override {
        if (!rgelt) return E_POINTER;
        if (celt > 1 && !pceltFetched) return E_INVALIDARG;
        for (ULONG i = 0; i < celt; ++i) rgelt[i] = nullptr;
        ULONG fetched = 0;
        while (fetched < celt && currentIndex_ < items_.size()) {
            const auto& bytes = items_[currentIndex_];
            auto* pidl = static_cast<PITEMID_CHILD>(CoTaskMemAlloc(bytes.size()));
            if (!pidl) return E_OUTOFMEMORY;
            std::memcpy(pidl, bytes.data(), bytes.size());
            rgelt[fetched] = pidl;
            ++fetched;
            ++currentIndex_;
        }
        if (pceltFetched) *pceltFetched = fetched;
        return fetched == celt ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG celt) override {
        size_t remaining = items_.size() - currentIndex_;
        if (celt > remaining) { currentIndex_ = items_.size(); return S_FALSE; }
        currentIndex_ += celt;
        return S_OK;
    }

    IFACEMETHODIMP Reset() override { currentIndex_ = 0; return S_OK; }

    IFACEMETHODIMP Clone(IEnumIDList** ppenum) override {
        if (!ppenum) return E_POINTER;
        auto copy = new (std::nothrow) StaticEnumIDList(items_);
        if (!copy) return E_OUTOFMEMORY;
        copy->currentIndex_ = currentIndex_;
        *ppenum = copy;
        return S_OK;
    }

private:
    std::atomic<ULONG> refCount_;
    std::vector<std::vector<std::uint8_t>> items_;
    size_t currentIndex_ = 0;
};

// ---------------------------------------------------------------------------
// HttpDownloadStream — IStream that wraps a live WinHTTP GET response.
// Read() pulls data incrementally so the shell copy engine can show progress.
// ---------------------------------------------------------------------------
class HttpDownloadStream : public IStream {
public:
    HttpDownloadStream(const HttpConnectionOptions& options, const std::wstring& remotePath)
        : options_(options), remotePath_(remotePath) {}

    ~HttpDownloadStream() {
        if (request_) WinHttpCloseHandle(request_);
        if (connection_) WinHttpCloseHandle(connection_);
        if (session_) WinHttpCloseHandle(session_);
    }

    HRESULT Initialize() {
        session_ = WinHttpOpen(L"ShellTabs/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                               WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (!session_) return HRESULT_FROM_WIN32(GetLastError());

        connection_ = WinHttpConnect(session_, options_.host.c_str(), options_.port, 0);
        if (!connection_) return HRESULT_FROM_WIN32(GetLastError());

        // Build full path: basePath + remotePath
        std::wstring fullPath = options_.basePath;
        if (!fullPath.empty() && fullPath.back() != L'/') fullPath.push_back(L'/');
        if (!remotePath_.empty() && remotePath_.front() == L'/') {
            fullPath += remotePath_.substr(1);
        } else {
            fullPath += remotePath_;
        }

        DWORD flags = options_.useHttps ? WINHTTP_FLAG_SECURE : 0;
        request_ = WinHttpOpenRequest(connection_, L"GET", fullPath.c_str(), nullptr,
                                       WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
        if (!request_) return HRESULT_FROM_WIN32(GetLastError());

        DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS;
        WinHttpSetOption(request_, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

        if (!WinHttpSendRequest(request_, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                                WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (!WinHttpReceiveResponse(request_, nullptr)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        DWORD statusCode = 0;
        DWORD statusSize = sizeof(statusCode);
        WinHttpQueryHeaders(request_, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                            WINHTTP_HEADER_NAME_BY_INDEX, &statusCode, &statusSize,
                            WINHTTP_NO_HEADER_INDEX);
        if (statusCode < 200 || statusCode >= 300) {
            LogMessage(LogLevel::Error, L"[HttpDownloadStream] HTTP %lu for %ls",
                       statusCode, fullPath.c_str());
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        wchar_t clStr[64] = {};
        DWORD clSize = sizeof(clStr);
        if (WinHttpQueryHeaders(request_, WINHTTP_QUERY_CONTENT_LENGTH,
                                WINHTTP_HEADER_NAME_BY_INDEX, clStr, &clSize,
                                WINHTTP_NO_HEADER_INDEX)) {
            totalSize_ = _wtoi64(clStr);
        }

        initialized_ = true;
        return S_OK;
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IStream || riid == IID_ISequentialStream) {
            *ppv = static_cast<IStream*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return ++refCount_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --refCount_;
        if (count == 0) delete this;
        return count;
    }

    // ISequentialStream
    IFACEMETHODIMP Read(void* pv, ULONG cb, ULONG* pcbRead) override {
        if (!pv) return STG_E_INVALIDPOINTER;
        if (pcbRead) *pcbRead = 0;
        if (!initialized_ || !request_) return STG_E_READFAULT;
        if (eof_) return S_FALSE;

        DWORD bytesRead = 0;
        if (!WinHttpReadData(request_, pv, cb, &bytesRead)) {
            return STG_E_READFAULT;
        }
        if (bytesRead == 0) {
            eof_ = true;
            return S_FALSE;
        }
        bytesRead_ += bytesRead;
        if (pcbRead) *pcbRead = bytesRead;
        return S_OK;
    }
    IFACEMETHODIMP Write(const void*, ULONG, ULONG*) override { return STG_E_ACCESSDENIED; }

    // IStream
    IFACEMETHODIMP Seek(LARGE_INTEGER, DWORD, ULARGE_INTEGER*) override { return E_NOTIMPL; }
    IFACEMETHODIMP SetSize(ULARGE_INTEGER) override { return E_NOTIMPL; }
    IFACEMETHODIMP CopyTo(IStream* pstm, ULARGE_INTEGER cb,
                          ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten) override {
        if (!pstm) return STG_E_INVALIDPOINTER;
        ULONGLONG totalRead = 0, totalWritten = 0;
        ULONGLONG remaining = cb.QuadPart;
        BYTE buf[65536];
        while (remaining > 0) {
            ULONG toRead = static_cast<ULONG>((std::min)(remaining, static_cast<ULONGLONG>(sizeof(buf))));
            ULONG r = 0;
            HRESULT hr = Read(buf, toRead, &r);
            if (FAILED(hr)) return hr;
            if (r == 0) break;
            totalRead += r;
            ULONG w = 0;
            hr = pstm->Write(buf, r, &w);
            if (FAILED(hr)) return hr;
            totalWritten += w;
            remaining -= r;
        }
        if (pcbRead) pcbRead->QuadPart = totalRead;
        if (pcbWritten) pcbWritten->QuadPart = totalWritten;
        return S_OK;
    }
    IFACEMETHODIMP Commit(DWORD) override { return S_OK; }
    IFACEMETHODIMP Revert() override { return E_NOTIMPL; }
    IFACEMETHODIMP LockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
    IFACEMETHODIMP UnlockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
    IFACEMETHODIMP Stat(STATSTG* pstatstg, DWORD grfStatFlag) override {
        if (!pstatstg) return E_POINTER;
        ZeroMemory(pstatstg, sizeof(*pstatstg));
        pstatstg->type = STGTY_STREAM;
        pstatstg->cbSize.QuadPart = totalSize_;
        if (!(grfStatFlag & STATFLAG_NONAME)) pstatstg->pwcsName = nullptr;
        return S_OK;
    }
    IFACEMETHODIMP Clone(IStream**) override { return E_NOTIMPL; }

private:
    HttpConnectionOptions options_;
    std::wstring remotePath_;
    HINTERNET session_ = nullptr;
    HINTERNET connection_ = nullptr;
    HINTERNET request_ = nullptr;
    ULONGLONG totalSize_ = 0;
    ULONGLONG bytesRead_ = 0;
    bool initialized_ = false;
    bool eof_ = false;
    std::atomic<ULONG> refCount_{1};
};

// ---------------------------------------------------------------------------
// HttpFileContextMenuCallback — IContextMenuCB for file items.
// SHCreateDefaultContextMenu invokes this callback when the default verb is
// activated (double-click, Enter, or right-click → Open). This is the
// documented mechanism for namespace extensions to handle item activation.
// ---------------------------------------------------------------------------
class HttpFileContextMenuCallback : public IContextMenuCB {
public:
    HttpFileContextMenuCallback(HttpUrlParts rootParts,
                                std::vector<std::wstring> segments,
                                std::wstring fileName)
        : rootParts_(std::move(rootParts)),
          segments_(std::move(segments)),
          fileName_(std::move(fileName)) {}

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IContextMenuCB) {
            *ppv = static_cast<IContextMenuCB*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return ++refCount_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --refCount_;
        if (count == 0) delete this;
        return count;
    }

    // IContextMenuCB
    IFACEMETHODIMP CallBack(IShellFolder*, HWND hwndOwner, IDataObject*,
                            UINT uMsg, WPARAM wParam, LPARAM lParam) override {
        switch (uMsg) {
        case DFM_MERGECONTEXTMENU: {
            // Insert "Open" as the first/default item
            auto* qcm = reinterpret_cast<QCMINFO*>(lParam);
            if (qcm && qcm->hmenu) {
                openCmdId_ = qcm->idCmdFirst;
                InsertMenuW(qcm->hmenu, 0, MF_BYPOSITION | MF_STRING, openCmdId_, L"&Open");
                SetMenuDefaultItem(qcm->hmenu, 0, TRUE);
                ++qcm->idCmdFirst;
            }
            return S_OK;
        }
        case DFM_INVOKECOMMAND: {
            UINT idCmd = static_cast<UINT>(wParam);
            LogMessage(LogLevel::Info,
                L"[HttpFileContextMenuCB] DFM_INVOKECOMMAND idCmd=%u openCmdId=%u file='%ls'",
                idCmd, openCmdId_, fileName_.c_str());

            // idCmd == offset of our "Open" item (0) → download and open
            if (idCmd == 0) {
                return DownloadAndOpen(hwndOwner);
            }
            // Let the default handler process other commands
            return S_FALSE;
        }
        default:
            return S_FALSE;  // Not handled — let default processing occur
        }
    }

private:
    HRESULT DownloadAndOpen(HWND /*hwnd*/) {
        auto rootParts  = rootParts_;
        auto segments   = segments_;
        auto fileName   = fileName_;

        std::thread([rootParts = std::move(rootParts),
                     segments  = std::move(segments),
                     fileName  = std::move(fileName)]() {
            CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

            wchar_t tempDir[MAX_PATH];
            if (GetTempPathW(MAX_PATH, tempDir) == 0) {
                CoUninitialize();
                return;
            }

            std::wstring subDir = std::wstring(tempDir) + L"ShellTabs\\";
            CreateDirectoryW(subDir.c_str(), nullptr);
            std::wstring tempFile = subDir + fileName;
            DeleteFileW(tempFile.c_str());

            ComPtr<IProgressDialog> progress;
            if (SUCCEEDED(CoCreateInstance(CLSID_ProgressDialog, nullptr, CLSCTX_INPROC_SERVER,
                                           IID_PPV_ARGS(&progress)))) {
                progress->SetTitle(L"Downloading");
                std::wstring line = L"Downloading " + fileName + L"...";
                progress->SetLine(1, line.c_str(), FALSE, nullptr);
                progress->StartProgressDialog(nullptr, nullptr, PROGDLG_AUTOTIME, nullptr);
            }

            HttpConnectionOptions options;
            options.host     = rootParts.host;
            options.port     = rootParts.port;
            options.basePath = rootParts.basePath;
            options.useHttps = rootParts.useHttps;

            std::wstring remotePath = JoinSegments({}, segments);

            HttpClient client;
            HRESULT hr = client.DownloadFile(options, remotePath, tempFile,
                [&](ULONGLONG downloaded, ULONGLONG total) -> bool {
                    if (progress) {
                        progress->SetProgress64(downloaded, total);
                        return !progress->HasUserCancelled();
                    }
                    return true;
                });

            if (progress) {
                progress->StopProgressDialog();
            }

            if (SUCCEEDED(hr)) {
                LogMessage(LogLevel::Info, L"[HttpFileOpen] Downloaded '%ls', opening",
                           tempFile.c_str());
                ShellExecuteW(nullptr, L"open", tempFile.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            } else {
                LogMessage(LogLevel::Error, L"[HttpFileOpen] Download failed hr=0x%08X for '%ls'",
                           hr, remotePath.c_str());
                DeleteFileW(tempFile.c_str());
            }

            CoUninitialize();
        }).detach();

        return S_OK;
    }

    std::atomic<ULONG> refCount_{1};
    HttpUrlParts rootParts_;
    std::vector<std::wstring> segments_;
    std::wstring fileName_;
    UINT openCmdId_ = 0;
};

// ---------------------------------------------------------------------------
// HttpVirtualFileDataObject — wraps SHCreateDataObject and adds
// CFSTR_FILEDESCRIPTORW / CFSTR_FILECONTENTS for virtual copy/drag-drop.
// When the shell's copy engine requests file contents, we download on demand.
// ---------------------------------------------------------------------------
class HttpVirtualFileDataObject : public IDataObject {
public:
    struct FileEntry {
        std::wstring name;           // Display name or relative backslash path (e.g., "sub\\file.txt")
        std::wstring remotePath;     // Forward-slash relative path for download (e.g., "sub/file.txt")
        ULONGLONG size = 0;
        DWORD attributes = 0;
        FILETIME lastWriteTime{};
        bool hasSize = false;
        bool isDirectory = false;
    };

    HttpVirtualFileDataObject(ComPtr<IDataObject> inner, HttpUrlParts rootParts,
                              std::vector<std::wstring> parentSegments,
                              std::vector<FileEntry> files)
        : m_inner(std::move(inner)),
          m_rootParts(std::move(rootParts)),
          m_parentSegments(std::move(parentSegments)),
          m_files(std::move(files)) {
        m_cfDescriptor = static_cast<CLIPFORMAT>(RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW));
        m_cfContents = static_cast<CLIPFORMAT>(RegisterClipboardFormatW(CFSTR_FILECONTENTS));
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IDataObject) {
            *ppv = static_cast<IDataObject*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return ++m_refCount; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --m_refCount;
        if (count == 0) delete this;
        return count;
    }

    // IDataObject
    IFACEMETHODIMP GetData(FORMATETC* pFE, STGMEDIUM* pMedium) override {
        if (!pFE || !pMedium) return E_INVALIDARG;
        ZeroMemory(pMedium, sizeof(*pMedium));

        if (pFE->cfFormat == m_cfDescriptor && (pFE->tymed & TYMED_HGLOBAL)) {
            return GetFileDescriptor(pMedium);
        }
        if (pFE->cfFormat == m_cfContents && (pFE->tymed & TYMED_ISTREAM)) {
            return GetFileContents(pFE->lindex, pMedium);
        }

        // Delegate to inner data object for standard formats
        if (m_inner) {
            return m_inner->GetData(pFE, pMedium);
        }
        return DV_E_FORMATETC;
    }

    IFACEMETHODIMP GetDataHere(FORMATETC* pFE, STGMEDIUM* pMedium) override {
        if (m_inner) return m_inner->GetDataHere(pFE, pMedium);
        return E_NOTIMPL;
    }

    IFACEMETHODIMP QueryGetData(FORMATETC* pFE) override {
        if (!pFE) return E_INVALIDARG;
        if (pFE->cfFormat == m_cfDescriptor && (pFE->tymed & TYMED_HGLOBAL)) return S_OK;
        if (pFE->cfFormat == m_cfContents && (pFE->tymed & TYMED_ISTREAM)) return S_OK;
        if (m_inner) return m_inner->QueryGetData(pFE);
        return DV_E_FORMATETC;
    }

    IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC* pIn, FORMATETC* pOut) override {
        if (m_inner) return m_inner->GetCanonicalFormatEtc(pIn, pOut);
        if (pOut) *pOut = *pIn;
        return DATA_S_SAMEFORMATETC;
    }

    IFACEMETHODIMP SetData(FORMATETC* pFE, STGMEDIUM* pMedium, BOOL fRelease) override {
        if (m_inner) return m_inner->SetData(pFE, pMedium, fRelease);
        return E_NOTIMPL;
    }

    IFACEMETHODIMP EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppEnum) override {
        if (dwDirection != DATADIR_GET) {
            if (m_inner) return m_inner->EnumFormatEtc(dwDirection, ppEnum);
            return E_NOTIMPL;
        }

        // Build format list: our virtual formats + whatever the inner has
        std::vector<FORMATETC> formats;

        // CFSTR_FILEDESCRIPTORW
        FORMATETC feDesc{};
        feDesc.cfFormat = m_cfDescriptor;
        feDesc.dwAspect = DVASPECT_CONTENT;
        feDesc.lindex = -1;
        feDesc.tymed = TYMED_HGLOBAL;
        formats.push_back(feDesc);

        // CFSTR_FILECONTENTS (one per file, but advertised as lindex=-1)
        FORMATETC feContents{};
        feContents.cfFormat = m_cfContents;
        feContents.dwAspect = DVASPECT_CONTENT;
        feContents.lindex = -1;
        feContents.tymed = TYMED_ISTREAM;
        formats.push_back(feContents);

        // Add inner formats
        if (m_inner) {
            ComPtr<IEnumFORMATETC> innerEnum;
            if (SUCCEEDED(m_inner->EnumFormatEtc(DATADIR_GET, &innerEnum))) {
                FORMATETC fe{};
                while (innerEnum->Next(1, &fe, nullptr) == S_OK) {
                    if (fe.cfFormat != m_cfDescriptor && fe.cfFormat != m_cfContents) {
                        formats.push_back(fe);
                    }
                }
            }
        }

        return SHCreateStdEnumFmtEtc(static_cast<UINT>(formats.size()), formats.data(), ppEnum);
    }

    IFACEMETHODIMP DAdvise(FORMATETC* pFE, DWORD advf, IAdviseSink* pSink, DWORD* pdwConn) override {
        if (m_inner) return m_inner->DAdvise(pFE, advf, pSink, pdwConn);
        return OLE_E_ADVISENOTSUPPORTED;
    }
    IFACEMETHODIMP DUnadvise(DWORD dwConn) override {
        if (m_inner) return m_inner->DUnadvise(dwConn);
        return OLE_E_ADVISENOTSUPPORTED;
    }
    IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA** ppEnum) override {
        if (m_inner) return m_inner->EnumDAdvise(ppEnum);
        return OLE_E_ADVISENOTSUPPORTED;
    }

private:
    HRESULT GetFileDescriptor(STGMEDIUM* pMedium) {
        ExpandDirectories();
        UINT count = static_cast<UINT>(m_files.size());
        LogMessage(LogLevel::Info,
            L"[HttpVirtualFileDataObject] GetFileDescriptor: %u files", count);
        SIZE_T allocSize = sizeof(FILEGROUPDESCRIPTORW) +
                           (count > 0 ? (count - 1) * sizeof(FILEDESCRIPTORW) : 0);
        HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, allocSize);
        if (!hGlobal) return E_OUTOFMEMORY;

        auto* fgd = static_cast<FILEGROUPDESCRIPTORW*>(GlobalLock(hGlobal));
        if (!fgd) {
            GlobalFree(hGlobal);
            return E_OUTOFMEMORY;
        }
        fgd->cItems = count;

        for (UINT i = 0; i < count; i++) {
            auto& src = m_files[i];
            auto& fd = fgd->fgd[i];
            fd.dwFlags = static_cast<DWORD>(FD_UNICODE | FD_ATTRIBUTES | FD_PROGRESSUI);
            fd.dwFileAttributes = src.attributes;

            if (src.hasSize && !src.isDirectory) {
                fd.dwFlags |= FD_FILESIZE;
                fd.nFileSizeHigh = static_cast<DWORD>(src.size >> 32);
                fd.nFileSizeLow = static_cast<DWORD>(src.size & 0xFFFFFFFF);
            }

            if (src.lastWriteTime.dwHighDateTime != 0 || src.lastWriteTime.dwLowDateTime != 0) {
                fd.dwFlags |= FD_WRITESTIME;
                fd.ftLastWriteTime = src.lastWriteTime;
            }

            StringCchCopyW(fd.cFileName, MAX_PATH, src.name.c_str());
        }

        GlobalUnlock(hGlobal);
        pMedium->tymed = TYMED_HGLOBAL;
        pMedium->hGlobal = hGlobal;
        pMedium->pUnkForRelease = nullptr;
        return S_OK;
    }

    HRESULT GetFileContents(LONG lindex, STGMEDIUM* pMedium) {
        ExpandDirectories();  // Ensure recursive expansion has happened
        if (lindex < 0 || lindex >= static_cast<LONG>(m_files.size())) {
            LogMessage(LogLevel::Error,
                L"[HttpVirtualFileDataObject] GetFileContents: lindex=%ld out of range (size=%zu)",
                lindex, m_files.size());
            return DV_E_LINDEX;
        }

        auto& file = m_files[static_cast<size_t>(lindex)];
        LogMessage(LogLevel::Info,
            L"[HttpVirtualFileDataObject] GetFileContents: lindex=%ld name='%ls' isDir=%d",
            lindex, file.name.c_str(), file.isDirectory ? 1 : 0);
        if (file.isDirectory) {
            // Return an empty stream — the shell creates the folder from the descriptor
            ComPtr<IStream> stream;
            HRESULT hr = CreateStreamOnHGlobal(nullptr, TRUE, &stream);
            if (FAILED(hr)) return hr;
            pMedium->tymed = TYMED_ISTREAM;
            pMedium->pstm = stream.Detach();
            pMedium->pUnkForRelease = nullptr;
            return S_OK;
        }

        // Build the remote path for this file
        std::vector<std::wstring> segments = m_parentSegments;
        if (!file.remotePath.empty()) {
            SplitAndAppend(file.remotePath, segments);
        } else {
            segments.push_back(file.name);
        }
        std::wstring remotePath = JoinSegments({}, segments);

        HttpConnectionOptions options;
        options.host = m_rootParts.host;
        options.port = m_rootParts.port;
        options.basePath = m_rootParts.basePath;
        options.useHttps = m_rootParts.useHttps;

        // Return a streaming IStream that reads from WinHTTP incrementally.
        // This lets the shell copy engine show its progress dialog as it reads.
        auto* stream = new (std::nothrow) HttpDownloadStream(options, remotePath);
        if (!stream) return E_OUTOFMEMORY;

        HRESULT hr = stream->Initialize();
        if (FAILED(hr)) {
            stream->Release();
            return hr;
        }

        pMedium->tymed = TYMED_ISTREAM;
        pMedium->pstm = stream;  // Already has refcount 1
        pMedium->pUnkForRelease = nullptr;
        return S_OK;
    }

    void ExpandDirectories() {
        std::call_once(m_expandOnce, [this]() {
            std::vector<FileEntry> expanded;
            for (auto& entry : m_files) {
                if (!entry.isDirectory) {
                    expanded.push_back(std::move(entry));
                } else {
                    ExpandDirectory(entry.name, entry.name, entry, expanded, 0);
                }
            }
            LogMessage(LogLevel::Info,
                L"[HttpVirtualFileDataObject] ExpandDirectories: %zu -> %zu entries",
                m_files.size(), expanded.size());
            m_files = std::move(expanded);
        });
    }

    void ExpandDirectory(const std::wstring& relativePrefix,
                         const std::wstring& remotePrefix,
                         const FileEntry& dirEntry,
                         std::vector<FileEntry>& out,
                         int depth) {
        if (depth >= 32) return;

        // Add the directory entry itself — shell creates the folder
        FileEntry dirOut;
        dirOut.name = relativePrefix;
        dirOut.remotePath = remotePrefix;
        dirOut.attributes = FILE_ATTRIBUTE_DIRECTORY;
        dirOut.isDirectory = true;
        dirOut.lastWriteTime = dirEntry.lastWriteTime;
        out.push_back(std::move(dirOut));

        // Build the remote URL path for listing
        std::vector<std::wstring> segments = m_parentSegments;
        SplitAndAppend(remotePrefix, segments);
        std::wstring listPath = JoinSegments({}, segments);
        if (!listPath.empty() && listPath.back() != L'/') {
            listPath.push_back(L'/');
        }

        HttpConnectionOptions options;
        options.host = m_rootParts.host;
        options.port = m_rootParts.port;
        options.basePath = m_rootParts.basePath;
        options.useHttps = m_rootParts.useHttps;

        HttpClient client;
        std::vector<DirectoryEntry> children;
        HRESULT hr = client.ListDirectory(options, listPath, &children);
        if (FAILED(hr)) {
            LogMessage(LogLevel::Error,
                L"[HttpVirtualFileDataObject] ExpandDirectory: ListDirectory failed hr=0x%08X path=%ls",
                hr, listPath.c_str());
            return;
        }
        LogMessage(LogLevel::Info,
            L"[HttpVirtualFileDataObject] ExpandDirectory: %zu children at %ls",
            children.size(), listPath.c_str());

        for (auto& child : children) {
            // Build backslash-separated name for shell descriptor
            std::wstring childRelative = relativePrefix + L"\\" + child.name;
            // Build forward-slash remote path for download
            std::wstring childRemote = remotePrefix + L"/" + child.name;

            if (childRelative.size() >= MAX_PATH) continue;

            if (child.isDirectory) {
                FileEntry childDir;
                childDir.name = childRelative;
                childDir.remotePath = childRemote;
                childDir.lastWriteTime = child.lastWriteTime;
                childDir.isDirectory = true;
                childDir.attributes = FILE_ATTRIBUTE_DIRECTORY;
                ExpandDirectory(childRelative, childRemote, childDir, out, depth + 1);
            } else {
                FileEntry childFile;
                childFile.name = childRelative;
                childFile.remotePath = childRemote;
                childFile.size = child.size;
                childFile.hasSize = true;  // Size is always known from directory listing
                childFile.attributes = FILE_ATTRIBUTE_NORMAL;
                childFile.lastWriteTime = child.lastWriteTime;
                childFile.isDirectory = false;
                out.push_back(std::move(childFile));
            }
        }
    }

    ~HttpVirtualFileDataObject() = default;

    std::atomic<ULONG> m_refCount{1};
    ComPtr<IDataObject> m_inner;
    HttpUrlParts m_rootParts;
    std::vector<std::wstring> m_parentSegments;
    std::vector<FileEntry> m_files;
    CLIPFORMAT m_cfDescriptor = 0;
    CLIPFORMAT m_cfContents = 0;
    std::once_flag m_expandOnce;
};

}  // namespace

class HttpShellFolder::ViewCallback : public IShellFolderViewCB {
public:
    explicit ViewCallback(HttpShellFolder* owner) : owner_(owner) {
        if (owner && owner->absolutePidl_) {
            notifyPidl_ = ClonePidl(owner->absolutePidl_.get());
        }
    }

    void Invalidate() {
        std::lock_guard<std::mutex> lock(mutex_);
        owner_ = nullptr;
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IShellFolderViewCB) {
            *ppv = static_cast<IShellFolderViewCB*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return ++refCount_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG count = --refCount_;
        if (count == 0) delete this;
        return count;
    }

    IFACEMETHODIMP MessageSFVCB(UINT uMsg, WPARAM wParam, LPARAM lParam) override {
        switch (uMsg) {
            case SFVM_MERGEMENU: {
                auto* qcm = reinterpret_cast<QCMINFO*>(lParam);
                if (qcm && qcm->hmenu) {
                    filterCmdId_ = qcm->idCmdFirst;
                    InsertMenuW(qcm->hmenu, static_cast<UINT>(-1), MF_BYPOSITION | MF_STRING,
                                filterCmdId_, L"Filter...");
                    ++qcm->idCmdFirst;
                }
                return S_OK;
            }
            case SFVM_INVOKECOMMAND: {
                if (static_cast<UINT>(wParam) == filterCmdId_) {
                    std::wstring filterText;
                    std::lock_guard<std::mutex> lock(mutex_);
                    if (owner_) {
                        if (ShowFilterDialog(nullptr, owner_->filterString_.c_str(), &filterText)) {
                            owner_->filterString_ = filterText;
                            if (owner_->absolutePidl_) {
                                SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST,
                                               owner_->absolutePidl_.get(), nullptr);
                            }
                        }
                    }
                    return S_OK;
                }
                return E_FAIL;
            }
            case SFVM_DBLCLK: {
                // Return S_FALSE so DefView invokes the default verb through
                // SHCreateDefaultContextMenu → IContextMenuCB → DFM_INVOKECOMMAND.
                return S_FALSE;
            }
            case SFVM_GETNOTIFY: {
                if (notifyPidl_) {
                    auto** ppidl = reinterpret_cast<PCIDLIST_ABSOLUTE*>(wParam);
                    auto* pEvents = reinterpret_cast<LONG*>(lParam);
                    if (ppidl) *ppidl = notifyPidl_.get();
                    if (pEvents) *pEvents = SHCNE_UPDATEDIR | SHCNE_CREATE | SHCNE_DELETE;
                }
                return S_OK;
            }
            default:
                return E_NOTIMPL;
        }
    }

private:
    std::mutex mutex_;
    HttpShellFolder* owner_ = nullptr;
    UniquePidl notifyPidl_;
    std::atomic<ULONG> refCount_{1};
    UINT filterCmdId_ = 0;
};

HttpShellFolder::HttpShellFolder() {
    ModuleAddRef();
    LogMessage(LogLevel::Info, L"[HttpShellFolder] Default constructor, this=%p", this);
}

HttpShellFolder::HttpShellFolder(const HttpUrlParts& root, const std::vector<std::wstring>& segments)
    : rootParts_(root), pathSegments_(segments) {
    ModuleAddRef();
    LogMessage(LogLevel::Info, L"[HttpShellFolder] Constructor host=%ls basePath=%ls segments=%zu https=%d port=%d",
               root.host.c_str(), root.basePath.c_str(), segments.size(),
               root.useHttps ? 1 : 0, static_cast<int>(root.port));
    EnsurePidl();
}

HttpShellFolder::~HttpShellFolder() {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] Destructor this=%p isRoot=%d host=%ls",
               this, isNamespaceRoot_ ? 1 : 0, rootParts_.host.c_str());
    if (viewCallback_) {
        static_cast<ViewCallback*>(viewCallback_.Get())->Invalidate();
    }
    ModuleRelease();
}

HRESULT HttpShellFolder::Create(const HttpUrlParts& root, const std::vector<std::wstring>& segments, REFIID riid,
                                void** ppv) {
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    HttpShellFolder* folder = new (std::nothrow) HttpShellFolder(root, segments);
    if (!folder) {
        return E_OUTOFMEMORY;
    }
    HRESULT hr = folder->QueryInterface(riid, ppv);
    folder->Release();
    return hr;
}

HRESULT HttpShellFolder::CreateWithParentPidl(const HttpUrlParts& root, const std::vector<std::wstring>& segments,
                                               PCIDLIST_ABSOLUTE parentPidl, PCUIDLIST_RELATIVE childPidl,
                                               REFIID riid, void** ppv) {
    {
        std::wstring segStr;
        for (const auto& s : segments) { if (!segStr.empty()) segStr += L" / "; segStr += s; }
        LogMessage(LogLevel::Info, L"[HttpShellFolder] CreateWithParentPidl: host=%ls basePath=%ls segments=[%ls] parentPidl=%p childPidl=%p",
                   root.host.c_str(), root.basePath.c_str(), segStr.c_str(), parentPidl, childPidl);
    }
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    HttpShellFolder* folder = new (std::nothrow) HttpShellFolder();
    if (!folder) {
        return E_OUTOFMEMORY;
    }
    folder->rootParts_ = root;
    folder->pathSegments_ = segments;

    // Build absolute PIDL by combining parent's absolute PIDL with the child relative PIDL
    if (parentPidl && childPidl) {
        folder->absolutePidl_ = UniquePidl(ILCombine(parentPidl, childPidl));
    } else if (parentPidl) {
        folder->absolutePidl_ = ClonePidl(parentPidl);
    }
    folder->initialized_ = folder->absolutePidl_ != nullptr;

    LogMessage(LogLevel::Info, L"[HttpShellFolder] CreateWithParentPidl: initialized=%d absolutePidl=%p",
               folder->initialized_ ? 1 : 0, folder->absolutePidl_.get());

    HRESULT hr = folder->QueryInterface(riid, ppv);
    folder->Release();
    LogMessage(LogLevel::Info, L"[HttpShellFolder] CreateWithParentPidl: QI result hr=0x%08X", hr);
    return hr;
}

IFACEMETHODIMP HttpShellFolder::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IShellFolder || riid == IID_IShellFolder2) {
        *object = static_cast<IShellFolder*>(this);
    } else if (riid == IID_IPersist || riid == IID_IPersistFolder || riid == IID_IPersistFolder2) {
        *object = static_cast<IPersistFolder2*>(this);
    } else {
        *object = nullptr;
        LogMessage(LogLevel::Verbose, L"[HttpShellFolder] QI FAILED for %ls (this=%p)",
                   GuidToString(riid).c_str(), this);
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) HttpShellFolder::AddRef() {
    return ++refCount_;
}

IFACEMETHODIMP_(ULONG) HttpShellFolder::Release() {
    ULONG count = --refCount_;
    if (count == 0) {
        delete this;
    }
    return count;
}

HRESULT HttpShellFolder::EnsurePidl() {
    if (initialized_ && absolutePidl_) {
        return S_OK;
    }
    if (rootParts_.host.empty()) {
        // This is the namespace root "Web Folders" — it has no URL of its own.
        LogMessage(LogLevel::Info, L"[HttpShellFolder] EnsurePidl: host is empty, setting namespace root");
        isNamespaceRoot_ = true;
        initialized_ = true;
        return S_OK;
    }
    absolutePidl_ = CreatePidlFromHttpUrl(rootParts_);
    if (!absolutePidl_) {
        return E_OUTOFMEMORY;
    }
    // Append path segments beyond basePath
    if (!pathSegments_.empty()) {
        PidlBuilder builder;
        // We need to rebuild: start from the root pidl, then append segments
        // Actually, just combine root pidl with segment pidls
        for (const auto& segment : pathSegments_) {
            PidlBuilder segBuilder;
            ComponentDefinition nameComp{ComponentType::Name, segment.c_str(), segment.size() * sizeof(wchar_t)};
            HRESULT hr = segBuilder.Append(ItemType::Directory, {nameComp});
            if (FAILED(hr)) {
                return hr;
            }
            UniquePidl segPidl = segBuilder.Finalize();
            if (!segPidl) {
                return E_OUTOFMEMORY;
            }
            UniquePidl combined(ILCombine(absolutePidl_.get(), segPidl.get()));
            if (!combined) {
                return E_OUTOFMEMORY;
            }
            absolutePidl_ = std::move(combined);
        }
    }
    initialized_ = true;
    return S_OK;
}

bool HttpShellFolder::ExtractRelativeSegments(PCUIDLIST_RELATIVE pidl, std::vector<std::wstring>* segments,
                                              bool* isDirectory) const {
    if (!segments) {
        return false;
    }
    segments->clear();
    if (isDirectory) {
        *isDirectory = true;
    }

    // Walk the relative PIDL and extract names from HTTP items
    const BYTE* cursor = reinterpret_cast<const BYTE*>(pidl);
    while (true) {
        const auto* item = reinterpret_cast<const SHITEMID*>(cursor);
        if (item->cb == 0) {
            break;
        }
        if (IsHttpItemId(*item)) {
            ItemType type = GetItemType(*item);
            if (type == ItemType::Root) {
                // Skip root items in relative navigation
            } else {
                std::wstring name;
                if (TryGetComponentString(*item, ComponentType::Name, &name)) {
                    segments->push_back(std::move(name));
                }
                if (isDirectory) {
                    *isDirectory = (type == ItemType::Directory);
                }
            }
        }
        cursor += item->cb;
    }

    return !segments->empty() || (pidl->mkid.cb == 0);
}

std::wstring HttpShellFolder::BuildRemotePath(const std::vector<std::wstring>& extra) const {
    return JoinSegments(pathSegments_, extra);
}

// Static members for download queue management
std::mutex HttpShellFolder::s_queueMutex;
std::unordered_map<std::wstring, std::unique_ptr<HttpDownloadQueue>> HttpShellFolder::s_downloadQueues;

HttpDownloadQueue* HttpShellFolder::GetDownloadQueue(const std::wstring& host,
                                                      int maxConcurrent, int speedLimitKBps) {
    std::lock_guard<std::mutex> lock(s_queueMutex);
    auto it = s_downloadQueues.find(host);
    if (it == s_downloadQueues.end()) {
        auto queue = std::make_unique<HttpDownloadQueue>();
        queue->Configure(maxConcurrent, speedLimitKBps);
        auto* ptr = queue.get();
        s_downloadQueues[host] = std::move(queue);
        return ptr;
    }
    it->second->Configure(maxConcurrent, speedLimitKBps);
    return it->second.get();
}

HRESULT HttpShellFolder::DownloadFileToStream(const std::vector<std::wstring>& segments,
                                              ComPtr<IStream>* stream) const {
    if (!stream) {
        return E_POINTER;
    }
    stream->Reset();
    if (segments.empty()) {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    wchar_t tempPath[MAX_PATH];
    DWORD written = GetTempPathW(static_cast<DWORD>(std::size(tempPath)), tempPath);
    if (written == 0 || written > std::size(tempPath)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    wchar_t tempFile[MAX_PATH];
    if (!GetTempFileNameW(tempPath, L"htp", 0, tempFile)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    HttpConnectionOptions options;
    options.host = rootParts_.host;
    options.port = rootParts_.port;
    options.basePath = rootParts_.basePath;
    options.useHttps = rootParts_.useHttps;

    std::wstring remotePath = JoinSegments(pathSegments_, segments);

    HttpClient client;
    HRESULT hr = client.DownloadFile(options, remotePath, tempFile);
    if (FAILED(hr)) {
        DeleteFileW(tempFile);
        return hr;
    }

    hr = SHCreateStreamOnFileEx(tempFile, STGM_READ | STGM_SHARE_DENY_NONE | STGM_DELETEONRELEASE,
                                FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, stream->GetAddressOf());
    if (FAILED(hr)) {
        DeleteFileW(tempFile);
        return hr;
    }
    return S_OK;
}

HRESULT HttpShellFolder::BindToChild(const std::vector<std::wstring>& segments, REFIID riid, void** ppv) const {
    return HttpShellFolder::Create(rootParts_, segments, riid, ppv);
}

HRESULT HttpShellFolder::EnumRootEntries(HWND, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: grfFlags=0x%08X", static_cast<DWORD>(grfFlags));
    // Only show folders at the root level
    if ((grfFlags & (SHCONTF_FOLDERS | kShcontfAllFolders)) == 0) {
        LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: no folder flags, returning empty");
        // Return empty enumerator
        auto enumerator = new (std::nothrow) StaticEnumIDList({});
        if (!enumerator) return E_OUTOFMEMORY;
        *ppenumIDList = enumerator;
        return S_OK;
    }

    auto opts = OptionsStore::Instance().Get();
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: %zu web folder entries in options",
               opts.webFolderEntries.size());
    std::vector<std::vector<std::uint8_t>> pidlBytes;

    for (const auto& entry : opts.webFolderEntries) {
        if (!entry.enabled) {
            LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: skipping disabled entry '%ls'",
                       entry.displayName.c_str());
            continue;
        }
        HttpUrlParts parts;
        if (!TryParseHttpUrl(entry.url, &parts)) {
            LogMessage(LogLevel::Warning, L"[HttpShellFolder] EnumRootEntries: failed to parse URL '%ls'",
                       entry.url.c_str());
            continue;
        }
        parts.displayName = entry.displayName;

        LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: building PIDL for '%ls' -> host=%ls basePath=%ls https=%d port=%d",
                   entry.displayName.c_str(), parts.host.c_str(), parts.basePath.c_str(),
                   parts.useHttps ? 1 : 0, static_cast<int>(parts.port));

        // Build a single-item PIDL (Root type) for this entry
        std::wstring scheme = parts.useHttps ? L"https" : L"http";
        std::uint16_t port = parts.port;

        PidlBuilder builder;
        ComponentDefinition hostComp{ComponentType::Host, parts.host.c_str(), parts.host.size() * sizeof(wchar_t)};
        ComponentDefinition portComp{ComponentType::Port, &port, sizeof(port)};
        ComponentDefinition basePathComp{ComponentType::BasePath, parts.basePath.c_str(),
                                         parts.basePath.size() * sizeof(wchar_t)};
        ComponentDefinition schemeComp{ComponentType::Scheme, scheme.c_str(), scheme.size() * sizeof(wchar_t)};
        ComponentDefinition nameComp{ComponentType::Name, parts.displayName.c_str(),
                                      parts.displayName.size() * sizeof(wchar_t)};

        HRESULT hr = builder.Append(ItemType::Root, {hostComp, portComp, basePathComp, schemeComp, nameComp});
        if (FAILED(hr)) {
            LogMessage(LogLevel::Error, L"[HttpShellFolder] EnumRootEntries: PidlBuilder::Append failed hr=0x%08X", hr);
            continue;
        }
        UniquePidl pidl = builder.Finalize();
        if (!pidl) {
            LogMessage(LogLevel::Error, L"[HttpShellFolder] EnumRootEntries: PidlBuilder::Finalize returned null");
            continue;
        }
        UINT size = ILGetSize(pidl.get());
        LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: built PIDL for '%ls', size=%u bytes",
                   entry.displayName.c_str(), size);
        std::vector<std::uint8_t> bytes(size);
        std::memcpy(bytes.data(), pidl.get(), size);
        pidlBytes.push_back(std::move(bytes));
    }

    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRootEntries: returning %zu entries", pidlBytes.size());
    auto enumerator = new (std::nothrow) StaticEnumIDList(std::move(pidlBytes));
    if (!enumerator) {
        return E_OUTOFMEMORY;
    }
    *ppenumIDList = enumerator;
    return S_OK;
}

HRESULT HttpShellFolder::EnumRemoteDirectory(HWND /*hwnd*/, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRemoteDirectory: host=%ls basePath=%ls segments=%zu",
               rootParts_.host.c_str(), rootParts_.basePath.c_str(), pathSegments_.size());
    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] EnumRemoteDirectory: EnsurePidl failed hr=0x%08X", hr);
        return hr;
    }

    // Fetch the directory listing synchronously so all items are available
    // before DefView starts iterating.  This prevents "This folder is empty"
    // from flashing while an async HTTP request is in progress.
    HttpConnectionOptions options;
    options.host = rootParts_.host;
    options.port = rootParts_.port;
    options.basePath = rootParts_.basePath;
    options.useHttps = rootParts_.useHttps;

    std::wstring remotePath = JoinSegments(pathSegments_, {});
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRemoteDirectory: ListDirectory remotePath=%ls",
               remotePath.c_str());

    HttpClient client;
    std::vector<DirectoryEntry> entries;
    hr = client.ListDirectory(options, remotePath, &entries);
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRemoteDirectory: ListDirectory returned hr=0x%08X entries=%zu",
               hr, entries.size());
    if (FAILED(hr)) {
        return hr;
    }

    // Build PIDLs for all qualifying entries
    std::vector<std::vector<std::uint8_t>> pidlList;
    pidlList.reserve(entries.size());
    int dirCount = 0, fileCount = 0, filteredOut = 0;
    for (const auto& entry : entries) {
        const bool isDirectory = entry.isDirectory;
        if (isDirectory) {
            ++dirCount;
            if ((grfFlags & (SHCONTF_FOLDERS | kShcontfAllFolders)) == 0) { ++filteredOut; continue; }
        } else {
            ++fileCount;
            if ((grfFlags & (SHCONTF_NONFOLDERS | SHCONTF_STORAGE)) == 0) { ++filteredOut; continue; }
        }
        if (!filterString_.empty() && !ContainsIgnoreCase(entry.name, filterString_)) {
            continue;
        }

        WIN32_FIND_DATAW findData{};
        findData.dwFileAttributes = isDirectory ? (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_READONLY)
                                                : (FILE_ATTRIBUTE_ARCHIVE | FILE_ATTRIBUTE_READONLY);
        findData.ftCreationTime = entry.lastWriteTime;
        findData.ftLastAccessTime = entry.lastWriteTime;
        findData.ftLastWriteTime = entry.lastWriteTime;
        findData.nFileSizeHigh = static_cast<DWORD>(entry.size >> 32);
        findData.nFileSizeLow = static_cast<DWORD>(entry.size & 0xFFFFFFFFULL);
        StringCchCopyW(findData.cFileName, ARRAYSIZE(findData.cFileName), entry.name.c_str());

        PidlBuilder builder;
        ComponentDefinition nameComponent{ComponentType::Name, entry.name.c_str(), entry.name.size() * sizeof(wchar_t)};
        ComponentDefinition dataComponent{ComponentType::FindData, &findData, sizeof(findData)};
        ItemType type = isDirectory ? ItemType::Directory : ItemType::File;
        hr = builder.Append(type, {nameComponent, dataComponent});
        if (FAILED(hr)) {
            LogMessage(LogLevel::Error, L"[HttpShellFolder] EnumRemoteDirectory: PidlBuilder::Append failed hr=0x%08X for '%ls'",
                       hr, entry.name.c_str());
            continue;
        }
        UniquePidl pidl = builder.Finalize();
        if (!pidl) continue;
        UINT size = ILGetSize(pidl.get());
        std::vector<std::uint8_t> bytes(size);
        std::memcpy(bytes.data(), pidl.get(), size);
        pidlList.push_back(std::move(bytes));
    }

    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumRemoteDirectory: returning %zu items (dirs=%d files=%d filtered=%d)",
               pidlList.size(), dirCount, fileCount, filteredOut);
    auto* enumerator = new (std::nothrow) StaticEnumIDList(std::move(pidlList));
    if (!enumerator) {
        return E_OUTOFMEMORY;
    }
    *ppenumIDList = enumerator;
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::ParseDisplayName(HWND, IBindCtx*, PWSTR pszName, ULONG* pchEaten,
                                                 PIDLIST_RELATIVE* ppidl, ULONG* pdwAttributes) {
    if (!ppidl) {
        return E_POINTER;
    }
    *ppidl = nullptr;
    if (!pszName) {
        return E_INVALIDARG;
    }

    std::wstring input(pszName);

    // Handle http:// or https:// URL input
    if (input.rfind(L"http://", 0) == 0 || input.rfind(L"https://", 0) == 0) {
        HttpUrlParts parts;
        if (!TryParseHttpUrl(input, &parts)) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (isNamespaceRoot_) {
            // At the root, parse the URL to create a Root PIDL
            std::wstring scheme = parts.useHttps ? L"https" : L"http";
            std::uint16_t port = parts.port;
            PidlBuilder builder;
            ComponentDefinition hostComp{ComponentType::Host, parts.host.c_str(), parts.host.size() * sizeof(wchar_t)};
            ComponentDefinition portComp{ComponentType::Port, &port, sizeof(port)};
            ComponentDefinition basePathComp{ComponentType::BasePath, parts.basePath.c_str(),
                                             parts.basePath.size() * sizeof(wchar_t)};
            ComponentDefinition schemeComp{ComponentType::Scheme, scheme.c_str(), scheme.size() * sizeof(wchar_t)};
            ComponentDefinition nameComp{ComponentType::Name, parts.host.c_str(), parts.host.size() * sizeof(wchar_t)};
            HRESULT hr = builder.Append(ItemType::Root, {hostComp, portComp, basePathComp, schemeComp, nameComp});
            if (FAILED(hr)) return hr;
            UniquePidl rel = builder.Finalize();
            if (!rel) return E_OUTOFMEMORY;
            *ppidl = reinterpret_cast<PIDLIST_RELATIVE>(rel.release());
            if (pchEaten) *pchEaten = static_cast<ULONG>(wcslen(pszName));
            if (pdwAttributes) *pdwAttributes = SFGAO_FOLDER | SFGAO_HASSUBFOLDER;
            return S_OK;
        }
        // Verify same host
        if (!EqualsIgnoreCase(parts.host, rootParts_.host) || parts.port != rootParts_.port) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
    }

    // Relative path parsing
    std::vector<std::wstring> segments;
    bool isDirectory = true;
    std::wstring buffer(input);
    if (!buffer.empty() && (buffer.back() == L'/' || buffer.back() == L'\\')) {
        isDirectory = true;
        buffer.pop_back();
    }
    std::wstring current;
    for (wchar_t ch : buffer) {
        if (ch == L'/' || ch == L'\\') {
            if (!current.empty()) {
                segments.push_back(current);
                current.clear();
            }
        } else {
            current.push_back(ch);
        }
    }
    if (!current.empty()) {
        segments.push_back(current);
    }
    if (!isDirectory && !segments.empty() && segments.back().find(L'.') != std::wstring::npos) {
        isDirectory = false;
    } else if (!segments.empty()) {
        isDirectory = true;
    }

    PidlBuilder builder;
    for (size_t i = 0; i < segments.size(); ++i) {
        ItemType type = (i + 1 == segments.size() && !isDirectory) ? ItemType::File : ItemType::Directory;
        const std::wstring& segment = segments[i];
        ComponentDefinition component{ComponentType::Name, segment.data(), segment.size() * sizeof(wchar_t)};
        HRESULT hr = builder.Append(type, {component});
        if (FAILED(hr)) {
            return hr;
        }
    }
    UniquePidl rel = builder.Finalize();
    if (!rel) {
        return E_OUTOFMEMORY;
    }
    *ppidl = reinterpret_cast<PIDLIST_RELATIVE>(rel.release());
    if (pchEaten) {
        *pchEaten = static_cast<ULONG>(wcslen(pszName));
    }
    if (pdwAttributes) {
        *pdwAttributes = isDirectory
                             ? (SFGAO_FOLDER | SFGAO_HASSUBFOLDER | SFGAO_BROWSABLE | SFGAO_STORAGE | SFGAO_READONLY | SFGAO_CANCOPY)
                             : (SFGAO_STREAM | SFGAO_STORAGE | SFGAO_READONLY | SFGAO_CANCOPY);
    }
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumObjects: isRoot=%d host=%ls grfFlags=0x%08X hwnd=%p",
               isNamespaceRoot_ ? 1 : 0, rootParts_.host.c_str(), static_cast<DWORD>(grfFlags), hwnd);
    if (!ppenumIDList) {
        return E_POINTER;
    }
    *ppenumIDList = nullptr;

    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] EnumObjects: EnsurePidl failed hr=0x%08X", hr);
        return hr;
    }

    if (isNamespaceRoot_) {
        LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumObjects: calling EnumRootEntries");
        hr = EnumRootEntries(hwnd, grfFlags, ppenumIDList);
        LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumObjects: EnumRootEntries returned hr=0x%08X enum=%p",
                   hr, ppenumIDList ? *ppenumIDList : nullptr);
        return hr;
    }
    // Update the navigation tracker so the site-root's GetCurFolder returns
    // a PIDL that includes our path segments (e.g. /eXo rather than just /).
    if (!rootParts_.host.empty() && absolutePidl_) {
        auto key = MakeNavTrackerKey(rootParts_);
        std::shared_ptr<ViewNavigationTracker> tracker;
        {
            std::lock_guard lock(s_navTrackerMutex);
            auto it = s_navTrackers.find(key);
            if (it != s_navTrackers.end()) {
                tracker = it->second;
            }
        }
        if (tracker) {
            std::lock_guard lock(tracker->mutex);
            tracker->absolutePidl = ClonePidl(absolutePidl_.get());
            tracker->pathSegments = pathSegments_;
            LogMessage(LogLevel::Info,
                       L"[HttpShellFolder] EnumObjects: updated nav tracker for %ls segments=%zu",
                       key.c_str(), pathSegments_.size());
        }
    }

    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumObjects: calling EnumRemoteDirectory for %ls%ls",
               rootParts_.host.c_str(), rootParts_.basePath.c_str());
    hr = EnumRemoteDirectory(hwnd, grfFlags, ppenumIDList);
    LogMessage(LogLevel::Info, L"[HttpShellFolder] EnumObjects: EnumRemoteDirectory returned hr=0x%08X", hr);
    return hr;
}

IFACEMETHODIMP HttpShellFolder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx*, REFIID riid, void** ppv) {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] BindToObject: isRoot=%d riid=%ls pidl=%p this=%p",
               isNamespaceRoot_ ? 1 : 0, GuidToString(riid).c_str(), pidl, this);
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;

    HRESULT hrPidl = EnsurePidl();
    if (FAILED(hrPidl)) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] BindToObject: EnsurePidl failed hr=0x%08X", hrPidl);
        return hrPidl;
    }

    if (isNamespaceRoot_) {
        // The child PIDL is a Root item — extract its URL parts and create a new folder
        if (!pidl || pidl->mkid.cb == 0) {
            LogMessage(LogLevel::Error, L"[HttpShellFolder] BindToObject: namespace root but pidl is null/empty");
            return E_INVALIDARG;
        }

        bool isHttpItem = IsHttpItemId(pidl->mkid);
        ItemType itemType = isHttpItem ? GetItemType(pidl->mkid) : ItemType::File;
        LogMessage(LogLevel::Info, L"[HttpShellFolder] BindToObject: namespace root child - isHttpItem=%d itemType=%d cb=%u",
                   isHttpItem ? 1 : 0, static_cast<int>(itemType), pidl->mkid.cb);

        if (isHttpItem && itemType == ItemType::Root) {
            HttpUrlParts parts;
            TryGetComponentString(pidl->mkid, ComponentType::Host, &parts.host);
            std::uint16_t port = 0;
            if (TryGetComponentUint16(pidl->mkid, ComponentType::Port, &port)) {
                parts.port = port;
            }
            TryGetComponentString(pidl->mkid, ComponentType::BasePath, &parts.basePath);
            std::wstring scheme;
            if (TryGetComponentString(pidl->mkid, ComponentType::Scheme, &scheme)) {
                parts.useHttps = (scheme == L"https");
            }
            TryGetComponentString(pidl->mkid, ComponentType::Name, &parts.displayName);

            LogMessage(LogLevel::Info, L"[HttpShellFolder] BindToObject: Root item -> host=%ls basePath=%ls scheme=%ls port=%d name=%ls",
                       parts.host.c_str(), parts.basePath.c_str(), scheme.c_str(), static_cast<int>(parts.port),
                       parts.displayName.c_str());

            // Check if there are further items after the Root
            std::vector<std::wstring> segments;
            const BYTE* cursor = reinterpret_cast<const BYTE*>(pidl);
            cursor += pidl->mkid.cb;
            int extraItemCount = 0;
            while (true) {
                const auto* item = reinterpret_cast<const SHITEMID*>(cursor);
                if (item->cb == 0) break;
                ++extraItemCount;
                bool isHttp = IsHttpItemId(*item);
                ItemType extraType = isHttp ? GetItemType(*item) : ItemType::File;
                LogMessage(LogLevel::Info,
                           L"[HttpShellFolder] BindToObject: extra item #%d cb=%u isHttp=%d type=%d",
                           extraItemCount, item->cb, isHttp ? 1 : 0, static_cast<int>(extraType));
                if (isHttp) {
                    std::wstring name;
                    if (TryGetComponentString(*item, ComponentType::Name, &name)) {
                        LogMessage(LogLevel::Info,
                                   L"[HttpShellFolder] BindToObject: extra segment '%ls'", name.c_str());
                        segments.push_back(std::move(name));
                    }
                }
                cursor += item->cb;
            }

            LogMessage(LogLevel::Info, L"[HttpShellFolder] BindToObject: creating child folder, extra segments=%zu", segments.size());
            HRESULT hr = CreateWithParentPidl(parts, segments, absolutePidl_.get(), pidl, riid, ppv);
            LogMessage(LogLevel::Info, L"[HttpShellFolder] BindToObject: CreateWithParentPidl returned hr=0x%08X ppv=%p",
                       hr, ppv ? *ppv : nullptr);
            return hr;
        }
        LogMessage(LogLevel::Warning, L"[HttpShellFolder] BindToObject: namespace root child is NOT a Root item, falling through");
    }

    std::vector<std::wstring> segments;
    bool directory = true;
    if (!ExtractRelativeSegments(pidl, &segments, &directory)) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] BindToObject: ExtractRelativeSegments failed");
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (!directory) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] BindToObject: target is not a directory");
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }
    std::vector<std::wstring> combined = pathSegments_;
    combined.insert(combined.end(), segments.begin(), segments.end());

    // Detailed diagnostics: log every segment so we can trace path accumulation
    {
        std::wstring parentSegs, childSegs, combinedSegs;
        for (const auto& s : pathSegments_) { if (!parentSegs.empty()) parentSegs += L" / "; parentSegs += s; }
        for (const auto& s : segments) { if (!childSegs.empty()) childSegs += L" / "; childSegs += s; }
        for (const auto& s : combined) { if (!combinedSegs.empty()) combinedSegs += L" / "; combinedSegs += s; }
        LogMessage(LogLevel::Info,
                   L"[HttpShellFolder] BindToObject NON-ROOT: this=%p basePath=%ls parentSegs=[%ls] childSegs=[%ls] combined=[%ls] total=%zu",
                   this, rootParts_.basePath.c_str(), parentSegs.c_str(), childSegs.c_str(), combinedSegs.c_str(), combined.size());
    }
    // Use parent PIDL combined with child relative PIDL for correct absolute path
    HRESULT hr = CreateWithParentPidl(rootParts_, combined, absolutePidl_.get(), pidl, riid, ppv);
    LogMessage(LogLevel::Info, L"[HttpShellFolder] BindToObject: result hr=0x%08X", hr);
    return hr;
}

IFACEMETHODIMP HttpShellFolder::BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx*, REFIID riid, void** ppv) {
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    if (riid != IID_IUnknown && riid != IID_IStream) {
        return E_NOINTERFACE;
    }
    std::vector<std::wstring> segments;
    bool directory = true;
    if (!ExtractRelativeSegments(pidl, &segments, &directory)) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (segments.empty() || directory) {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }
    ComPtr<IStream> stream;
    HRESULT hr = DownloadFileToStream(segments, &stream);
    if (FAILED(hr)) {
        return hr;
    }
    return stream->QueryInterface(riid, ppv);
}

IFACEMETHODIMP HttpShellFolder::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) {
    if (!pidl1 || !pidl2) {
        return E_POINTER;
    }
    int column = static_cast<int>(LOWORD(lParam));
    if (column < 0 || column >= static_cast<int>(kColumnCount)) {
        column = 0;
    }
    WIN32_FIND_DATAW left{};
    WIN32_FIND_DATAW right{};
    const bool hasLeft = TryGetFindData(pidl1, &left);
    const bool hasRight = TryGetFindData(pidl2, &right);
    int comparison = 0;
    switch (column) {
        case 1: {
            if (hasLeft && hasRight) {
                ULONGLONG sizeLeft = GetFileSizeFromFindData(left);
                ULONGLONG sizeRight = GetFileSizeFromFindData(right);
                if (sizeLeft < sizeRight) comparison = -1;
                else if (sizeLeft > sizeRight) comparison = 1;
            }
            break;
        }
        case 2: {
            if (hasLeft && hasRight) {
                comparison = static_cast<int>(CompareFileTime(&left.ftLastWriteTime, &right.ftLastWriteTime));
            }
            break;
        }
        default:
            break;
    }
    if (comparison == 0) {
        std::wstring nameLeft;
        std::wstring nameRight;
        if (hasLeft) nameLeft.assign(left.cFileName);
        else TryGetNameFromPidl(pidl1, &nameLeft);
        if (hasRight) nameRight.assign(right.cFileName);
        else TryGetNameFromPidl(pidl2, &nameRight);
        comparison = _wcsicmp(nameLeft.c_str(), nameRight.c_str());
    }
    if (comparison < 0) return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0xFFFF);
    if (comparison > 0) return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 1);
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::CreateViewObject(HWND hwnd, REFIID riid, void** ppv) {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] CreateViewObject: riid=%ls isRoot=%d host=%ls hwnd=%p",
               GuidToString(riid).c_str(), isNamespaceRoot_ ? 1 : 0, rootParts_.host.c_str(), hwnd);
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;

    if (riid == IID_IContextMenu || riid == IID_IContextMenu2 || riid == IID_IContextMenu3) {
        HRESULT hrPidl = EnsurePidl();
        if (FAILED(hrPidl)) return hrPidl;

        NamespaceMenuContext ctx;
        ctx.ownerWindow = hwnd;
        ctx.folderPidl = absolutePidl_.get();

        if (isNamespaceRoot_) {
            ctx.kind = NamespaceMenuKind::HttpBackground;
        } else {
            ctx.kind = NamespaceMenuKind::HttpRemoteBackground;
            // Build the full URL for "Open in Browser"
            std::wstring scheme = rootParts_.useHttps ? L"https" : L"http";
            ctx.httpUrl = scheme + L"://" + rootParts_.host;
            if ((rootParts_.useHttps && rootParts_.port != 443) ||
                (!rootParts_.useHttps && rootParts_.port != 80)) {
                ctx.httpUrl += L":" + std::to_wstring(rootParts_.port);
            }
            ctx.httpUrl += rootParts_.basePath;
            for (const auto& segment : pathSegments_) {
                ctx.httpUrl += L"/" + segment;
            }
            if (!ctx.httpUrl.empty() && ctx.httpUrl.back() != L'/') {
                ctx.httpUrl += L"/";
            }
        }

        auto* menu = new (std::nothrow) NamespaceContextMenu(ctx);
        if (!menu) return E_OUTOFMEMORY;
        HRESULT hr = menu->QueryInterface(riid, ppv);
        menu->Release();
        return hr;
    }

    if (riid != IID_IShellView) {
        LogMessage(LogLevel::Info, L"[HttpShellFolder] CreateViewObject: not IShellView, returning E_NOINTERFACE");
        return E_NOINTERFACE;
    }
    SFV_CREATE create{};
    create.cbSize = sizeof(create);
    FOLDERSETTINGS settings{};
    settings.ViewMode = FVM_DETAILS;
    settings.fFlags = FWF_SHOWSELALWAYS | FWF_AUTOARRANGE;
    AssignFolderSettings(create, settings);
    HRESULT hr = QueryInterface(IID_PPV_ARGS(&create.pshf));
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] CreateViewObject: QI for IShellFolder failed hr=0x%08X", hr);
        return hr;
    }
    if (!viewCallback_) {
        auto* cb = new (std::nothrow) ViewCallback(this);
        if (cb) {
            viewCallback_.Attach(cb);
        }
    }
    create.psfvcb = viewCallback_.Get();
    hr = SHCreateShellFolderView(&create, reinterpret_cast<IShellView**>(ppv));
    LogMessage(LogLevel::Info, L"[HttpShellFolder] CreateViewObject: SHCreateShellFolderView returned hr=0x%08X ppv=%p",
               hr, ppv ? *ppv : nullptr);
    if (create.pshf) {
        create.pshf->Release();
    }
    if (SUCCEEDED(hr) && ppv && *ppv) {
        auto* shellView = static_cast<IShellView*>(*ppv);

        ComPtr<IFolderView2> folderView2;
        if (SUCCEEDED(shellView->QueryInterface(IID_PPV_ARGS(&folderView2)))) {
            ComPtr<IFolderView> folderView;
            if (SUCCEEDED(folderView2.As(&folderView))) {
                folderView->SetCurrentViewMode(settings.ViewMode);
            }
            folderView2->SetCurrentFolderFlags(settings.fFlags, settings.fFlags);
        }

        // Register a navigation tracker for this site so that subfolder
        // navigations update the PIDL returned by GetCurFolder.
        if (!isNamespaceRoot_ && !rootParts_.host.empty() && absolutePidl_) {
            auto key = MakeNavTrackerKey(rootParts_);
            auto tracker = std::make_shared<ViewNavigationTracker>();
            {
                std::lock_guard lock(tracker->mutex);
                tracker->absolutePidl = ClonePidl(absolutePidl_.get());
                tracker->pathSegments = pathSegments_;
            }
            {
                std::lock_guard lock(s_navTrackerMutex);
                s_navTrackers[key] = tracker;
            }
            LogMessage(LogLevel::Info,
                       L"[HttpShellFolder] CreateViewObject: registered nav tracker for %ls",
                       key.c_str());
        }
    }
    return hr;
}

IFACEMETHODIMP HttpShellFolder::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, ULONG* rgfInOut) {
    if (!rgfInOut) {
        return E_POINTER;
    }
    ULONG mask = *rgfInOut;
    LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetAttributesOf: cidl=%u mask=0x%08X isRoot=%d",
               cidl, mask, isNamespaceRoot_ ? 1 : 0);
    if (cidl == 0) {
        ULONG folderFlags = SFGAO_FOLDER | SFGAO_STORAGE | SFGAO_HASSUBFOLDER |
                            SFGAO_READONLY | SFGAO_BROWSABLE | SFGAO_FILESYSANCESTOR | SFGAO_FILESYSTEM;
        *rgfInOut = mask == 0 ? folderFlags : (folderFlags & mask);
        LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetAttributesOf: cidl=0 (folder itself) -> 0x%08X", *rgfInOut);
        return S_OK;
    }
    if (!apidl) {
        return E_INVALIDARG;
    }
    ULONG relevantMask = mask == 0 ? 0xFFFFFFFFu : mask;
    ULONG result = relevantMask;
    for (UINT index = 0; index < cidl; ++index) {
        WIN32_FIND_DATAW findData{};
        ULONG itemFlags = 0;
        if (TryGetFindData(apidl[index], &findData)) {
            itemFlags = MapFindDataToAttributes(findData);
        } else {
            const ItemType type = (apidl[index] && apidl[index]->mkid.cb != 0) ? GetItemType(apidl[index]->mkid)
                                                                               : ItemType::File;
            if (type == ItemType::Directory || type == ItemType::Root) {
                itemFlags = SFGAO_FOLDER | SFGAO_STORAGE | SFGAO_HASSUBFOLDER |
                            SFGAO_READONLY | SFGAO_BROWSABLE | SFGAO_CANCOPY |
                            SFGAO_FILESYSANCESTOR | SFGAO_FILESYSTEM;
            } else {
                itemFlags = SFGAO_STREAM | SFGAO_STORAGE | SFGAO_READONLY | SFGAO_CANCOPY;
            }
        }
        if (index == 0) {
            LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetAttributesOf: item[0] flags=0x%08X", itemFlags);
        }
        result &= itemFlags;
    }
    if (mask != 0) {
        result &= mask;
    }
    *rgfInOut = result;
    LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetAttributesOf: result=0x%08X (BROWSABLE=%d FOLDER=%d)",
               result, (result & SFGAO_BROWSABLE) ? 1 : 0, (result & SFGAO_FOLDER) ? 1 : 0);
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::GetUIObjectOf(HWND hwnd, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
                                              UINT*, void** ppv) {
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    if (cidl > 0 && !apidl) {
        return E_INVALIDARG;
    }
    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        return hr;
    }

    LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetUIObjectOf: cidl=%u riid=%08X-%04X-%04X",
               cidl, riid.Data1, riid.Data2, riid.Data3);

    if (riid == IID_IExtractIconW) {
        if (cidl != 1 || !apidl || !apidl[0]) {
            return E_INVALIDARG;
        }
        // Determine if item is a directory/root or file
        bool isDirectory = false;
        WIN32_FIND_DATAW findData{};
        if (TryGetFindData(apidl[0], &findData)) {
            isDirectory = (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        } else {
            ItemType type = (apidl[0]->mkid.cb != 0) ? GetItemType(apidl[0]->mkid) : ItemType::File;
            isDirectory = (type == ItemType::Directory || type == ItemType::Root);
        }

        if (isDirectory) {
            // Return standard folder icon via shell associations
            ComPtr<IDefaultExtractIconInit> iconInit;
            hr = SHCreateDefaultExtractIcon(IID_PPV_ARGS(&iconInit));
            if (SUCCEEDED(hr)) {
                iconInit->SetNormalIcon(L"shell32.dll", 3);
                iconInit->SetOpenIcon(L"shell32.dll", 4);
                return iconInit->QueryInterface(riid, ppv);
            }
        } else {
            // For files, use the file extension to get the associated icon
            std::wstring name;
            if (TryGetFindData(apidl[0], &findData)) {
                name = findData.cFileName;
            } else {
                TryGetNameFromPidl(apidl[0], &name);
            }
            const wchar_t* ext = PathFindExtensionW(name.c_str());
            ComPtr<IDefaultExtractIconInit> iconInit;
            hr = SHCreateDefaultExtractIcon(IID_PPV_ARGS(&iconInit));
            if (SUCCEEDED(hr)) {
                // Use SHGetFileInfo to find the icon for this extension
                SHFILEINFOW sfi{};
                if (ext && *ext) {
                    SHGetFileInfoW(ext, FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi),
                                   SHGFI_ICON | SHGFI_USEFILEATTRIBUTES | SHGFI_SMALLICON);
                    if (sfi.hIcon) {
                        DestroyIcon(sfi.hIcon);
                    }
                    // Use the extension-based icon via AssocQueryString
                    wchar_t iconPath[MAX_PATH] = {};
                    DWORD iconPathSize = MAX_PATH;
                    if (SUCCEEDED(AssocQueryStringW(ASSOCF_INIT_DEFAULTTOSTAR, ASSOCSTR_DEFAULTICON,
                                                    ext, nullptr, iconPath, &iconPathSize))) {
                        // Parse "path,index" format
                        wchar_t* comma = wcsrchr(iconPath, L',');
                        int iconIndex = 0;
                        if (comma) {
                            *comma = 0;
                            iconIndex = _wtoi(comma + 1);
                        }
                        iconInit->SetNormalIcon(iconPath, iconIndex);
                    } else {
                        iconInit->SetNormalIcon(L"shell32.dll", 0);
                    }
                } else {
                    iconInit->SetNormalIcon(L"shell32.dll", 0);
                }
                return iconInit->QueryInterface(riid, ppv);
            }
        }
        return E_NOINTERFACE;
    }

    if (riid == IID_IDataObject) {
        if (!absolutePidl_) return E_FAIL;
        if (cidl == 0 || !apidl) {
            return SHCreateDataObject(absolutePidl_.get(), 0, nullptr, nullptr, riid, ppv);
        }

        // Build an inner data object for standard shell formats
        ComPtr<IDataObject> inner;
        hr = SHCreateDataObject(absolutePidl_.get(), cidl, apidl, nullptr, IID_PPV_ARGS(&inner));
        if (FAILED(hr)) return hr;

        // Extract file entries from the selected PIDLs for virtual file transfer
        std::vector<HttpVirtualFileDataObject::FileEntry> fileEntries;
        fileEntries.reserve(cidl);
        for (UINT i = 0; i < cidl; i++) {
            if (!apidl[i]) continue;
            HttpVirtualFileDataObject::FileEntry entry;
            WIN32_FIND_DATAW findData{};
            if (TryGetFindData(apidl[i], &findData)) {
                entry.name = findData.cFileName;
                entry.size = GetFileSizeFromFindData(findData);
                entry.attributes = findData.dwFileAttributes;
                entry.lastWriteTime = findData.ftLastWriteTime;
                entry.hasSize = true;
                entry.isDirectory = (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            } else {
                TryGetNameFromPidl(apidl[i], &entry.name);
                ItemType type = GetItemType(apidl[i]->mkid);
                entry.isDirectory = (type == ItemType::Directory || type == ItemType::Root);
                entry.attributes = entry.isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
            }
            if (!entry.name.empty()) {
                fileEntries.push_back(std::move(entry));
            }
        }

        // When Explorer reuses the site-root folder for subfolder views,
        // pathSegments_ is empty.  Read the nav tracker for the actual segments.
        std::vector<std::wstring> effectiveSegments = pathSegments_;
        if (effectiveSegments.empty() && !isNamespaceRoot_ && !rootParts_.host.empty()) {
            GetNavTrackerState(rootParts_, &effectiveSegments, nullptr);
        }
        LogMessage(LogLevel::Info,
                   L"[HttpShellFolder] GetUIObjectOf(IDataObject): segments=%zu effective=%zu",
                   pathSegments_.size(), effectiveSegments.size());

        auto* obj = new (std::nothrow) HttpVirtualFileDataObject(
            std::move(inner), rootParts_, std::move(effectiveSegments), std::move(fileEntries));
        if (!obj) return E_OUTOFMEMORY;
        hr = obj->QueryInterface(riid, ppv);
        obj->Release();
        return hr;
    }

    if (riid == IID_IContextMenu || riid == IID_IContextMenu2 || riid == IID_IContextMenu3) {
        // For namespace root items, provide our custom context menu
        if (isNamespaceRoot_ && cidl == 1 && apidl && apidl[0] && apidl[0]->mkid.cb != 0 &&
            IsHttpItemId(apidl[0]->mkid) && GetItemType(apidl[0]->mkid) == ItemType::Root) {
            NamespaceMenuContext ctx;
            ctx.kind = NamespaceMenuKind::HttpRootItem;
            ctx.ownerWindow = hwnd;
            ctx.folderPidl = absolutePidl_.get();

            // Find the matching entry in OptionsStore
            std::wstring host;
            TryGetComponentString(apidl[0]->mkid, ComponentType::Host, &host);
            std::uint16_t port = 80;
            TryGetComponentUint16(apidl[0]->mkid, ComponentType::Port, &port);

            auto options = OptionsStore::Instance().Get();
            for (int i = 0; i < static_cast<int>(options.webFolderEntries.size()); ++i) {
                // Match by display name or URL containing host
                if (options.webFolderEntries[i].url.find(host) != std::wstring::npos) {
                    ctx.entryIndex = i;
                    ctx.itemEnabled = options.webFolderEntries[i].enabled;
                    ctx.itemName = options.webFolderEntries[i].displayName;
                    break;
                }
            }

            // Build item's absolute PIDL for in-place navigation
            UniquePidl itemPidl(ILCombine(absolutePidl_.get(), apidl[0]));
            if (itemPidl) {
                ctx.itemAbsolutePidl = itemPidl.get();
            }

            auto* menu = new (std::nothrow) NamespaceContextMenu(ctx);
            if (!menu) return E_OUTOFMEMORY;
            hr = menu->QueryInterface(riid, ppv);
            menu->Release();
            return hr;
        }

        // For non-root items, provide association keys so the default context menu
        // knows the correct verbs (e.g., "open" for folders, "copy" for files)
        HKEY keys[2] = {};
        UINT keyCount = 0;
        bool isFolder = false;
        if (cidl >= 1 && apidl && apidl[0]) {
            WIN32_FIND_DATAW findData{};
            if (TryGetFindData(apidl[0], &findData)) {
                isFolder = (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            } else if (apidl[0]->mkid.cb != 0) {
                ItemType type = GetItemType(apidl[0]->mkid);
                isFolder = (type == ItemType::Directory);
            }
        }
        if (isFolder) {
            RegOpenKeyExW(HKEY_CLASSES_ROOT, L"Folder", 0, KEY_READ, &keys[keyCount]);
            if (keys[keyCount]) ++keyCount;
        } else {
            // For files, provide the "*" (all files) association key
            RegOpenKeyExW(HKEY_CLASSES_ROOT, L"*", 0, KEY_READ, &keys[keyCount]);
            if (keys[keyCount]) ++keyCount;
        }

        LogMessage(LogLevel::Info, L"[HttpShellFolder] GetUIObjectOf: IContextMenu cidl=%u isFolder=%d keyCount=%u",
                   cidl, isFolder ? 1 : 0, keyCount);

        // For single file items, use IContextMenuCB so that double-click / Enter
        // routes through DFM_INVOKECOMMAND → DownloadAndOpen.
        ComPtr<IContextMenuCB> callback;
        if (!isFolder && cidl == 1 && apidl && apidl[0]) {
            std::wstring fileName;
            WIN32_FIND_DATAW fd{};
            if (TryGetFindData(apidl[0], &fd)) {
                fileName = fd.cFileName;
            } else {
                TryGetNameFromPidl(apidl[0], &fileName);
            }

            std::vector<std::wstring> segments = pathSegments_;
            if (segments.empty() && !isNamespaceRoot_ && !rootParts_.host.empty()) {
                GetNavTrackerState(rootParts_, &segments, nullptr);
            }
            segments.push_back(fileName);

            callback.Attach(new (std::nothrow) HttpFileContextMenuCallback(
                rootParts_, std::move(segments), std::move(fileName)));
        }

        DEFCONTEXTMENU def{};
        def.hwnd = hwnd;
        def.pidlFolder = absolutePidl_.get();
        def.psf = this;
        def.cidl = cidl;
        def.apidl = apidl;
        if (keyCount > 0) {
            def.aKeys = keys;
            def.cKeys = keyCount;
        }
        if (callback) {
            def.pcmcb = callback.Get();
        }

        hr = SHCreateDefaultContextMenu(&def, riid, ppv);
        LogMessage(LogLevel::Info, L"[HttpShellFolder] GetUIObjectOf: SHCreateDefaultContextMenu hr=0x%08X pcmcb=%p",
                   hr, callback.Get());
        for (UINT ki = 0; ki < keyCount; ++ki) {
            if (keys[ki]) RegCloseKey(keys[ki]);
        }
        return hr;
    }

    if (riid == IID_IQueryAssociations) {
        ComPtr<IQueryAssociations> associations;
        hr = AssocCreate(CLSID_QueryAssociations, IID_PPV_ARGS(&associations));
        if (FAILED(hr)) {
            return hr;
        }

        ASSOCF flags = ASSOCF_INIT_DEFAULTTOSTAR | ASSOCF_INIT_IGNOREUNKNOWN;
        const wchar_t* assoc = L"*";
        bool isDirectory = false;
        if (cidl == 0 || !apidl) {
            flags = ASSOCF_INIT_DEFAULTTOFOLDER;
            assoc = nullptr;
            isDirectory = true;
        } else {
            std::vector<std::wstring> segments;
            if (ExtractRelativeSegments(apidl[0], &segments, &isDirectory) && !segments.empty() && !isDirectory) {
                std::wstring name;
                WIN32_FIND_DATAW findData{};
                if (TryGetFindData(apidl[0], &findData)) {
                    name.assign(findData.cFileName);
                } else {
                    TryGetNameFromPidl(apidl[0], &name);
                }
                if (!name.empty()) {
                    const wchar_t* extension = PathFindExtensionW(name.c_str());
                    if (extension && *extension) {
                        assoc = extension;
                    }
                }
            }
            if (isDirectory) {
                flags = ASSOCF_INIT_DEFAULTTOFOLDER;
                assoc = L"Folder";
            }
        }

        hr = associations->Init(flags, assoc, nullptr, hwnd);
        if (FAILED(hr)) {
            hr = associations->Init(ASSOCF_INIT_DEFAULTTOSTAR | ASSOCF_INIT_IGNOREUNKNOWN, L"*", nullptr, hwnd);
            if (FAILED(hr)) {
                return hr;
            }
        }
        *ppv = associations.Detach();
        return S_OK;
    }

    return E_NOINTERFACE;
}

IFACEMETHODIMP HttpShellFolder::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET* pName) {
    if (!pidl || !pName) {
        return E_POINTER;
    }

    LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetDisplayNameOf: uFlags=0x%08X isRoot=%d pidl.cb=%u host=%ls segments=%zu",
               uFlags, isNamespaceRoot_ ? 1 : 0, pidl ? pidl->mkid.cb : 0,
               rootParts_.host.c_str(), pathSegments_.size());

    if (uFlags & SHGDN_FORPARSING) {
        if (uFlags & SHGDN_INFOLDER) {
            WIN32_FIND_DATAW findData{};
            if (TryGetFindData(pidl, &findData)) {
                LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetDisplayNameOf: FORPARSING|INFOLDER -> '%ls'",
                           findData.cFileName);
                return AssignToStrRet(findData.cFileName, pName);
            }
            std::wstring name;
            if (TryGetNameFromPidl(pidl, &name)) {
                LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetDisplayNameOf: FORPARSING|INFOLDER -> '%ls'",
                           name.c_str());
                return AssignToStrRet(name, pName);
            }
            return E_FAIL;
        }
        // Full parsing name
        if (isNamespaceRoot_) {
            // For root items, return the display name as parsing name.
            // Returning a URL here causes Explorer to open the browser when navigating in.
            if (IsHttpItemId(pidl->mkid) && GetItemType(pidl->mkid) == ItemType::Root) {
                std::wstring name;
                TryGetComponentString(pidl->mkid, ComponentType::Name, &name);
                if (name.empty()) {
                    TryGetComponentString(pidl->mkid, ComponentType::Host, &name);
                }
                LogMessage(LogLevel::Verbose, L"[HttpShellFolder] GetDisplayNameOf: Root FORPARSING -> '%ls'",
                           name.c_str());
                return AssignToStrRet(name, pName);
            }
        }
        HRESULT hr = EnsurePidl();
        if (FAILED(hr)) {
            return hr;
        }
        if (absolutePidl_) {
            UniquePidl combined(ILCombine(absolutePidl_.get(), pidl));
            if (combined) {
                std::wstring url = BuildUrlFromHttpPidl(combined.get());
                if (!url.empty()) {
                    LogMessage(LogLevel::Info,
                               L"[HttpShellFolder] GetDisplayNameOf FORPARSING: this=%p url=%ls segments=%zu",
                               this, url.c_str(), pathSegments_.size());
                    return AssignToStrRet(url, pName);
                }
            }
        }
        // Fallback to name
        std::wstring name;
        if (TryGetNameFromPidl(pidl, &name)) {
            return AssignToStrRet(name, pName);
        }
        return E_FAIL;
    }

    // Display name
    WIN32_FIND_DATAW findData{};
    if (TryGetFindData(pidl, &findData)) {
        return AssignToStrRet(findData.cFileName, pName);
    }
    std::wstring name;
    if (!TryGetNameFromPidl(pidl, &name)) {
        return E_FAIL;
    }
    return AssignToStrRet(name, pName);
}

IFACEMETHODIMP HttpShellFolder::SetNameOf(HWND, PCUITEMID_CHILD, PCWSTR, SHGDNF, PIDLIST_RELATIVE*) {
    return E_NOTIMPL;  // Read-only
}

IFACEMETHODIMP HttpShellFolder::GetDefaultSearchGUID(GUID* pguid) {
    if (!pguid) {
        return E_POINTER;
    }
    *pguid = GUID_NULL;
    return E_NOTIMPL;
}

IFACEMETHODIMP HttpShellFolder::EnumSearches(IEnumExtraSearch** ppEnum) {
    if (!ppEnum) {
        return E_POINTER;
    }
    *ppEnum = nullptr;
    return E_NOTIMPL;
}

IFACEMETHODIMP HttpShellFolder::GetDefaultColumn(DWORD, ULONG* pSort, ULONG* pDisplay) {
    if (!pSort || !pDisplay) {
        return E_POINTER;
    }
    *pSort = 0;
    *pDisplay = 0;
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::GetDefaultColumnState(UINT iColumn, SHCOLSTATEF* pcsFlags) {
    if (!pcsFlags) {
        return E_POINTER;
    }
    if (iColumn >= kColumnCount) {
        return E_INVALIDARG;
    }
    *pcsFlags = kColumnDefinitions[iColumn].state;
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID* pscid, VARIANT* pv) {
    if (!pscid || !pv) {
        return E_POINTER;
    }
    VariantInit(pv);
    if (!pidl) {
        return E_INVALIDARG;
    }
    const PROPERTYKEY key{pscid->fmtid, pscid->pid};
    WIN32_FIND_DATAW findData{};
    if (!TryGetFindData(pidl, &findData)) {
        // For Root items at namespace root, show name
        if (IsEqualPropertyKey(key, PKEY_ItemNameDisplay)) {
            std::wstring name;
            if (TryGetNameFromPidl(pidl, &name)) {
                return InitVariantFromString(name.c_str(), pv);
            }
        }
        return S_FALSE;
    }
    if (IsEqualPropertyKey(key, PKEY_ItemNameDisplay)) {
        return InitVariantFromString(findData.cFileName, pv);
    }
    if (IsEqualPropertyKey(key, PKEY_Size)) {
        ULONGLONG size = GetFileSizeFromFindData(findData);
        return InitVariantFromUInt64(size, pv);
    }
    if (IsEqualPropertyKey(key, PKEY_DateModified)) {
        return InitVariantFromFileTime(&findData.ftLastWriteTime, pv);
    }
    return S_FALSE;
}

IFACEMETHODIMP HttpShellFolder::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS* pDetails) {
    if (!pDetails) {
        return E_POINTER;
    }
    if (iColumn >= kColumnCount) {
        return E_FAIL;
    }
    pDetails->fmt = kColumnDefinitions[iColumn].format;
    pDetails->cxChar = kColumnDefinitions[iColumn].width;
    if (!pidl) {
        return AssignToStrRet(kColumnDefinitions[iColumn].name, &pDetails->str);
    }
    WIN32_FIND_DATAW findData{};
    std::wstring value;
    if (TryGetFindData(pidl, &findData)) {
        switch (iColumn) {
            case 0:
                value.assign(findData.cFileName);
                break;
            case 1:
                if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
                    value = FormatSizeString(GetFileSizeFromFindData(findData));
                }
                break;
            case 2:
                value = FormatDateString(findData.ftLastWriteTime);
                break;
        }
    } else if (iColumn == 0) {
        TryGetNameFromPidl(pidl, &value);
    }
    return AssignToStrRet(value, &pDetails->str);
}

IFACEMETHODIMP HttpShellFolder::MapColumnToSCID(UINT iColumn, SHCOLUMNID* pscid) {
    if (!pscid) {
        return E_POINTER;
    }
    if (iColumn >= kColumnCount) {
        return E_INVALIDARG;
    }
    pscid->fmtid = kColumnDefinitions[iColumn].key.fmtid;
    pscid->pid = kColumnDefinitions[iColumn].key.pid;
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::GetClassID(CLSID* pClassID) {
    if (!pClassID) {
        return E_POINTER;
    }
    *pClassID = isNamespaceRoot_ ? CLSID_ShellTabsHttpRoot : CLSID_ShellTabsHttpFolder;
    LogMessage(LogLevel::Info, L"[HttpShellFolder] GetClassID: isRoot=%d -> %ls",
               isNamespaceRoot_ ? 1 : 0, GuidToString(*pClassID).c_str());
    return S_OK;
}

IFACEMETHODIMP HttpShellFolder::Initialize(PCIDLIST_ABSOLUTE pidl) {
    LogMessage(LogLevel::Info, L"[HttpShellFolder] Initialize called, pidl=%p, this=%p", pidl, this);
    if (!pidl) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] Initialize: pidl is null");
        return E_INVALIDARG;
    }

    // Log PIDL structure
    UINT pidlSize = ILGetSize(pidl);
    UINT pidlDepth = 0;
    {
        const BYTE* cursor = reinterpret_cast<const BYTE*>(pidl);
        while (true) {
            const auto* item = reinterpret_cast<const SHITEMID*>(cursor);
            if (item->cb == 0) break;
            ++pidlDepth;
            cursor += item->cb;
        }
    }
    LogMessage(LogLevel::Info, L"[HttpShellFolder] Initialize: PIDL size=%u depth=%u", pidlSize, pidlDepth);

    HttpUrlParts parts;
    std::vector<std::wstring> segments;
    bool directory = true;
    if (!TryParseHttpPidl(pidl, &parts, &segments, &directory)) {
        // This might be the namespace root being initialized with just a namespace prefix.
        // In that case, we are the root "Web Folders" node.
        LogMessage(LogLevel::Info, L"[HttpShellFolder] Initialize: TryParseHttpPidl failed -> namespace root");
        isNamespaceRoot_ = true;
        absolutePidl_ = ClonePidl(pidl);
        initialized_ = absolutePidl_ != nullptr;
        LogMessage(LogLevel::Info, L"[HttpShellFolder] Initialize: namespace root, initialized=%d", initialized_ ? 1 : 0);
        return initialized_ ? S_OK : E_OUTOFMEMORY;
    }
    rootParts_ = parts;
    rootParts_.basePath = parts.basePath;
    pathSegments_ = segments;
    if (!directory && !pathSegments_.empty()) {
        pathSegments_.pop_back();
    }
    absolutePidl_ = ClonePidl(pidl);
    initialized_ = absolutePidl_ != nullptr;
    isNamespaceRoot_ = false;
    {
        std::wstring segStr;
        for (const auto& s : pathSegments_) { if (!segStr.empty()) segStr += L" / "; segStr += s; }
        LogMessage(LogLevel::Info, L"[HttpShellFolder] Initialize: this=%p host=%ls basePath=%ls segments=[%ls] dir=%d initialized=%d",
                   this, rootParts_.host.c_str(), rootParts_.basePath.c_str(), segStr.c_str(),
                   directory ? 1 : 0, initialized_ ? 1 : 0);
    }
    return initialized_ ? S_OK : E_OUTOFMEMORY;
}

IFACEMETHODIMP HttpShellFolder::GetCurFolder(PIDLIST_ABSOLUTE* ppidl) {
    if (!ppidl) {
        return E_POINTER;
    }
    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpShellFolder] GetCurFolder: EnsurePidl failed hr=0x%08X", hr);
        return hr;
    }
    if (!absolutePidl_) {
        LogMessage(LogLevel::Warning, L"[HttpShellFolder] GetCurFolder: absolutePidl_ is null, returning S_FALSE");
        *ppidl = nullptr;
        return S_FALSE;
    }

    // For site-root folders (pathSegments_ empty, e.g. the Myrient root),
    // Explorer reuses our DefView for subfolder navigation and polls
    // GetCurFolder to build absolute PIDLs for child navigation.  Use the
    // navigation tracker to return the actual current-location PIDL rather
    // than our stale root-level PIDL.
    if (!isNamespaceRoot_ && !rootParts_.host.empty() && pathSegments_.empty()) {
        auto key = MakeNavTrackerKey(rootParts_);
        std::shared_ptr<ViewNavigationTracker> tracker;
        {
            std::lock_guard lock(s_navTrackerMutex);
            auto it = s_navTrackers.find(key);
            if (it != s_navTrackers.end()) {
                tracker = it->second;
            }
        }
        if (tracker) {
            std::lock_guard lock(tracker->mutex);
            if (tracker->absolutePidl) {
                *ppidl = ClonePidl(tracker->absolutePidl.get()).release();
                LogMessage(LogLevel::Info,
                           L"[HttpShellFolder] GetCurFolder: using nav tracker, segments=%zu pidl=%p",
                           tracker->pathSegments.size(), *ppidl);
                return *ppidl ? S_OK : E_OUTOFMEMORY;
            }
        }
    }

    *ppidl = ClonePidl(absolutePidl_.get()).release();
    HRESULT result = *ppidl ? S_OK : E_OUTOFMEMORY;
    LogMessage(LogLevel::Info, L"[HttpShellFolder] GetCurFolder: returning own pidl=%p segments=%zu hr=0x%08X",
               *ppidl, pathSegments_.size(), result);
    return result;
}

}  // namespace shelltabs::http
