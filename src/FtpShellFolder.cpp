#include "FtpShellFolder.h"

#include <shlobj.h>
#include <shlwapi.h>

#include <algorithm>
#include <atomic>
#include <condition_variable>
#include <cstdint>
#include <cwchar>
#include <cwctype>
#include <iterator>
#include <memory>
#include <mutex>
#include <string_view>
#include <thread>
#include <vector>

#include <commctrl.h>
#include <propvarutil.h>
#include <strsafe.h>
#include <winnls.h>
#include <wininet.h>

#ifndef INTERNET_OPTION_PASSIVE
#define INTERNET_OPTION_PASSIVE 52
#endif

#ifdef _MSC_VER
#pragma comment(lib, "propsys.lib")
#endif

using Microsoft::WRL::ComPtr;

namespace shelltabs::ftp {

namespace {

#ifdef SHCONTF_ALLFOLDERS
constexpr SHCONTF kShcontfAllFolders = SHCONTF_ALLFOLDERS;
#else
constexpr SHCONTF kShcontfAllFolders = static_cast<SHCONTF>(0x00000080);
#endif

constexpr GUID kFtpSearchProviderGuid =
    {0x9a3df3a4, 0x8d1a, 0x4a26, {0x9f, 0x48, 0xf8, 0x43, 0x61, 0x5b, 0xd9, 0x5e}};

class FtpSearchEnumerator : public IEnumExtraSearch {
public:
    explicit FtpSearchEnumerator(std::vector<EXTRASEARCH> entries)
        : refCount_(1), entries_(std::move(entries)) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IEnumExtraSearch) {
            *object = static_cast<IEnumExtraSearch*>(this);
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

    IFACEMETHODIMP Next(ULONG celt, EXTRASEARCH* rgelt, ULONG* pceltFetched) override {
        if (!rgelt) {
            return E_POINTER;
        }
        if (celt > 1 && !pceltFetched) {
            return E_INVALIDARG;
        }
        ULONG fetched = 0;
        while (fetched < celt && currentIndex_ < entries_.size()) {
            rgelt[fetched] = entries_[currentIndex_++];
            ++fetched;
        }
        if (pceltFetched) {
            *pceltFetched = fetched;
        }
        return fetched == celt ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG celt) override {
        size_t remaining = entries_.size() - currentIndex_;
        if (celt > remaining) {
            currentIndex_ = entries_.size();
            return S_FALSE;
        }
        currentIndex_ += celt;
        return S_OK;
    }

    IFACEMETHODIMP Reset() override {
        currentIndex_ = 0;
        return S_OK;
    }

    IFACEMETHODIMP Clone(IEnumExtraSearch** clone) override {
        if (!clone) {
            return E_POINTER;
        }
        auto copy = new (std::nothrow) FtpSearchEnumerator(entries_);
        if (!copy) {
            return E_OUTOFMEMORY;
        }
        copy->currentIndex_ = currentIndex_;
        *clone = copy;
        return S_OK;
    }

private:
    std::atomic<ULONG> refCount_;
    std::vector<EXTRASEARCH> entries_;
    size_t currentIndex_ = 0;
};

struct InternetHandle {
    InternetHandle() = default;
    explicit InternetHandle(HINTERNET value) : handle(value) {}
    ~InternetHandle() {
        if (handle) {
            InternetCloseHandle(handle);
        }
    }

    InternetHandle(const InternetHandle&) = delete;
    InternetHandle& operator=(const InternetHandle&) = delete;

    InternetHandle(InternetHandle&& other) noexcept : handle(other.handle) { other.handle = nullptr; }
    InternetHandle& operator=(InternetHandle&& other) noexcept {
        if (this != &other) {
            if (handle) {
                InternetCloseHandle(handle);
            }
            handle = other.handle;
            other.handle = nullptr;
        }
        return *this;
    }

    HINTERNET get() const noexcept { return handle; }
    HINTERNET release() noexcept {
        HINTERNET value = handle;
        handle = nullptr;
        return value;
    }
    void reset(HINTERNET value = nullptr) {
        if (handle) {
            InternetCloseHandle(handle);
        }
        handle = value;
    }

private:
    HINTERNET handle = nullptr;
};

HRESULT RenameRemoteItem(const FtpConnectionOptions& options, const FtpCredential& credential,
                         const std::wstring& directory, const std::wstring& oldName,
                         const std::wstring& newName) {
    InternetHandle internet(InternetOpenW(L"ShellTabs", INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0));
    if (!internet.get()) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    InternetHandle connection(InternetConnectW(internet.get(), options.host.c_str(), options.port,
                                               credential.userName.c_str(), credential.password.c_str(),
                                               INTERNET_SERVICE_FTP, 0, 0));
    if (!connection.get()) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    DWORD passive = options.passiveMode ? 1u : 0u;
    InternetSetOptionW(connection.get(), INTERNET_OPTION_PASSIVE, &passive, sizeof(passive));
    if (!directory.empty() && directory != L"/") {
        if (!FtpSetCurrentDirectoryW(connection.get(), directory.c_str())) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }
    if (!FtpRenameFileW(connection.get(), oldName.c_str(), newName.c_str())) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}

template <typename>
inline constexpr bool kDependentFalse = false;

template <typename SFVCreate>
void AssignFolderSettings(SFVCreate& create, const FOLDERSETTINGS& settings) {
    // Windows 10/11 SDKs expose the FOLDERSETTINGS pointer as pViewSettings,
    // while older headers still publish pfs or pfolderSettings. Detect the
    // available alias at compile time so the view gets initialized correctly.
    if constexpr (requires(SFVCreate& candidate) { candidate.pfs = &settings; }) {
        create.pfs = &settings;
    } else if constexpr (requires(SFVCreate& candidate) { candidate.pfolderSettings = &settings; }) {
        create.pfolderSettings = &settings;
    } else if constexpr (requires(SFVCreate& candidate) { candidate.pViewSettings = &settings; }) {
        create.pViewSettings = &settings;
    } else {
        static_assert(kDependentFalse<SFVCreate>, "SFV_CREATE is missing a folder settings member");
    }
}

std::wstring BuildCanonicalUrl(const FtpUrlParts& parts) {
    std::wstring url = L"ftp://";
    if (!parts.userName.empty()) {
        url += parts.userName;
        if (!parts.password.empty()) {
            url += L":" + parts.password;
        }
        url += L"@";
    }
    url += parts.host;
    if (parts.port != INTERNET_DEFAULT_FTP_PORT && parts.port != 0) {
        url += L":" + std::to_wstring(parts.port);
    }
    std::wstring path = parts.path;
    if (path.empty()) {
        path = L"/";
    }
    if (!path.empty() && path.front() != L'/') {
        url.push_back(L'/');
    }
    url += path.front() == L'/' ? path.substr(1) : path;
    if (!url.empty() && url.back() != L'/') {
        url.push_back(L'/');
    }
    return url;
}

std::vector<std::wstring> SplitSegments(std::wstring_view path) {
    std::vector<std::wstring> segments;
    std::wstring current;
    for (wchar_t ch : path) {
        if (ch == L'/') {
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
    return segments;
}

FtpUrlParts CombineParts(const FtpUrlParts& root, const std::vector<std::wstring>& segments) {
    FtpUrlParts combined = root;
    std::vector<std::wstring> base;
    if (!root.path.empty() && root.path != L"/") {
        std::wstring normalized = root.path;
        if (!normalized.empty() && normalized.front() == L'/') {
            normalized.erase(normalized.begin());
        }
        if (!normalized.empty() && normalized.back() == L'/') {
            normalized.pop_back();
        }
        base = SplitSegments(normalized);
    }
    std::vector<std::wstring> total = base;
    total.insert(total.end(), segments.begin(), segments.end());
    if (total.empty()) {
        combined.path = L"/";
    } else {
        combined.path.clear();
        for (const auto& segment : total) {
            combined.path.push_back(L'/');
            combined.path += segment;
        }
        combined.path.push_back(L'/');
    }
    combined.canonicalUrl = BuildCanonicalUrl(combined);
    return combined;
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
    ULONG attributes = SFGAO_STORAGE | SFGAO_CANCOPY;
    const bool isDirectory = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    if (isDirectory) {
        attributes |= SFGAO_FOLDER | SFGAO_FILESYSANCESTOR | SFGAO_STORAGEANCESTOR | SFGAO_HASSUBFOLDER;
    } else {
        attributes |= SFGAO_STREAM;
    }
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_READONLY) != 0) {
        attributes |= SFGAO_READONLY;
    } else {
        attributes |= SFGAO_CANDELETE | SFGAO_CANMOVE | SFGAO_CANRENAME;
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

class EnumerationState : public std::enable_shared_from_this<EnumerationState> {
public:
    EnumerationState(const FtpUrlParts& parts, std::vector<std::wstring> segments, std::vector<std::uint8_t> absolute,
                     SHCONTF flags, HWND owner)
        : rootParts_(parts),
          pathSegments_(std::move(segments)),
          absolutePidlBytes_(std::move(absolute)),
          flags_(flags),
          ownerWindow_(owner) {}

    ~EnumerationState() {
        Cancel();
        if (worker_.joinable()) {
            worker_.join();
        }
    }

    void Start() {
        worker_ = std::thread([self = shared_from_this()]() { self->WorkerProc(); });
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
        HRESULT hr = S_OK;
        bool coInitialized = false;
        HRESULT init = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (SUCCEEDED(init)) {
            coInitialized = true;
        }

        FtpConnectionOptions options;
        options.host = rootParts_.host;
        options.port = rootParts_.port;
        options.passiveMode = true;
        options.preferMlsd = true;
        std::wstring remotePath = JoinSegments(pathSegments_, {});
        if (!remotePath.empty() && remotePath != L"/") {
            options.initialPath = remotePath;
        }

        FtpCredential credential;
        credential.userName = rootParts_.userName.empty() ? L"anonymous" : rootParts_.userName;
        credential.password = rootParts_.password;

        FtpClient client;
        std::vector<FtpDirectoryEntry> entries;
        hr = client.ListDirectory(options, &credential, &entries, ownerWindow_);
        if (SUCCEEDED(hr)) {
            for (const auto& entry : entries) {
                if (cancelled_.load(std::memory_order_acquire)) {
                    hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    break;
                }
                if (!ShouldInclude(entry)) {
                    continue;
                }
                WIN32_FIND_DATAW findData{};
                findData.dwFileAttributes = entry.attributes == 0
                                                ? (entry.isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_ARCHIVE)
                                                : entry.attributes;
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
        }

        if (cancelled_.load(std::memory_order_acquire) && SUCCEEDED(hr)) {
            hr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (SUCCEEDED(hr) && !absolutePidlBytes_.empty()) {
            PIDLIST_ABSOLUTE notify = reinterpret_cast<PIDLIST_ABSOLUTE>(CoTaskMemAlloc(absolutePidlBytes_.size()));
            if (notify) {
                std::memcpy(notify, absolutePidlBytes_.data(), absolutePidlBytes_.size());
                SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_IDLIST, notify, nullptr);
                CoTaskMemFree(notify);
            }
        }

        {
            std::lock_guard<std::mutex> guard(mutex_);
            workerFinished_ = true;
            result_ = hr;
        }
        cv_.notify_all();

        if (coInitialized) {
            CoUninitialize();
        }
    }

    bool ShouldInclude(const FtpDirectoryEntry& entry) const {
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
        if ((flags_ & SHCONTF_INCLUDEHIDDEN) == 0) {
            const DWORD attributes = entry.attributes;
            if ((attributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) != 0) {
                return false;
            }
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
    FtpUrlParts rootParts_{};
    std::vector<std::wstring> pathSegments_;
    std::vector<std::uint8_t> absolutePidlBytes_;
    SHCONTF flags_ = 0;
    HWND ownerWindow_ = nullptr;
};

class FtpEnumIDList : public IEnumIDList {
public:
    explicit FtpEnumIDList(std::shared_ptr<EnumerationState> state, size_t index)
        : refCount_(1), state_(std::move(state)), currentIndex_(index) {}

    // IUnknown
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

    // IEnumIDList
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
        auto clone = new (std::nothrow) FtpEnumIDList(state_, currentIndex_);
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

}  // namespace

FtpShellFolder::FtpShellFolder() {
    ModuleAddRef();
}

FtpShellFolder::FtpShellFolder(const FtpUrlParts& root, const std::vector<std::wstring>& segments)
    : rootParts_(root), pathSegments_(segments) {
    ModuleAddRef();
    EnsurePidl();
}

FtpShellFolder::~FtpShellFolder() {
    ModuleRelease();
}

HRESULT FtpShellFolder::Create(const FtpUrlParts& root, const std::vector<std::wstring>& segments, REFIID riid,
                               void** ppv) {
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    FtpShellFolder* folder = new (std::nothrow) FtpShellFolder(root, segments);
    if (!folder) {
        return E_OUTOFMEMORY;
    }
    HRESULT hr = folder->QueryInterface(riid, ppv);
    folder->Release();
    return hr;
}

IFACEMETHODIMP FtpShellFolder::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IShellFolder || riid == IID_IShellFolder2) {
        *object = static_cast<IShellFolder*>(this);
    } else if (riid == IID_IPersist || riid == IID_IPersistFolder || riid == IID_IPersistFolder2) {
        *object = static_cast<IPersistFolder2*>(this);
    } else {
        *object = nullptr;
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

IFACEMETHODIMP_(ULONG) FtpShellFolder::AddRef() {
    return ++refCount_;
}

IFACEMETHODIMP_(ULONG) FtpShellFolder::Release() {
    ULONG count = --refCount_;
    if (count == 0) {
        delete this;
    }
    return count;
}

HRESULT FtpShellFolder::EnsurePidl() {
    if (initialized_ && absolutePidl_) {
        return S_OK;
    }
    FtpUrlParts parts = CombineParts(rootParts_, pathSegments_);
    absolutePidl_ = CreatePidlFromFtpUrl(parts);
    initialized_ = absolutePidl_ != nullptr;
    return initialized_ ? S_OK : E_OUTOFMEMORY;
}

HRESULT FtpShellFolder::ParseInputToSegments(std::wstring_view input, std::vector<std::wstring>* segments,
                                             bool* isDirectory) const {
    if (!segments || !isDirectory) {
        return E_POINTER;
    }
    segments->clear();
    if (input.empty()) {
        *isDirectory = true;
        return S_OK;
    }
    bool trailing = false;
    if (!input.empty() && (input.back() == L'/' || input.back() == L'\\')) {
        trailing = true;
        input.remove_suffix(1);
    }
    std::wstring buffer(input);
    std::wstring current;
    for (wchar_t ch : buffer) {
        if (ch == L'/' || ch == L'\\') {
            if (!current.empty()) {
                segments->push_back(current);
                current.clear();
            }
            continue;
        }
        current.push_back(ch);
    }
    if (!current.empty()) {
        segments->push_back(current);
    }
    *isDirectory = trailing;
    if (!trailing && !segments->empty()) {
        const std::wstring& final = segments->back();
        if (final.find(L'.') != std::wstring::npos) {
            *isDirectory = false;
        } else {
            *isDirectory = true;
        }
    }
    return S_OK;
}

bool FtpShellFolder::ExtractRelativeSegments(PCUIDLIST_RELATIVE pidl, std::vector<std::wstring>* segments,
                                             bool* isDirectory) const {
    if (!segments) {
        return false;
    }
    segments->clear();
    if (!absolutePidl_) {
        return false;
    }
    UniquePidl combined(ILCombine(absolutePidl_.get(), pidl));
    if (!combined) {
        return false;
    }
    FtpUrlParts parts;
    std::vector<std::wstring> parsedSegments;
    bool terminal = true;
    if (!TryParseFtpPidl(combined.get(), &parts, &parsedSegments, &terminal)) {
        return false;
    }
    if (parsedSegments.size() < pathSegments_.size()) {
        return false;
    }
    if (!std::equal(pathSegments_.begin(), pathSegments_.end(), parsedSegments.begin())) {
        return false;
    }
    segments->assign(parsedSegments.begin() + pathSegments_.size(), parsedSegments.end());
    if (isDirectory) {
        *isDirectory = terminal;
    }
    return true;
}

std::wstring FtpShellFolder::BuildFolderPath(const std::vector<std::wstring>& extra) const {
    return JoinSegments(pathSegments_, extra);
}

std::wstring FtpShellFolder::BuildFileName(const std::vector<std::wstring>& segments, std::wstring* directoryOut) const {
    if (segments.empty()) {
        return {};
    }
    std::vector<std::wstring> folderSegments(segments.begin(), segments.end() - 1);
    if (directoryOut) {
        *directoryOut = JoinSegments(pathSegments_, folderSegments);
    }
    return segments.back();
}

HRESULT FtpShellFolder::DownloadFileToStream(const std::vector<std::wstring>& segments, ComPtr<IStream>* stream) const {
    if (!stream) {
        return E_POINTER;
    }
    stream->Reset();
    std::wstring directory;
    std::wstring fileName = BuildFileName(segments, &directory);
    if (fileName.empty()) {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    wchar_t tempPath[MAX_PATH];
    DWORD written = GetTempPathW(static_cast<DWORD>(std::size(tempPath)), tempPath);
    if (written == 0 || written > std::size(tempPath)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    wchar_t tempFile[MAX_PATH];
    if (!GetTempFileNameW(tempPath, L"ftp", 0, tempFile)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    FtpConnectionOptions options;
    options.host = rootParts_.host;
    options.port = rootParts_.port;
    options.passiveMode = true;
    options.initialPath = directory == L"/" ? std::wstring() : directory;

    FtpCredential credential;
    credential.userName = rootParts_.userName.empty() ? L"anonymous" : rootParts_.userName;
    credential.password = rootParts_.password;

    FtpClient client;
    HRESULT hr = client.DownloadFile(options, &credential, fileName, tempFile, nullptr, nullptr);
    if (FAILED(hr)) {
        DeleteFileW(tempFile);
        return hr;
    }

    hr = SHCreateStreamOnFileEx(tempFile, STGM_READ | STGM_SHARE_DENY_NONE | STGM_DELETEONRELEASE, FILE_ATTRIBUTE_NORMAL,
                                FALSE, nullptr, stream->GetAddressOf());
    if (FAILED(hr)) {
        DeleteFileW(tempFile);
        return hr;
    }
    return S_OK;
}

HRESULT FtpShellFolder::BindToChild(const std::vector<std::wstring>& segments, REFIID riid, void** ppv) const {
    return FtpShellFolder::Create(rootParts_, segments, riid, ppv);
}

IFACEMETHODIMP FtpShellFolder::ParseDisplayName(HWND, IBindCtx*, PWSTR pszName, ULONG* pchEaten, PIDLIST_RELATIVE* ppidl,
                                               ULONG* pdwAttributes) {
    if (!ppidl) {
        return E_POINTER;
    }
    *ppidl = nullptr;
    if (!pszName) {
        return E_INVALIDARG;
    }
    std::wstring input(pszName);
    std::vector<std::wstring> segments;
    bool isDirectory = true;

    if (input.rfind(L"ftp://", 0) == 0) {
        FtpUrlParts parts;
        if (!TryParseFtpUrl(input, &parts)) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (!EqualsIgnoreCase(parts.host, rootParts_.host) || parts.port != rootParts_.port) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        std::vector<std::wstring> absoluteSegments;
        bool terminal = true;
        UniquePidl full = CreatePidlFromFtpUrl(parts);
        if (!full) {
            return E_OUTOFMEMORY;
        }
        if (!TryParseFtpPidl(full.get(), &parts, &absoluteSegments, &terminal)) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (absoluteSegments.size() < pathSegments_.size()) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (!std::equal(pathSegments_.begin(), pathSegments_.end(), absoluteSegments.begin())) {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        segments.assign(absoluteSegments.begin() + pathSegments_.size(), absoluteSegments.end());
        isDirectory = terminal;
    } else {
        HRESULT hr = ParseInputToSegments(input, &segments, &isDirectory);
        if (FAILED(hr)) {
            return hr;
        }
    }

    PidlBuilder builder;
    if (segments.empty()) {
        UniquePidl empty = builder.Finalize();
        if (!empty) {
            return E_OUTOFMEMORY;
        }
        *ppidl = reinterpret_cast<PIDLIST_RELATIVE>(empty.release());
    } else {
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
    }

    if (pchEaten) {
        *pchEaten = static_cast<ULONG>(wcslen(pszName));
    }
    if (pdwAttributes) {
        *pdwAttributes = isDirectory ? (SFGAO_FOLDER | SFGAO_HASSUBFOLDER) : SFGAO_STREAM;
    }
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::EnumObjects(HWND hwnd, SHCONTF grfFlags, IEnumIDList** ppenumIDList) {
    if (!ppenumIDList) {
        return E_POINTER;
    }
    *ppenumIDList = nullptr;
    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        return hr;
    }
    std::vector<std::uint8_t> absoluteBytes = SerializeFtpPidl(absolutePidl_.get());
    auto state = std::make_shared<EnumerationState>(rootParts_, pathSegments_, std::move(absoluteBytes), grfFlags, hwnd);
    if (!state) {
        return E_OUTOFMEMORY;
    }
    auto enumerator = new (std::nothrow) FtpEnumIDList(state, 0);
    if (!enumerator) {
        return E_OUTOFMEMORY;
    }
    state->Start();
    *ppenumIDList = enumerator;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::BindToObject(PCUIDLIST_RELATIVE pidl, IBindCtx*, REFIID riid, void** ppv) {
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    std::vector<std::wstring> segments;
    bool directory = true;
    if (!ExtractRelativeSegments(pidl, &segments, &directory)) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (!directory) {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }
    std::vector<std::wstring> combined = pathSegments_;
    combined.insert(combined.end(), segments.begin(), segments.end());
    return BindToChild(combined, riid, ppv);
}

IFACEMETHODIMP FtpShellFolder::BindToStorage(PCUIDLIST_RELATIVE pidl, IBindCtx*, REFIID riid, void** ppv) {
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

IFACEMETHODIMP FtpShellFolder::CompareIDs(LPARAM lParam, PCUIDLIST_RELATIVE pidl1, PCUIDLIST_RELATIVE pidl2) {
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
                if (sizeLeft < sizeRight) {
                    comparison = -1;
                } else if (sizeLeft > sizeRight) {
                    comparison = 1;
                }
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
        if (hasLeft) {
            nameLeft.assign(left.cFileName);
        } else {
            TryGetNameFromPidl(pidl1, &nameLeft);
        }
        if (hasRight) {
            nameRight.assign(right.cFileName);
        } else {
            TryGetNameFromPidl(pidl2, &nameRight);
        }
        comparison = _wcsicmp(nameLeft.c_str(), nameRight.c_str());
    }
    if (comparison < 0) {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0xFFFF);
    }
    if (comparison > 0) {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 1);
    }
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::CreateViewObject(HWND, REFIID riid, void** ppv) {
    if (!ppv) {
        return E_POINTER;
    }
    *ppv = nullptr;
    if (riid != IID_IShellView) {
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
        return hr;
    }
    hr = SHCreateShellFolderView(&create, reinterpret_cast<IShellView**>(ppv));
    if (create.pshf) {
        create.pshf->Release();
    }
    return hr;
}

IFACEMETHODIMP FtpShellFolder::GetAttributesOf(UINT cidl, PCUITEMID_CHILD_ARRAY apidl, ULONG* rgfInOut) {
    if (!rgfInOut) {
        return E_POINTER;
    }
    ULONG mask = *rgfInOut;
    if (cidl == 0) {
        ULONG folderFlags = SFGAO_FOLDER | SFGAO_STORAGE | SFGAO_FILESYSANCESTOR | SFGAO_HASSUBFOLDER | SFGAO_CANCOPY |
                            SFGAO_CANMOVE | SFGAO_CANRENAME | SFGAO_CANDELETE;
        *rgfInOut = mask == 0 ? folderFlags : (folderFlags & mask);
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
            if (type == ItemType::Directory) {
                itemFlags = SFGAO_FOLDER | SFGAO_STORAGE | SFGAO_FILESYSANCESTOR | SFGAO_HASSUBFOLDER |
                            SFGAO_CANCOPY | SFGAO_CANMOVE | SFGAO_CANRENAME | SFGAO_CANDELETE;
            } else {
                itemFlags = SFGAO_STREAM | SFGAO_STORAGE | SFGAO_CANCOPY | SFGAO_CANMOVE | SFGAO_CANRENAME |
                            SFGAO_CANDELETE;
            }
        }
        result &= itemFlags;
    }
    if (mask != 0) {
        result &= mask;
    }
    *rgfInOut = result;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::GetUIObjectOf(HWND hwnd, UINT cidl, PCUITEMID_CHILD_ARRAY apidl, REFIID riid,
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

    if (riid == IID_IDataObject) {
        return SHCreateDataObject(absolutePidl_.get(), cidl, apidl, nullptr, riid, ppv);
    }

    if (riid == IID_IContextMenu || riid == IID_IContextMenu2 || riid == IID_IContextMenu3) {
        DEFCONTEXTMENU def{};
        def.hwnd = hwnd;
        def.pidlFolder = absolutePidl_.get();
        def.psf = this;
        def.cidl = cidl;
        def.apidl = apidl;

        Microsoft::WRL::ComPtr<IDataObject> dataObject;
        if (cidl > 0) {
            hr = SHCreateDataObject(absolutePidl_.get(), cidl, apidl, nullptr, IID_PPV_ARGS(&dataObject));
            if (FAILED(hr)) {
                return hr;
            }
            def.pdtobj = dataObject.Get();
        }

        return SHCreateDefaultContextMenu(&def, riid, ppv);
    }

    if (riid == IID_IQueryAssociations) {
        Microsoft::WRL::ComPtr<IQueryAssociations> associations;
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
                WIN32_FIND_DATAW findData{};
                std::wstring name;
                if (TryGetFindData(apidl[0], &findData)) {
                    name.assign(findData.cFileName);
                } else if (TryGetNameFromPidl(apidl[0], &name)) {
                    // Name already assigned.
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
                assoc = nullptr;
            }
        }

        hr = associations->Init(flags, assoc, nullptr, hwnd);
        if (FAILED(hr)) {
            if (!assoc) {
                hr = associations->Init(ASSOCF_INIT_DEFAULTTOSTAR | ASSOCF_INIT_IGNOREUNKNOWN, L"*", nullptr, hwnd);
                if (FAILED(hr)) {
                    return hr;
                }
            } else {
                return hr;
            }
        }
        *ppv = associations.Detach();
        return S_OK;
    }

    return E_NOINTERFACE;
}

IFACEMETHODIMP FtpShellFolder::GetDisplayNameOf(PCUITEMID_CHILD pidl, SHGDNF uFlags, STRRET* pName) {
    if (!pidl || !pName) {
        return E_POINTER;
    }
    if (uFlags & SHGDN_FORPARSING) {
        if (uFlags & SHGDN_INFOLDER) {
            WIN32_FIND_DATAW findData{};
            if (TryGetFindData(pidl, &findData)) {
                return AssignToStrRet(findData.cFileName, pName);
            }
            std::wstring name;
            if (TryGetNameFromPidl(pidl, &name)) {
                return AssignToStrRet(name, pName);
            }
            return E_FAIL;
        }
        HRESULT hr = EnsurePidl();
        if (FAILED(hr)) {
            return hr;
        }
        UniquePidl combined(ILCombine(absolutePidl_.get(), pidl));
        if (!combined) {
            return E_OUTOFMEMORY;
        }
        std::wstring url = BuildUrlFromFtpPidl(combined.get());
        if (url.empty()) {
            return E_FAIL;
        }
        return AssignToStrRet(url, pName);
    }
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

IFACEMETHODIMP FtpShellFolder::SetNameOf(HWND hwnd, PCUITEMID_CHILD pidl, PCWSTR pszName, SHGDNF, PIDLIST_RELATIVE* ppidlOut) {
    UNREFERENCED_PARAMETER(hwnd);
    if (!pidl || !pszName) {
        return E_INVALIDARG;
    }

    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        return hr;
    }

    std::wstring newName(pszName);
    if (newName.empty() || newName == L"." || newName == L"..") {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }
    if (newName.find_first_of(L"/\\:") != std::wstring::npos) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    std::vector<std::wstring> segments;
    bool isDirectory = true;
    if (!ExtractRelativeSegments(pidl, &segments, &isDirectory) || segments.empty()) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::wstring directory;
    std::wstring oldName = BuildFileName(segments, &directory);
    if (oldName.empty()) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (oldName == newName) {
        if (ppidlOut) {
            *ppidlOut = nullptr;
        }
        return S_OK;
    }

    FtpConnectionOptions options;
    options.host = rootParts_.host;
    options.port = rootParts_.port;
    options.passiveMode = true;
    std::wstring remoteDirectory = directory == L"/" ? std::wstring() : directory;

    FtpCredential credential;
    credential.userName = rootParts_.userName.empty() ? L"anonymous" : rootParts_.userName;
    credential.password = rootParts_.password;

    hr = RenameRemoteItem(options, credential, remoteDirectory, oldName, newName);
    if (FAILED(hr)) {
        return hr;
    }

    WIN32_FIND_DATAW findData{};
    const bool hasFindData = TryGetFindData(pidl, &findData) &&
                             SUCCEEDED(StringCchCopyW(findData.cFileName, ARRAYSIZE(findData.cFileName), newName.c_str()));

    PidlBuilder builder;
    for (size_t i = 0; i < segments.size(); ++i) {
        const std::wstring& segment = (i + 1 == segments.size()) ? newName : segments[i];
        ItemType type = (i + 1 == segments.size() && !isDirectory) ? ItemType::File : ItemType::Directory;
        ComponentDefinition nameComponent{ComponentType::Name, segment.data(), segment.size() * sizeof(wchar_t)};
        if (i + 1 == segments.size() && hasFindData) {
            ComponentDefinition dataComponent{ComponentType::FindData, &findData, sizeof(findData)};
            hr = builder.Append(type, {nameComponent, dataComponent});
        } else {
            hr = builder.Append(type, {nameComponent});
        }
        if (FAILED(hr)) {
            return hr;
        }
    }
    UniquePidl renamedRelative = builder.Finalize();
    if (!renamedRelative) {
        return E_OUTOFMEMORY;
    }

    if (ppidlOut) {
        UniquePidl clone = CloneRelativeFtpPidl(reinterpret_cast<PCUIDLIST_RELATIVE>(renamedRelative.get()));
        if (!clone) {
            return E_OUTOFMEMORY;
        }
        *ppidlOut = reinterpret_cast<PIDLIST_RELATIVE>(clone.release());
    }

    if (SUCCEEDED(EnsurePidl())) {
        UniquePidl oldAbsolute(ILCombine(absolutePidl_.get(), pidl));
        UniquePidl newAbsolute(ILCombine(absolutePidl_.get(), renamedRelative.get()));
        if (oldAbsolute && newAbsolute) {
            LONG eventId = isDirectory ? SHCNE_RENAMEFOLDER : SHCNE_RENAMEITEM;
            SHChangeNotify(eventId, SHCNF_IDLIST, oldAbsolute.get(), newAbsolute.get());
        }
    }

    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::GetDefaultSearchGUID(GUID* pguid) {
    if (!pguid) {
        return E_POINTER;
    }
    *pguid = kFtpSearchProviderGuid;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::EnumSearches(IEnumExtraSearch** ppEnum) {
    if (!ppEnum) {
        return E_POINTER;
    }
    *ppEnum = nullptr;
    HRESULT hr = S_OK;
    FtpUrlParts parts = CombineParts(rootParts_, pathSegments_);
    std::wstring searchUrl = L"search-ms:query=%1&crumb=location:" + parts.canonicalUrl;

    EXTRASEARCH search{};
    search.guidSearch = kFtpSearchProviderGuid;
    hr = StringCchCopyW(search.wszFriendlyName, ARRAYSIZE(search.wszFriendlyName), L"Search this FTP site");
    if (FAILED(hr)) {
        return hr;
    }
    hr = StringCchCopyW(search.wszUrl, ARRAYSIZE(search.wszUrl), searchUrl.c_str());
    if (FAILED(hr)) {
        return hr;
    }

    std::vector<EXTRASEARCH> entries;
    entries.push_back(search);
    auto enumerator = new (std::nothrow) FtpSearchEnumerator(std::move(entries));
    if (!enumerator) {
        return E_OUTOFMEMORY;
    }
    *ppEnum = enumerator;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::GetDefaultColumn(DWORD, ULONG* pSort, ULONG* pDisplay) {
    if (!pSort || !pDisplay) {
        return E_POINTER;
    }
    *pSort = 0;
    *pDisplay = 0;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::GetDefaultColumnState(UINT iColumn, SHCOLSTATEF* pcsFlags) {
    if (!pcsFlags) {
        return E_POINTER;
    }
    if (iColumn >= kColumnCount) {
        return E_INVALIDARG;
    }
    *pcsFlags = kColumnDefinitions[iColumn].state;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::GetDetailsEx(PCUITEMID_CHILD pidl, const SHCOLUMNID* pscid, VARIANT* pv) {
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

IFACEMETHODIMP FtpShellFolder::GetDetailsOf(PCUITEMID_CHILD pidl, UINT iColumn, SHELLDETAILS* pDetails) {
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

IFACEMETHODIMP FtpShellFolder::MapColumnToSCID(UINT iColumn, SHCOLUMNID* pscid) {
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

IFACEMETHODIMP FtpShellFolder::GetClassID(CLSID* pClassID) {
    if (!pClassID) {
        return E_POINTER;
    }
    *pClassID = CLSID_NULL;
    return S_OK;
}

IFACEMETHODIMP FtpShellFolder::Initialize(PCIDLIST_ABSOLUTE pidl) {
    if (!pidl) {
        return E_INVALIDARG;
    }
    FtpUrlParts parts;
    std::vector<std::wstring> segments;
    bool directory = true;
    if (!TryParseFtpPidl(pidl, &parts, &segments, &directory)) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    rootParts_ = parts;
    rootParts_.path.clear();
    pathSegments_ = segments;
    if (!directory && !pathSegments_.empty()) {
        pathSegments_.pop_back();
    }
    absolutePidl_ = ClonePidl(pidl);
    initialized_ = absolutePidl_ != nullptr;
    return initialized_ ? S_OK : E_OUTOFMEMORY;
}

IFACEMETHODIMP FtpShellFolder::GetCurFolder(PIDLIST_ABSOLUTE* ppidl) {
    if (!ppidl) {
        return E_POINTER;
    }
    HRESULT hr = EnsurePidl();
    if (FAILED(hr)) {
        return hr;
    }
    *ppidl = ClonePidl(absolutePidl_.get()).release();
    return *ppidl ? S_OK : E_OUTOFMEMORY;
}

}  // namespace shelltabs::ftp

