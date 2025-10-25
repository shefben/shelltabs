#include "FtpShellFolder.h"

#include <shlobj.h>
#include <shlwapi.h>

#include <algorithm>
#include <cwctype>
#include <iterator>
#include <string_view>

using Microsoft::WRL::ComPtr;

namespace shelltabs::ftp {

namespace {

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
    if (riid == IID_IUnknown || riid == IID_IShellFolder) {
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

IFACEMETHODIMP FtpShellFolder::EnumObjects(HWND, SHCONTF, IEnumIDList**) {
    return E_NOTIMPL;
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

IFACEMETHODIMP FtpShellFolder::CompareIDs(LPARAM, PCUIDLIST_RELATIVE, PCUIDLIST_RELATIVE) {
    return E_NOTIMPL;
}

IFACEMETHODIMP FtpShellFolder::CreateViewObject(HWND, REFIID, void**) { return E_NOTIMPL; }

IFACEMETHODIMP FtpShellFolder::GetAttributesOf(UINT, PCUITEMID_CHILD_ARRAY, ULONG* rgfInOut) {
    if (rgfInOut) {
        *rgfInOut = 0;
    }
    return E_NOTIMPL;
}

IFACEMETHODIMP FtpShellFolder::GetUIObjectOf(HWND, UINT, PCUITEMID_CHILD_ARRAY, REFIID, UINT*, void**) {
    return E_NOTIMPL;
}

IFACEMETHODIMP FtpShellFolder::GetDisplayNameOf(PCUITEMID_CHILD, SHGDNF, STRRET*) { return E_NOTIMPL; }

IFACEMETHODIMP FtpShellFolder::SetNameOf(HWND, PCUITEMID_CHILD, PCWSTR, SHGDNF, PIDLIST_RELATIVE**) {
    return E_NOTIMPL;
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

