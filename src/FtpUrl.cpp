#include "Utilities.h"
#include "FtpPidl.h"

#include <shlobj.h>
#include <shlwapi.h>
#include <urlmon.h>
#include <wininet.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <string>
#include <utility>

#include <wrl/client.h>

namespace shelltabs {
namespace {
constexpr unsigned short kDefaultFtpPort = INTERNET_DEFAULT_FTP_PORT;

std::wstring TakeBstr(BSTR value) {
    if (!value) {
        return {};
    }
    std::wstring result(value, SysStringLen(value));
    SysFreeString(value);
    return result;
}

std::wstring GetUriPropertyString(IUri* uri, Uri_PROPERTY property) {
    if (!uri) {
        return {};
    }
    BSTR value = nullptr;
    if (FAILED(uri->GetPropertyBSTR(property, &value, 0)) || !value) {
        if (value) {
            SysFreeString(value);
        }
        return {};
    }
    return TakeBstr(value);
}

std::wstring DecodeUriComponent(const std::wstring& value) {
    if (value.empty()) {
        return value;
    }
    DWORD length = static_cast<DWORD>(value.size() + 1);
    std::wstring buffer(length, L'\0');
    const size_t copyCount = std::min<size_t>(value.size(), buffer.size() - 1);
    std::wmemcpy(buffer.data(), value.c_str(), copyCount);
    buffer[copyCount] = L'\0';
    HRESULT hr = UrlUnescapeW(buffer.data(), buffer.data(), &length, URL_UNESCAPE_AS_UTF8);
    if (hr == E_POINTER) {
        buffer.resize(length);
        const size_t secondCopy = std::min<size_t>(value.size(), buffer.size() - 1);
        std::wmemcpy(buffer.data(), value.c_str(), secondCopy);
        buffer[secondCopy] = L'\0';
        hr = UrlUnescapeW(buffer.data(), buffer.data(), &length, URL_UNESCAPE_AS_UTF8);
    }
    if (FAILED(hr) || length == 0) {
        return value;
    }
    buffer.resize(length - 1);
    return buffer;
}

std::wstring EscapeUriSegment(const std::wstring& segment) {
    if (segment.empty()) {
        return segment;
    }
    DWORD length = static_cast<DWORD>(segment.size() * 4 + 1);
    std::wstring buffer(length, L'\0');
    HRESULT hr = UrlEscapeW(segment.c_str(), buffer.data(), &length,
                            URL_ESCAPE_SEGMENT_ONLY | URL_ESCAPE_PERCENT | URL_ESCAPE_AS_UTF8);
    if (hr == E_POINTER) {
        buffer.resize(length);
        hr = UrlEscapeW(segment.c_str(), buffer.data(), &length,
                        URL_ESCAPE_SEGMENT_ONLY | URL_ESCAPE_PERCENT | URL_ESCAPE_AS_UTF8);
    }
    if (FAILED(hr) || length == 0) {
        return segment;
    }
    buffer.resize(length - 1);
    return buffer;
}

std::wstring NormalizeFtpPath(const std::wstring& path) {
    if (path.empty()) {
        return L"/";
    }
    std::wstring normalized = path;
    std::replace(normalized.begin(), normalized.end(), L'\\', L'/');
    if (normalized.front() != L'/') {
        normalized.insert(normalized.begin(), L'/');
    }
    return normalized;
}

std::wstring EncodeFtpPath(const std::wstring& path) {
    std::wstring sanitized = NormalizeFtpPath(path);
    std::wstring result;
    result.reserve(sanitized.size() * 3);
    size_t index = 0;
    while (index < sanitized.size()) {
        if (sanitized[index] == L'/') {
            result.push_back(L'/');
            ++index;
            continue;
        }
        size_t nextSlash = sanitized.find(L'/', index);
        std::wstring segment = sanitized.substr(index, nextSlash == std::wstring::npos ? sanitized.size() - index
                                                                                       : nextSlash - index);
        if (!segment.empty()) {
            result += EscapeUriSegment(segment);
        }
        if (nextSlash == std::wstring::npos) {
            break;
        }
        index = nextSlash;
    }
    if (result.empty()) {
        result = L"/";
    }
    return result;
}

std::wstring BuildCanonicalFtpUrl(const FtpUrlParts& parts) {
    std::wstring url = L"ftp://";
    if (!parts.userName.empty()) {
        url += EscapeUriSegment(parts.userName);
        if (!parts.password.empty()) {
            url += L":";
            url += EscapeUriSegment(parts.password);
        }
        url += L"@";
    }
    url += parts.host;
    if (parts.port != kDefaultFtpPort && parts.port != 0) {
        url += L":";
        url += std::to_wstring(parts.port);
    }
    url += EncodeFtpPath(parts.path);
    return url;
}

}  // namespace

bool TryParseFtpUrl(const std::wstring& url, FtpUrlParts* parts) {
    if (!parts || url.empty()) {
        return false;
    }

    Microsoft::WRL::ComPtr<IUri> uri;
    HRESULT hr = CreateUri(url.c_str(), Uri_CREATE_CANONICALIZE, 0, &uri);
    if (FAILED(hr) || !uri) {
        return false;
    }

    DWORD scheme = static_cast<DWORD>(URL_SCHEME_INVALID);
    if (FAILED(uri->GetScheme(&scheme)) || scheme != URL_SCHEME_FTP) {
        return false;
    }

    FtpUrlParts result;
    result.host = GetUriPropertyString(uri.Get(), Uri_PROPERTY_HOST);
    if (result.host.empty()) {
        return false;
    }

    std::transform(result.host.begin(), result.host.end(), result.host.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(towlower(static_cast<wint_t>(ch))); });

    result.userName = DecodeUriComponent(GetUriPropertyString(uri.Get(), Uri_PROPERTY_USER_NAME));
    result.password = DecodeUriComponent(GetUriPropertyString(uri.Get(), Uri_PROPERTY_PASSWORD));
    std::wstring path = DecodeUriComponent(GetUriPropertyString(uri.Get(), Uri_PROPERTY_PATH));
    result.path = NormalizeFtpPath(path);

    DWORD port = 0;
    if (FAILED(uri->GetPort(&port)) || port == 0 || port == static_cast<DWORD>(-1)) {
        port = kDefaultFtpPort;
    }
    result.port = static_cast<unsigned short>(port);

    if (result.userName.empty()) {
        result.userName = L"anonymous";
        result.password.clear();
    }

    result.canonicalUrl = BuildCanonicalFtpUrl(result);
    *parts = std::move(result);
    return true;
}

UniquePidl CreateFtpPidlFromUrl(const FtpUrlParts& parts) {
    if (parts.host.empty()) {
        return nullptr;
    }
    return ftp::CreatePidlFromFtpUrl(parts);
}

}  // namespace shelltabs
