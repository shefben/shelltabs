#include "Utilities.h"
#include "HttpPidl.h"

#include <shlobj.h>
#include <shlwapi.h>
#include <urlmon.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <string>
#include <utility>

#include <wrl/client.h>

namespace shelltabs {

namespace {

std::wstring TakeBstr(BSTR value) {
    if (!value) {
        return {};
    }
    std::wstring result(value, SysStringLen(value));
    SysFreeString(value);
    return result;
}

std::wstring GetUriPropertyString(IUri* uri, Uri_PROPERTY property) {
    BSTR value = nullptr;
    if (SUCCEEDED(uri->GetPropertyBSTR(property, &value, 0)) && value) {
        return TakeBstr(value);
    }
    return {};
}

}  // namespace

bool TryParseHttpUrl(const std::wstring& url, HttpUrlParts* parts) {
    if (!parts) {
        return false;
    }

    Microsoft::WRL::ComPtr<IUri> uri;
    HRESULT hr = CreateUri(url.c_str(),
                           Uri_CREATE_CANONICALIZE | Uri_CREATE_DECODE_EXTRA_INFO | Uri_CREATE_NO_IE_SETTINGS,
                           0, &uri);
    if (FAILED(hr) || !uri) {
        return false;
    }

    DWORD scheme = 0;
    hr = uri->GetScheme(&scheme);
    if (FAILED(hr)) {
        return false;
    }
    if (scheme != URL_SCHEME_HTTP && scheme != URL_SCHEME_HTTPS) {
        return false;
    }

    parts->useHttps = (scheme == URL_SCHEME_HTTPS);
    parts->host = GetUriPropertyString(uri.Get(), Uri_PROPERTY_HOST);
    if (parts->host.empty()) {
        return false;
    }

    DWORD port = 0;
    hr = uri->GetPort(&port);
    if (SUCCEEDED(hr) && port != 0) {
        parts->port = static_cast<unsigned short>(port);
    } else {
        parts->port = parts->useHttps ? 443 : 80;
    }

    parts->basePath = GetUriPropertyString(uri.Get(), Uri_PROPERTY_PATH);
    if (parts->basePath.empty()) {
        parts->basePath = L"/";
    }

    // Build canonical URL
    std::wstring canonical = parts->useHttps ? L"https://" : L"http://";
    canonical += parts->host;
    const unsigned short defaultPort = parts->useHttps ? 443 : 80;
    if (parts->port != defaultPort) {
        canonical += L":" + std::to_wstring(parts->port);
    }
    canonical += parts->basePath;
    if (!canonical.empty() && canonical.back() != L'/') {
        canonical.push_back(L'/');
    }
    parts->canonicalUrl = std::move(canonical);

    return true;
}

}  // namespace shelltabs
