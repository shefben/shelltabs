#include "HttpPidl.h"
#include "Utilities.h"

#include <windows.h>
#include <objbase.h>

#include <cassert>
#include <cstdio>
#include <string>
#include <vector>

using namespace shelltabs;
using namespace shelltabs::http;

namespace {

void TestCreatePidlFromHttpUrl() {
    HttpUrlParts parts;
    parts.host = L"myrient.erista.me";
    parts.basePath = L"/files/";
    parts.port = 443;
    parts.useHttps = true;
    parts.canonicalUrl = L"https://myrient.erista.me/files/";
    parts.displayName = L"Myrient";

    UniquePidl pidl = CreatePidlFromHttpUrl(parts);
    assert(pidl != nullptr);

    // Walk to the first SHITEMID
    auto* raw = reinterpret_cast<PCITEMID_CHILD>(pidl.get());
    assert(raw->mkid.cb > 0);
    assert(IsHttpItemId(raw->mkid));
    assert(GetItemType(raw->mkid) == ItemType::Root);

    std::wstring host;
    assert(TryGetComponentString(raw->mkid, ComponentType::Host, &host));
    assert(host == L"myrient.erista.me");

    std::wstring basePath;
    assert(TryGetComponentString(raw->mkid, ComponentType::BasePath, &basePath));
    assert(basePath == L"/files/");

    std::uint16_t port = 0;
    assert(TryGetComponentUint16(raw->mkid, ComponentType::Port, &port));
    assert(port == 443);

    std::wstring scheme;
    assert(TryGetComponentString(raw->mkid, ComponentType::Scheme, &scheme));
    assert(scheme == L"https");

    std::wstring name;
    assert(TryGetComponentString(raw->mkid, ComponentType::Name, &name));
    assert(name == L"Myrient");

    wprintf(L"  PASS: TestCreatePidlFromHttpUrl\n");
}

void TestPidlRoundTrip() {
    HttpUrlParts parts;
    parts.host = L"example.com";
    parts.basePath = L"/pub/";
    parts.port = 8080;
    parts.useHttps = false;
    parts.canonicalUrl = L"http://example.com:8080/pub/";
    parts.displayName = L"Example";

    UniquePidl pidl = CreatePidlFromHttpUrl(parts);
    assert(pidl != nullptr);

    HttpUrlParts parsed;
    std::vector<std::wstring> segments;
    bool terminalIsDir = false;
    bool ok = TryParseHttpPidl(pidl.get(), &parsed, &segments, &terminalIsDir);
    assert(ok);
    assert(parsed.host == L"example.com");
    assert(parsed.basePath == L"/pub/");
    assert(parsed.port == 8080);
    assert(parsed.useHttps == false);
    assert(parsed.displayName == L"Example");
    assert(segments.empty());  // Root pidl has no extra segments

    wprintf(L"  PASS: TestPidlRoundTrip\n");
}

void TestPidlBuilderMultiLevel() {
    // Build a Root + Directory + File PIDL
    PidlBuilder builder;

    std::wstring host = L"files.example.org";
    std::wstring basePath = L"/";
    std::uint16_t port = 443;
    std::wstring scheme = L"https";
    std::wstring displayName = L"Files";

    HRESULT hr = builder.Append(ItemType::Root, {
        {ComponentType::Host, host.c_str(), (host.size() + 1) * sizeof(wchar_t)},
        {ComponentType::BasePath, basePath.c_str(), (basePath.size() + 1) * sizeof(wchar_t)},
        {ComponentType::Port, &port, sizeof(port)},
        {ComponentType::Scheme, scheme.c_str(), (scheme.size() + 1) * sizeof(wchar_t)},
        {ComponentType::Name, displayName.c_str(), (displayName.size() + 1) * sizeof(wchar_t)},
    });
    assert(SUCCEEDED(hr));

    std::wstring dirName = L"Games";
    WIN32_FIND_DATAW dirFindData = {};
    dirFindData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY;
    wcscpy_s(dirFindData.cFileName, dirName.c_str());

    hr = builder.Append(ItemType::Directory, {
        {ComponentType::Name, dirName.c_str(), (dirName.size() + 1) * sizeof(wchar_t)},
        {ComponentType::FindData, &dirFindData, sizeof(dirFindData)},
    });
    assert(SUCCEEDED(hr));

    std::wstring fileName = L"rom.zip";
    WIN32_FIND_DATAW fileFindData = {};
    fileFindData.dwFileAttributes = FILE_ATTRIBUTE_NORMAL;
    fileFindData.nFileSizeLow = 12345678;
    wcscpy_s(fileFindData.cFileName, fileName.c_str());

    hr = builder.Append(ItemType::File, {
        {ComponentType::Name, fileName.c_str(), (fileName.size() + 1) * sizeof(wchar_t)},
        {ComponentType::FindData, &fileFindData, sizeof(fileFindData)},
    });
    assert(SUCCEEDED(hr));

    assert(builder.item_count() == 3);
    UniquePidl pidl = builder.Finalize();
    assert(pidl != nullptr);

    // Parse back
    HttpUrlParts parsed;
    std::vector<std::wstring> segments;
    bool terminalIsDir = false;
    bool ok = TryParseHttpPidl(pidl.get(), &parsed, &segments, &terminalIsDir);
    assert(ok);
    assert(parsed.host == L"files.example.org");
    assert(parsed.basePath == L"/");
    assert(parsed.port == 443);
    assert(parsed.useHttps == true);
    assert(parsed.displayName == L"Files");
    assert(segments.size() == 2);
    assert(segments[0] == L"Games");
    assert(segments[1] == L"rom.zip");
    assert(terminalIsDir == false);

    wprintf(L"  PASS: TestPidlBuilderMultiLevel\n");
}

void TestSerializeDeserialize() {
    HttpUrlParts parts;
    parts.host = L"test.org";
    parts.basePath = L"/data/";
    parts.port = 443;
    parts.useHttps = true;
    parts.canonicalUrl = L"https://test.org/data/";
    parts.displayName = L"Test";

    UniquePidl pidl = CreatePidlFromHttpUrl(parts);
    assert(pidl != nullptr);

    std::vector<std::uint8_t> bytes = SerializeHttpPidl(pidl.get());
    assert(!bytes.empty());

    // Verify the serialized bytes start with a valid SHITEMID
    auto* restored = reinterpret_cast<PCIDLIST_ABSOLUTE>(bytes.data());
    auto* firstItem = reinterpret_cast<const SHITEMID*>(
        reinterpret_cast<const std::byte*>(restored));
    assert(firstItem->cb > 0);
    assert(IsHttpItemId(*firstItem));
    assert(GetItemType(*firstItem) == ItemType::Root);

    wprintf(L"  PASS: TestSerializeDeserialize\n");
}

void TestBuildUrlFromPidl() {
    PidlBuilder builder;

    std::wstring host = L"archive.org";
    std::wstring basePath = L"/download/";
    std::uint16_t port = 443;
    std::wstring scheme = L"https";
    std::wstring displayName = L"Archive";

    HRESULT hr = builder.Append(ItemType::Root, {
        {ComponentType::Host, host.c_str(), (host.size() + 1) * sizeof(wchar_t)},
        {ComponentType::BasePath, basePath.c_str(), (basePath.size() + 1) * sizeof(wchar_t)},
        {ComponentType::Port, &port, sizeof(port)},
        {ComponentType::Scheme, scheme.c_str(), (scheme.size() + 1) * sizeof(wchar_t)},
        {ComponentType::Name, displayName.c_str(), (displayName.size() + 1) * sizeof(wchar_t)},
    });
    assert(SUCCEEDED(hr));

    std::wstring dirName = L"collection";
    hr = builder.Append(ItemType::Directory, {
        {ComponentType::Name, dirName.c_str(), (dirName.size() + 1) * sizeof(wchar_t)},
    });
    assert(SUCCEEDED(hr));

    std::wstring fileName = L"file.txt";
    hr = builder.Append(ItemType::File, {
        {ComponentType::Name, fileName.c_str(), (fileName.size() + 1) * sizeof(wchar_t)},
    });
    assert(SUCCEEDED(hr));

    UniquePidl pidl = builder.Finalize();
    std::wstring url = BuildUrlFromHttpPidl(pidl.get());
    assert(url == L"https://archive.org/download/collection/file.txt");

    wprintf(L"  PASS: TestBuildUrlFromPidl\n");
}

void TestHttpUrlParsing() {
    HttpUrlParts parts;
    bool ok = TryParseHttpUrl(L"https://myrient.erista.me/files/", &parts);
    assert(ok);
    assert(parts.host == L"myrient.erista.me");
    assert(parts.basePath == L"/files/");
    assert(parts.port == 443);
    assert(parts.useHttps == true);

    ok = TryParseHttpUrl(L"http://example.com:8080/pub/data/", &parts);
    assert(ok);
    assert(parts.host == L"example.com");
    assert(parts.basePath == L"/pub/data/");
    assert(parts.port == 8080);
    assert(parts.useHttps == false);

    // Invalid URLs
    ok = TryParseHttpUrl(L"ftp://example.com/", &parts);
    assert(!ok);

    ok = TryParseHttpUrl(L"not a url", &parts);
    assert(!ok);

    wprintf(L"  PASS: TestHttpUrlParsing\n");
}

void TestIsHttpItemId() {
    // Build a valid HTTP PIDL and verify detection
    PidlBuilder builder;
    std::wstring name = L"test";
    HRESULT hr = builder.Append(ItemType::File, {
        {ComponentType::Name, name.c_str(), (name.size() + 1) * sizeof(wchar_t)},
    });
    assert(SUCCEEDED(hr));

    UniquePidl pidl = builder.Finalize();
    auto* item = reinterpret_cast<const SHITEMID*>(
        reinterpret_cast<const std::byte*>(pidl.get()));
    assert(IsHttpItemId(*item));
    assert(GetItemType(*item) == ItemType::File);

    // A zeroed-out SHITEMID should NOT be detected as HTTP
    SHITEMID fake = {};
    fake.cb = 16;
    assert(!IsHttpItemId(fake));

    wprintf(L"  PASS: TestIsHttpItemId\n");
}

}  // namespace

int main() {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    wprintf(L"Running HttpPidl tests...\n");

    TestCreatePidlFromHttpUrl();
    TestPidlRoundTrip();
    TestPidlBuilderMultiLevel();
    TestSerializeDeserialize();
    TestBuildUrlFromPidl();
    TestHttpUrlParsing();
    TestIsHttpItemId();

    wprintf(L"\nAll HttpPidl tests passed!\n");

    CoUninitialize();
    return 0;
}
