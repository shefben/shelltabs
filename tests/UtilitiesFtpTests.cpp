#include "Utilities.h"
#include "FtpPidl.h"
#include "FtpClient.h"

#include <windows.h>

#include <iostream>
#include <cstddef>
#include <string>
#include <vector>

namespace {

struct TestDefinition {
    const wchar_t* name;
    bool (*fn)();
};

void PrintFailure(const wchar_t* testName, const std::wstring& message) {
    std::wcerr << L"[" << testName << L"] " << message << std::endl;
}

bool TestParseFtpUrls() {
    struct FtpParseCase {
        std::wstring input;
        std::wstring expectedCanonical;
        std::wstring expectedUser;
        std::wstring expectedPassword;
        std::wstring expectedHost;
        std::wstring expectedPath;
        unsigned short expectedPort = 21;
    };

    const std::vector<FtpParseCase> cases = {
        {L"ftp://user:pass@example.com:2121/path/to/file",
         L"ftp://user:pass@example.com:2121/path/to/file", L"user", L"pass", L"example.com", L"/path/to/file", 2121},
        {L"ftp://example.com/some folder/with spaces",
         L"ftp://anonymous@example.com/some%20folder/with%20spaces", L"anonymous", L"", L"example.com",
         L"/some folder/with spaces", 21},
        {L"ftp://例子.com/文件/路径",
         L"ftp://anonymous@例子.com/%E6%96%87%E4%BB%B6/%E8%B7%AF%E5%BE%84", L"anonymous", L"", L"例子.com",
         L"/文件/路径", 21},
        {L"ftp://User:Pa%40ss@Example.com",
         L"ftp://User:Pa%40ss@example.com/", L"User", L"Pa@ss", L"example.com", L"/", 21},
    };

    bool success = true;
    for (const auto& testCase : cases) {
        shelltabs::FtpUrlParts parts;
        if (!shelltabs::TryParseFtpUrl(testCase.input, &parts)) {
            PrintFailure(L"TestParseFtpUrls", L"TryParseFtpUrl failed for " + testCase.input);
            success = false;
            continue;
        }
        if (parts.canonicalUrl != testCase.expectedCanonical) {
            PrintFailure(L"TestParseFtpUrls", L"Canonical URL mismatch for " + testCase.input);
            success = false;
        }
        if (parts.userName != testCase.expectedUser || parts.password != testCase.expectedPassword) {
            PrintFailure(L"TestParseFtpUrls", L"Credential mismatch for " + testCase.input);
            success = false;
        }
        if (parts.host != testCase.expectedHost) {
            PrintFailure(L"TestParseFtpUrls", L"Host mismatch for " + testCase.input);
            success = false;
        }
        if (parts.path != testCase.expectedPath) {
            PrintFailure(L"TestParseFtpUrls", L"Path mismatch for " + testCase.input);
            success = false;
        }
        if (parts.port != testCase.expectedPort) {
            PrintFailure(L"TestParseFtpUrls", L"Port mismatch for " + testCase.input);
            success = false;
        }
    }
    return success;
}

bool TestFtpPidlHelpers() {
    shelltabs::FtpUrlParts parts;
    parts.host = L"example.com";
    parts.userName = L"user";
    parts.password = L"pass";
    parts.path = L"/root/sub/file.txt";

    shelltabs::UniquePidl pidl = shelltabs::CreateFtpPidlFromUrl(parts);
    if (!pidl) {
        PrintFailure(L"TestFtpPidlHelpers", L"CreateFtpPidlFromUrl returned null");
        return false;
    }

    const auto* current = reinterpret_cast<const ITEMIDLIST*>(pidl.get());
    int index = 0;
    bool foundFile = false;
    while (current && current->mkid.cb) {
        const SHITEMID& item = current->mkid;
        if (!shelltabs::ftp::IsFtpItemId(item)) {
            PrintFailure(L"TestFtpPidlHelpers", L"Unexpected item signature");
            return false;
        }
        const auto type = shelltabs::ftp::GetItemType(item);
        if (index == 0) {
            if (type != shelltabs::ftp::ItemType::Root) {
                PrintFailure(L"TestFtpPidlHelpers", L"First item was not root");
                return false;
            }
            std::wstring host;
            if (!shelltabs::ftp::TryGetComponentString(item, shelltabs::ftp::ComponentType::Host, &host) || host != parts.host) {
                PrintFailure(L"TestFtpPidlHelpers", L"Root host component mismatch");
                return false;
            }
        } else {
            if (index == 1 && type != shelltabs::ftp::ItemType::Directory) {
                PrintFailure(L"TestFtpPidlHelpers", L"Expected directory component");
                return false;
            }
            if (type == shelltabs::ftp::ItemType::File) {
                foundFile = true;
            }
        }
        current = reinterpret_cast<const ITEMIDLIST*>(reinterpret_cast<const std::byte*>(current) + current->mkid.cb);
        ++index;
    }

    if (!foundFile) {
        PrintFailure(L"TestFtpPidlHelpers", L"File component was not detected");
        return false;
    }

    shelltabs::FtpUrlParts parsedParts;
    std::vector<std::wstring> segments;
    bool terminalDirectory = true;
    if (!shelltabs::ftp::TryParseFtpPidl(pidl.get(), &parsedParts, &segments, &terminalDirectory)) {
        PrintFailure(L"TestFtpPidlHelpers", L"TryParseFtpPidl failed");
        return false;
    }
    if (segments.size() != 3 || segments.back() != L"file.txt" || terminalDirectory) {
        PrintFailure(L"TestFtpPidlHelpers", L"Unexpected segments returned from TryParseFtpPidl");
        return false;
    }
    if (shelltabs::ftp::BuildUrlFromFtpPidl(pidl.get()) != parts.canonicalUrl) {
        PrintFailure(L"TestFtpPidlHelpers", L"BuildUrlFromFtpPidl mismatch");
        return false;
    }

    const UINT expectedSize = ILGetSize(pidl.get());
    const auto serialized = shelltabs::ftp::SerializeFtpPidl(pidl.get());
    if (serialized.empty() || serialized.size() != expectedSize) {
        PrintFailure(L"TestFtpPidlHelpers", L"SerializeFtpPidl size mismatch");
        return false;
    }

    return true;
}

bool TestMlsdDirectoryListing() {
    const std::string listing =
        "type=dir;modify=20231010153000;perm=el; unique=123; subdir\r\n"
        "type=file;size=1024;modify=20231011101010;perm=adfr; unique=124; sample.txt\r\n";

    std::vector<shelltabs::FtpDirectoryEntry> entries;
    const HRESULT hr = shelltabs::ftp::testhooks::ParseDirectoryListing(listing, &entries);
    if (FAILED(hr)) {
        PrintFailure(L"TestMlsdDirectoryListing", L"ParseDirectoryListing returned failure");
        return false;
    }
    if (entries.size() != 2) {
        PrintFailure(L"TestMlsdDirectoryListing", L"Unexpected number of MLSD entries");
        return false;
    }
    if (!entries[0].isDirectory || entries[0].name != L"subdir") {
        PrintFailure(L"TestMlsdDirectoryListing", L"Directory entry parsing mismatch");
        return false;
    }
    if (entries[1].isDirectory || entries[1].name != L"sample.txt" || entries[1].size != 1024) {
        PrintFailure(L"TestMlsdDirectoryListing", L"File entry parsing mismatch");
        return false;
    }
    return true;
}

bool TestLegacyDirectoryListingFallback() {
    const std::string listing = "-rw-r--r-- 1 user group 1234 Jan 01 12:34 legacy.txt\r\n";
    std::vector<shelltabs::FtpDirectoryEntry> entries;
    const HRESULT hr = shelltabs::ftp::testhooks::ParseDirectoryListing(listing, &entries);
    if (FAILED(hr) || entries.size() != 1) {
        PrintFailure(L"TestLegacyDirectoryListingFallback", L"Legacy listing did not parse");
        return false;
    }
    if (entries[0].name != L"-rw-r--r-- 1 user group 1234 Jan 01 12:34 legacy.txt" || entries[0].isDirectory) {
        PrintFailure(L"TestLegacyDirectoryListingFallback", L"Legacy entry values unexpected");
        return false;
    }
    return true;
}

}  // namespace

int wmain() {
    const HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::wcerr << L"CoInitializeEx failed: 0x" << std::hex << hr << std::endl;
        return 1;
    }

    const std::vector<TestDefinition> tests = {
        {L"TestParseFtpUrls", &TestParseFtpUrls},
        {L"TestFtpPidlHelpers", &TestFtpPidlHelpers},
        {L"TestMlsdDirectoryListing", &TestMlsdDirectoryListing},
        {L"TestLegacyDirectoryListingFallback", &TestLegacyDirectoryListingFallback},
    };

    bool success = true;
    for (const auto& test : tests) {
        if (!test.fn()) {
            std::wcerr << L"[FAILED] " << test.name << std::endl;
            success = false;
        } else {
            std::wcout << L"[PASSED] " << test.name << std::endl;
        }
    }

    CoUninitialize();

    if (!success) {
        std::wcerr << L"FTP utility tests failed." << std::endl;
        return 1;
    }

    std::wcout << L"FTP utility tests passed." << std::endl;
    return 0;
}
