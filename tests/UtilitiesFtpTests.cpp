#include "Utilities.h"
#include "FtpPidl.h"

#include <windows.h>

#include <iostream>
#include <vector>

namespace {
struct FtpParseCase {
    std::wstring input;
    std::wstring expectedCanonical;
    std::wstring expectedUser;
    std::wstring expectedPassword;
    std::wstring expectedHost;
    std::wstring expectedPath;
    unsigned short expectedPort = 21;
};

bool RunCase(const FtpParseCase& testCase) {
    shelltabs::FtpUrlParts parts;
    if (!shelltabs::TryParseFtpUrl(testCase.input, &parts)) {
        std::wcerr << L"TryParseFtpUrl failed for " << testCase.input << std::endl;
        return false;
    }

    bool success = true;
    if (parts.canonicalUrl != testCase.expectedCanonical) {
        std::wcerr << L"Canonical mismatch for " << testCase.input << L"\n  expected: "
                   << testCase.expectedCanonical << L"\n  actual:   " << parts.canonicalUrl << std::endl;
        success = false;
    }
    if (parts.userName != testCase.expectedUser) {
        std::wcerr << L"User mismatch for " << testCase.input << L"\n  expected: "
                   << testCase.expectedUser << L"\n  actual:   " << parts.userName << std::endl;
        success = false;
    }
    if (parts.password != testCase.expectedPassword) {
        std::wcerr << L"Password mismatch for " << testCase.input << L"\n  expected: "
                   << testCase.expectedPassword << L"\n  actual:   " << parts.password << std::endl;
        success = false;
    }
    if (parts.host != testCase.expectedHost) {
        std::wcerr << L"Host mismatch for " << testCase.input << L"\n  expected: "
                   << testCase.expectedHost << L"\n  actual:   " << parts.host << std::endl;
        success = false;
    }
    if (parts.path != testCase.expectedPath) {
        std::wcerr << L"Path mismatch for " << testCase.input << L"\n  expected: "
                   << testCase.expectedPath << L"\n  actual:   " << parts.path << std::endl;
        success = false;
    }
    if (parts.port != testCase.expectedPort) {
        std::wcerr << L"Port mismatch for " << testCase.input << L"\n  expected: "
                   << testCase.expectedPort << L"\n  actual:   " << parts.port << std::endl;
        success = false;
    }
    return success;
}
}  // namespace

int wmain() {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::wcerr << L"CoInitializeEx failed: 0x" << std::hex << hr << std::endl;
        return 1;
    }

    std::vector<FtpParseCase> cases = {
        {L"ftp://user:pass@example.com:2121/path/to/file",
         L"ftp://user:pass@example.com:2121/path/to/file", L"user", L"pass", L"example.com",
         L"/path/to/file", 2121},
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
        success &= RunCase(testCase);
    }

    shelltabs::FtpUrlParts ftpParts;
    if (shelltabs::TryParseFtpUrl(L"ftp://user:pass@example.com:21/root/path/", &ftpParts)) {
        shelltabs::UniquePidl pidl = shelltabs::CreateFtpPidlFromUrl(ftpParts);
        shelltabs::FtpUrlParts parsedParts;
        std::vector<std::wstring> segments;
        bool isDirectory = false;
        if (!pidl || !shelltabs::ftp::TryParseFtpPidl(pidl.get(), &parsedParts, &segments, &isDirectory)) {
            std::wcerr << L"Failed to round-trip FTP PIDL" << std::endl;
            success = false;
        } else {
            if (parsedParts.host != ftpParts.host || segments.empty() || segments.front() != L"root") {
                std::wcerr << L"Round-trip mismatch" << std::endl;
                success = false;
            }
        }
    }

    CoUninitialize();

    if (!success) {
        std::wcerr << L"FTP parser tests failed." << std::endl;
        return 1;
    }

    std::wcout << L"FTP parser tests passed." << std::endl;
    return 0;
}
