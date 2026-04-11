#pragma once

#include <string>
#include <string_view>
#include <vector>

#include <windows.h>

namespace shelltabs::http {

struct DirectoryEntry {
    std::wstring name;
    FILETIME lastWriteTime{};
    ULONGLONG size = 0;
    bool isDirectory = false;
};

enum class DirectoryFormat {
    Unknown,
    Apache,
    Nginx,
    TableIndex,  // Table-based autoindex (Apache FancyIndexing, custom themes like Myrient)
    Generic,
};

struct ParseResult {
    std::vector<DirectoryEntry> entries;
    DirectoryFormat format = DirectoryFormat::Unknown;
};

// Parses an HTML directory listing page and extracts file/directory entries.
// |html| is the raw HTML content (wide string, already decoded from UTF-8).
// |baseHost| is the host of the page being parsed, used to filter external links.
ParseResult ParseDirectoryListing(std::wstring_view html, std::wstring_view baseHost);

}  // namespace shelltabs::http
