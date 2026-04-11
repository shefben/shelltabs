#include "HtmlDirectoryParser.h"

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <string>
#include <string_view>

namespace shelltabs::http {

namespace {

bool StartsWithIgnoreCase(std::wstring_view text, std::wstring_view prefix) {
    if (text.size() < prefix.size()) {
        return false;
    }
    for (size_t i = 0; i < prefix.size(); ++i) {
        if (towlower(text[i]) != towlower(prefix[i])) {
            return false;
        }
    }
    return true;
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

std::wstring DecodeHtmlEntities(std::wstring_view text) {
    std::wstring result;
    result.reserve(text.size());
    size_t i = 0;
    while (i < text.size()) {
        if (text[i] == L'&') {
            if (StartsWithIgnoreCase(text.substr(i), L"&amp;")) {
                result.push_back(L'&');
                i += 5;
            } else if (StartsWithIgnoreCase(text.substr(i), L"&lt;")) {
                result.push_back(L'<');
                i += 4;
            } else if (StartsWithIgnoreCase(text.substr(i), L"&gt;")) {
                result.push_back(L'>');
                i += 4;
            } else if (StartsWithIgnoreCase(text.substr(i), L"&quot;")) {
                result.push_back(L'"');
                i += 6;
            } else if (StartsWithIgnoreCase(text.substr(i), L"&#39;") ||
                       StartsWithIgnoreCase(text.substr(i), L"&apos;")) {
                result.push_back(L'\'');
                i += (text.substr(i, 6).find(L';') == 4) ? 5 : 6;
            } else {
                result.push_back(text[i]);
                ++i;
            }
        } else {
            result.push_back(text[i]);
            ++i;
        }
    }
    return result;
}

std::wstring UrlDecode(std::wstring_view text) {
    auto hexVal = [](wchar_t ch) -> int {
        if (ch >= L'0' && ch <= L'9') return ch - L'0';
        if (ch >= L'a' && ch <= L'f') return ch - L'a' + 10;
        if (ch >= L'A' && ch <= L'F') return ch - L'A' + 10;
        return -1;
    };

    // Decode percent-encoded bytes into a UTF-8 byte buffer first,
    // then convert the entire buffer to a wide string. This correctly
    // handles multi-byte UTF-8 sequences like %E4%B8%AD.
    std::string utf8Bytes;
    utf8Bytes.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i) {
        if (text[i] == L'%' && i + 2 < text.size()) {
            int hiVal = hexVal(text[i + 1]);
            int loVal = hexVal(text[i + 2]);
            if (hiVal >= 0 && loVal >= 0) {
                utf8Bytes.push_back(static_cast<char>((hiVal << 4) | loVal));
                i += 2;
                continue;
            }
        }
        utf8Bytes.push_back(static_cast<char>(text[i] & 0xFF));
    }
    int needed = MultiByteToWideChar(CP_UTF8, 0, utf8Bytes.data(), static_cast<int>(utf8Bytes.size()), nullptr, 0);
    if (needed <= 0) {
        return std::wstring(text);
    }
    std::wstring result(static_cast<size_t>(needed), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8Bytes.data(), static_cast<int>(utf8Bytes.size()), result.data(), needed);
    return result;
}

struct LinkInfo {
    std::wstring href;
    std::wstring text;
    size_t endPosition = 0;  // position in HTML right after </a>
};

// Scans for <a href="...">text</a> patterns. Returns all found links.
std::vector<LinkInfo> ExtractLinks(std::wstring_view html) {
    std::vector<LinkInfo> links;
    size_t pos = 0;
    while (pos < html.size()) {
        // Find <a (case insensitive)
        size_t tagStart = std::wstring_view::npos;
        for (size_t i = pos; i + 2 < html.size(); ++i) {
            if (html[i] == L'<' && (html[i + 1] == L'a' || html[i + 1] == L'A') &&
                (html[i + 2] == L' ' || html[i + 2] == L'\t' || html[i + 2] == L'\n' || html[i + 2] == L'\r')) {
                tagStart = i;
                break;
            }
        }
        if (tagStart == std::wstring_view::npos) {
            break;
        }

        // Find href="..." or href='...'
        size_t tagEnd = html.find(L'>', tagStart);
        if (tagEnd == std::wstring_view::npos) {
            break;
        }

        std::wstring_view tagContent = html.substr(tagStart, tagEnd - tagStart + 1);
        std::wstring href;

        // Look for href=
        for (size_t i = 0; i + 5 < tagContent.size(); ++i) {
            if (StartsWithIgnoreCase(tagContent.substr(i), L"href=")) {
                size_t valueStart = i + 5;
                if (valueStart < tagContent.size()) {
                    wchar_t quote = tagContent[valueStart];
                    if (quote == L'"' || quote == L'\'') {
                        ++valueStart;
                        size_t valueEnd = tagContent.find(quote, valueStart);
                        if (valueEnd != std::wstring_view::npos) {
                            href = tagContent.substr(valueStart, valueEnd - valueStart);
                        }
                    } else {
                        // Unquoted href - take until space or >
                        size_t valueEnd = valueStart;
                        while (valueEnd < tagContent.size() && tagContent[valueEnd] != L' ' &&
                               tagContent[valueEnd] != L'>' && tagContent[valueEnd] != L'\t') {
                            ++valueEnd;
                        }
                        href = tagContent.substr(valueStart, valueEnd - valueStart);
                    }
                }
                break;
            }
        }

        if (href.empty()) {
            pos = tagEnd + 1;
            continue;
        }

        // Find the link text (between > and </a>)
        size_t textStart = tagEnd + 1;
        size_t closingTag = std::wstring_view::npos;
        for (size_t i = textStart; i + 3 < html.size(); ++i) {
            if (html[i] == L'<' && html[i + 1] == L'/' &&
                (html[i + 2] == L'a' || html[i + 2] == L'A') &&
                (html[i + 3] == L'>' || html[i + 3] == L' ')) {
                closingTag = i;
                break;
            }
        }

        std::wstring text;
        size_t endPos;
        if (closingTag != std::wstring_view::npos) {
            text = html.substr(textStart, closingTag - textStart);
            size_t closingEnd = html.find(L'>', closingTag);
            endPos = (closingEnd != std::wstring_view::npos) ? closingEnd + 1 : closingTag + 4;
        } else {
            text = html.substr(textStart, std::min<size_t>(256, html.size() - textStart));
            endPos = tagEnd + 1;
        }

        // Strip any HTML tags from the link text
        std::wstring cleanText;
        bool inTag = false;
        for (wchar_t ch : text) {
            if (ch == L'<') {
                inTag = true;
            } else if (ch == L'>') {
                inTag = false;
            } else if (!inTag) {
                cleanText.push_back(ch);
            }
        }

        // Trim whitespace
        while (!cleanText.empty() && iswspace(cleanText.front())) {
            cleanText.erase(cleanText.begin());
        }
        while (!cleanText.empty() && iswspace(cleanText.back())) {
            cleanText.pop_back();
        }

        LinkInfo info;
        info.href = DecodeHtmlEntities(href);
        info.text = DecodeHtmlEntities(cleanText);
        info.endPosition = endPos;
        links.push_back(std::move(info));
        pos = endPos;
    }
    return links;
}

bool IsFilteredLink(std::wstring_view href) {
    if (href.empty()) {
        return true;
    }
    // Filter parent directory and self links
    if (href == L"../" || href == L".." || href == L"/" || href == L"." || href == L"./") {
        return true;
    }
    // Filter absolute-path links (navigation/menu links, not directory entries).
    // Autoindex pages use relative URLs for actual entries; absolute paths like
    // /donate/ or /contact/ are site navigation and should be excluded.
    if (href.front() == L'/') {
        return true;
    }
    // Filter Apache sort query links
    if (href.find(L"?C=") != std::wstring_view::npos) {
        return true;
    }
    if (href.find(L"?c=") != std::wstring_view::npos) {
        return true;
    }
    // Filter fragment-only links
    if (href.front() == L'#') {
        return true;
    }
    // Filter javascript: links
    if (StartsWithIgnoreCase(href, L"javascript:")) {
        return true;
    }
    // Filter mailto: links
    if (StartsWithIgnoreCase(href, L"mailto:")) {
        return true;
    }
    return false;
}

bool IsExternalLink(std::wstring_view href, std::wstring_view baseHost) {
    if (StartsWithIgnoreCase(href, L"http://") || StartsWithIgnoreCase(href, L"https://")) {
        // Extract host from absolute URL
        size_t hostStart = href.find(L"//");
        if (hostStart == std::wstring_view::npos) {
            return true;
        }
        hostStart += 2;
        size_t hostEnd = href.find_first_of(L":/", hostStart);
        std::wstring_view linkHost = (hostEnd != std::wstring_view::npos) ? href.substr(hostStart, hostEnd - hostStart)
                                                                          : href.substr(hostStart);
        return !EqualsIgnoreCase(linkHost, baseHost);
    }
    return false;
}

// Try to extract file size and date from text following </a> in Apache/nginx autoindex pages.
// Apache format:  "                         2024-01-15 10:30  1.5G"
// Nginx format:   "                         01-Jan-2024 12:00        1234567"
struct MetadataResult {
    ULONGLONG size = 0;
    FILETIME lastWriteTime{};
    bool hasSize = false;
    bool hasDate = false;
};

ULONGLONG ParseSizeString(std::wstring_view token) {
    if (token.empty() || token == L"-") {
        return 0;
    }

    // Try to parse as raw number first
    double value = 0;
    size_t numEnd = 0;
    bool hasDigit = false;
    for (size_t i = 0; i < token.size(); ++i) {
        wchar_t ch = token[i];
        if ((ch >= L'0' && ch <= L'9') || ch == L'.') {
            hasDigit = true;
            numEnd = i + 1;
        } else {
            break;
        }
    }
    if (!hasDigit) {
        return 0;
    }

    std::wstring numStr(token.substr(0, numEnd));
    value = _wtof(numStr.c_str());

    std::wstring_view suffix = token.substr(numEnd);
    while (!suffix.empty() && iswspace(suffix.front())) {
        suffix.remove_prefix(1);
    }

    if (suffix.empty()) {
        return static_cast<ULONGLONG>(value);
    }

    wchar_t unit = towupper(suffix.front());
    switch (unit) {
        case L'K':
            return static_cast<ULONGLONG>(value * 1024.0);
        case L'M':
            return static_cast<ULONGLONG>(value * 1024.0 * 1024.0);
        case L'G':
            return static_cast<ULONGLONG>(value * 1024.0 * 1024.0 * 1024.0);
        case L'T':
            return static_cast<ULONGLONG>(value * 1024.0 * 1024.0 * 1024.0 * 1024.0);
        default:
            return static_cast<ULONGLONG>(value);
    }
}

// Month name to number (1-12)
int ParseMonthName(std::wstring_view name) {
    static const wchar_t* months[] = {L"jan", L"feb", L"mar", L"apr", L"may", L"jun",
                                       L"jul", L"aug", L"sep", L"oct", L"nov", L"dec"};
    for (int i = 0; i < 12; ++i) {
        if (name.size() >= 3 && StartsWithIgnoreCase(name, months[i])) {
            return i + 1;
        }
    }
    return 0;
}

bool TryParseDate(std::wstring_view text, FILETIME* ft) {
    // Try "YYYY-MM-DD HH:MM" (Apache)
    // or "DD-Mon-YYYY HH:MM" (Nginx)
    SYSTEMTIME st{};

    // Tokenize by spaces and dashes
    std::vector<std::wstring_view> tokens;
    size_t i = 0;
    while (i < text.size()) {
        while (i < text.size() && (iswspace(text[i]) || text[i] == L'-' || text[i] == L':')) {
            ++i;
        }
        size_t start = i;
        while (i < text.size() && !iswspace(text[i]) && text[i] != L'-' && text[i] != L':') {
            ++i;
        }
        if (i > start) {
            tokens.push_back(text.substr(start, i - start));
        }
    }

    if (tokens.size() >= 5) {
        // Try Apache: YYYY MM DD HH MM
        int v0 = _wtoi(std::wstring(tokens[0]).c_str());
        int v1 = _wtoi(std::wstring(tokens[1]).c_str());
        int v2 = _wtoi(std::wstring(tokens[2]).c_str());
        int v3 = _wtoi(std::wstring(tokens[3]).c_str());
        int v4 = _wtoi(std::wstring(tokens[4]).c_str());

        if (v0 >= 1970 && v0 <= 2100 && v1 >= 1 && v1 <= 12 && v2 >= 1 && v2 <= 31) {
            st.wYear = static_cast<WORD>(v0);
            st.wMonth = static_cast<WORD>(v1);
            st.wDay = static_cast<WORD>(v2);
            st.wHour = static_cast<WORD>(v3);
            st.wMinute = static_cast<WORD>(v4);
            return SystemTimeToFileTime(&st, ft) != FALSE;
        }

        // Try Nginx: DD Mon YYYY HH MM
        int month = ParseMonthName(tokens[1]);
        if (v0 >= 1 && v0 <= 31 && month > 0 && v2 >= 1970 && v2 <= 2100) {
            st.wDay = static_cast<WORD>(v0);
            st.wMonth = static_cast<WORD>(month);
            st.wYear = static_cast<WORD>(v2);
            st.wHour = static_cast<WORD>(v3);
            st.wMinute = static_cast<WORD>(v4);
            return SystemTimeToFileTime(&st, ft) != FALSE;
        }
    }

    return false;
}

MetadataResult ParseMetadataAfterLink(std::wstring_view text) {
    MetadataResult result;

    // Trim leading whitespace
    while (!text.empty() && iswspace(text.front())) {
        text.remove_prefix(1);
    }
    if (text.empty()) {
        return result;
    }

    // Split into whitespace-delimited tokens
    std::vector<std::wstring_view> tokens;
    size_t i = 0;
    while (i < text.size() && tokens.size() < 10) {
        while (i < text.size() && iswspace(text[i])) {
            ++i;
        }
        size_t start = i;
        while (i < text.size() && !iswspace(text[i])) {
            ++i;
        }
        if (i > start) {
            tokens.push_back(text.substr(start, i - start));
        }
    }

    // Look for date+time pattern and size
    // Apache: "2024-01-15 10:30  1.5G" → tokens: ["2024-01-15", "10:30", "1.5G"]
    // Nginx:  "01-Jan-2024 12:00     1234567" → tokens: ["01-Jan-2024", "12:00", "1234567"]

    for (size_t t = 0; t + 1 < tokens.size(); ++t) {
        // Check if token looks like a date (contains '-')
        if (tokens[t].find(L'-') != std::wstring_view::npos) {
            // Check if next token looks like a time (contains ':')
            if (tokens[t + 1].find(L':') != std::wstring_view::npos) {
                std::wstring dateTimeStr(tokens[t]);
                dateTimeStr.push_back(L' ');
                dateTimeStr.append(tokens[t + 1]);
                if (TryParseDate(dateTimeStr, &result.lastWriteTime)) {
                    result.hasDate = true;
                }

                // Size is typically the token after the time
                if (t + 2 < tokens.size()) {
                    ULONGLONG size = ParseSizeString(tokens[t + 2]);
                    if (size > 0 || tokens[t + 2] == L"-") {
                        result.size = size;
                        result.hasSize = true;
                    }
                }
                break;
            }
        }
    }

    // If we didn't find a date-time pattern, look for standalone size (nginx sometimes)
    if (!result.hasSize && !tokens.empty()) {
        // Last token might be the size
        ULONGLONG size = ParseSizeString(tokens.back());
        if (size > 0) {
            result.size = size;
            result.hasSize = true;
        }
    }

    return result;
}

DirectoryFormat DetectFormat(std::wstring_view html) {
    // Check for Apache autoindex markers (plain <pre> format)
    if (html.find(L"<pre>") != std::wstring_view::npos || html.find(L"<PRE>") != std::wstring_view::npos) {
        if (html.find(L"Index of") != std::wstring_view::npos || html.find(L"index of") != std::wstring_view::npos) {
            return DirectoryFormat::Apache;
        }
    }
    // Check for nginx autoindex (it typically doesn't use <pre> but has a specific date format)
    if (html.find(L"<title>Index of") != std::wstring_view::npos) {
        // Check for nginx-style dates (DD-Mon-YYYY)
        if (html.find(L"-Jan-") != std::wstring_view::npos || html.find(L"-Feb-") != std::wstring_view::npos ||
            html.find(L"-Mar-") != std::wstring_view::npos || html.find(L"-Apr-") != std::wstring_view::npos ||
            html.find(L"-May-") != std::wstring_view::npos || html.find(L"-Jun-") != std::wstring_view::npos ||
            html.find(L"-Jul-") != std::wstring_view::npos || html.find(L"-Aug-") != std::wstring_view::npos ||
            html.find(L"-Sep-") != std::wstring_view::npos || html.find(L"-Oct-") != std::wstring_view::npos ||
            html.find(L"-Nov-") != std::wstring_view::npos || html.find(L"-Dec-") != std::wstring_view::npos) {
            return DirectoryFormat::Nginx;
        }
        return DirectoryFormat::Apache;
    }
    // Check for table-based autoindex (Apache mod_autoindex with FancyIndexing, or custom themes
    // like Myrient). Detectable by Apache sort query parameters (?C=N, ?C=S, ?C=M) which are
    // standard mod_autoindex controls, combined with a <table> structure.
    if (html.find(L"<table") != std::wstring_view::npos) {
        if (html.find(L"?C=N") != std::wstring_view::npos || html.find(L"?C=S") != std::wstring_view::npos ||
            html.find(L"?C=M") != std::wstring_view::npos) {
            return DirectoryFormat::TableIndex;
        }
    }
    // Generic: has links but no recognized format
    return DirectoryFormat::Generic;
}

std::wstring_view GetTextAfterLink(std::wstring_view html, size_t endPos) {
    if (endPos >= html.size()) {
        return {};
    }
    // Find the next <a or <br or newline
    size_t lineEnd = endPos;
    while (lineEnd < html.size()) {
        if (html[lineEnd] == L'\n' || html[lineEnd] == L'\r') {
            break;
        }
        if (html[lineEnd] == L'<') {
            // Check if it's a new link or line break
            if (lineEnd + 1 < html.size()) {
                wchar_t next = towlower(html[lineEnd + 1]);
                if (next == L'a' || next == L'b' || next == L'/') {
                    break;
                }
            }
        }
        ++lineEnd;
    }
    return html.substr(endPos, lineEnd - endPos);
}

// Extract size and date text from sibling <td> cells after a link in table-based listings.
// Starting from the position after </a>, finds the enclosing </td> then reads subsequent
// <td> cells within the same <tr>.
struct TableCellData {
    std::wstring sizeText;
    std::wstring dateText;
};

TableCellData ExtractTableCellsAfterLink(std::wstring_view html, size_t endPos) {
    TableCellData result;
    if (endPos >= html.size()) {
        return result;
    }

    // Find closing </td> for the link's cell
    size_t tdClose = html.find(L"</td>", endPos);
    if (tdClose == std::wstring_view::npos) {
        return result;
    }

    size_t pos = tdClose + 5;
    int cellIndex = 0;
    while (pos < html.size() && cellIndex < 2) {
        // Stop if we hit </tr> (end of row)
        size_t trEnd = html.find(L"</tr>", pos);
        size_t nextTd = html.find(L"<td", pos);
        if (nextTd == std::wstring_view::npos) {
            break;
        }
        if (trEnd != std::wstring_view::npos && trEnd < nextTd) {
            break;
        }

        // Find the > that closes the <td tag
        size_t contentStart = html.find(L'>', nextTd);
        if (contentStart == std::wstring_view::npos) {
            break;
        }
        ++contentStart;

        // Find </td>
        size_t contentEnd = html.find(L"</td>", contentStart);
        if (contentEnd == std::wstring_view::npos) {
            break;
        }

        std::wstring cellText(html.substr(contentStart, contentEnd - contentStart));
        // Trim whitespace
        while (!cellText.empty() && iswspace(cellText.front())) {
            cellText.erase(cellText.begin());
        }
        while (!cellText.empty() && iswspace(cellText.back())) {
            cellText.pop_back();
        }

        if (cellIndex == 0) {
            result.sizeText = std::move(cellText);
        } else {
            result.dateText = std::move(cellText);
        }

        pos = contentEnd + 5;
        ++cellIndex;
    }
    return result;
}

}  // namespace

ParseResult ParseDirectoryListing(std::wstring_view html, std::wstring_view baseHost) {
    ParseResult result;
    result.format = DetectFormat(html);

    auto links = ExtractLinks(html);
    if (links.empty()) {
        return result;
    }

    for (auto& link : links) {
        if (IsFilteredLink(link.href)) {
            continue;
        }
        if (IsExternalLink(link.href, baseHost)) {
            continue;
        }

        DirectoryEntry entry;

        // Determine if it's a directory (href ends with /)
        bool isDirectory = false;
        std::wstring href = link.href;

        // Strip absolute URL prefix if it's the same host
        if (StartsWithIgnoreCase(href, L"http://") || StartsWithIgnoreCase(href, L"https://")) {
            size_t pathStart = href.find(L'/', href.find(L"//") + 2);
            if (pathStart != std::wstring::npos) {
                href = href.substr(pathStart);
            } else {
                href = L"/";
            }
        }

        if (!href.empty() && href.back() == L'/') {
            isDirectory = true;
            href.pop_back();
        }

        // Extract just the filename from the path
        size_t lastSlash = href.rfind(L'/');
        std::wstring name;
        if (lastSlash != std::wstring::npos) {
            name = href.substr(lastSlash + 1);
        } else {
            name = href;
        }

        // URL-decode the name
        name = UrlDecode(name);

        if (name.empty() || name == L"." || name == L"..") {
            continue;
        }

        entry.name = std::move(name);
        entry.isDirectory = isDirectory;

        // Try to extract metadata from text after the link (for Apache/nginx)
        if (result.format == DirectoryFormat::Apache || result.format == DirectoryFormat::Nginx) {
            std::wstring_view afterLink = GetTextAfterLink(html, link.endPosition);
            if (!afterLink.empty()) {
                MetadataResult metadata = ParseMetadataAfterLink(afterLink);
                if (metadata.hasSize) {
                    entry.size = metadata.size;
                }
                if (metadata.hasDate) {
                    entry.lastWriteTime = metadata.lastWriteTime;
                }
            }
        } else if (result.format == DirectoryFormat::TableIndex) {
            // Extract size/date from sibling <td> cells in the same table row
            TableCellData cells = ExtractTableCellsAfterLink(html, link.endPosition);
            if (!cells.sizeText.empty() && cells.sizeText != L"-") {
                entry.size = ParseSizeString(cells.sizeText);
            }
            if (!cells.dateText.empty() && cells.dateText != L"-") {
                TryParseDate(cells.dateText, &entry.lastWriteTime);
            }
        }

        result.entries.push_back(std::move(entry));
    }

    return result;
}

}  // namespace shelltabs::http
