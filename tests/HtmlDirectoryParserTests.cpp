#include "HtmlDirectoryParser.h"

#include <cassert>
#include <cstdio>
#include <string>

using namespace shelltabs::http;

namespace {

void TestApacheAutoindex() {
    const wchar_t* html =
        L"<html><head><title>Index of /files/</title></head>\n"
        L"<body><h1>Index of /files/</h1>\n"
        L"<pre><img src=\"/icons/blank.gif\" alt=\"Icon\">"
        L" <a href=\"?C=N;O=D\">Name</a>   <a href=\"?C=M;O=A\">Last modified</a>"
        L"   <a href=\"?C=S;O=A\">Size</a>\n"
        L"<hr><img src=\"/icons/back.gif\" alt=\"[PARENTDIR]\"> <a href=\"/\">Parent Directory</a>\n"
        L"<img src=\"/icons/folder.gif\" alt=\"[DIR]\"> <a href=\"No-Intro/\">No-Intro/</a>"
        L"                 2024-01-15 10:30    -\n"
        L"<img src=\"/icons/folder.gif\" alt=\"[DIR]\"> <a href=\"Redump/\">Redump/</a>"
        L"                   2024-02-20 14:45    -\n"
        L"<img src=\"/icons/unknown.gif\" alt=\"[   ]\"> <a href=\"readme.txt\">readme.txt</a>"
        L"              2024-03-10 08:00  1.5K\n"
        L"<img src=\"/icons/compressed.gif\" alt=\"[   ]\"> <a href=\"archive.zip\">archive.zip</a>"
        L"            2024-04-05 12:30  2.3G\n"
        L"<hr></pre></body></html>";

    ParseResult result = ParseDirectoryListing(html, L"example.com");

    assert(result.format == DirectoryFormat::Apache);
    assert(result.entries.size() == 4);

    assert(result.entries[0].name == L"No-Intro");
    assert(result.entries[0].isDirectory == true);

    assert(result.entries[1].name == L"Redump");
    assert(result.entries[1].isDirectory == true);

    assert(result.entries[2].name == L"readme.txt");
    assert(result.entries[2].isDirectory == false);

    assert(result.entries[3].name == L"archive.zip");
    assert(result.entries[3].isDirectory == false);

    wprintf(L"  PASS: TestApacheAutoindex\n");
}

void TestNginxAutoindex() {
    const wchar_t* html =
        L"<html>\n<head><title>Index of /files/</title></head>\n"
        L"<body>\n<h1>Index of /files/</h1><hr><pre>"
        L"<a href=\"../\">../</a>\n"
        L"<a href=\"Games/\">Games/</a>                                             15-Jan-2024 10:30       -\n"
        L"<a href=\"ROM.zip\">ROM.zip</a>                                           01-Mar-2024 12:00  1234567\n"
        L"</pre><hr></body>\n</html>\n";

    ParseResult result = ParseDirectoryListing(html, L"example.com");

    assert(result.entries.size() == 2);

    assert(result.entries[0].name == L"Games");
    assert(result.entries[0].isDirectory == true);

    assert(result.entries[1].name == L"ROM.zip");
    assert(result.entries[1].isDirectory == false);
    assert(result.entries[1].size == 1234567);

    wprintf(L"  PASS: TestNginxAutoindex\n");
}

void TestFilteredLinks() {
    const wchar_t* html =
        L"<html><head><title>Index of /</title></head><body>\n"
        L"<a href=\"../\">Parent</a>\n"
        L"<a href=\"?C=N;O=D\">Name</a>\n"
        L"<a href=\".\">Current</a>\n"
        L"<a href=\"javascript:void(0)\">Script</a>\n"
        L"<a href=\"https://other.com/page\">External</a>\n"
        L"<a href=\"valid/\">Valid</a>\n"
        L"</body></html>";

    ParseResult result = ParseDirectoryListing(html, L"example.com");

    assert(result.entries.size() == 1);
    assert(result.entries[0].name == L"valid");
    assert(result.entries[0].isDirectory == true);

    wprintf(L"  PASS: TestFilteredLinks\n");
}

void TestUrlDecodedNames() {
    const wchar_t* html =
        L"<html><body>\n"
        L"<a href=\"My%20Game%20(USA).zip\">My Game (USA).zip</a>\n"
        L"<a href=\"folder%20with%20spaces/\">folder with spaces/</a>\n"
        L"</body></html>";

    ParseResult result = ParseDirectoryListing(html, L"example.com");

    assert(result.entries.size() == 2);
    assert(result.entries[0].name == L"My Game (USA).zip");
    assert(result.entries[0].isDirectory == false);
    assert(result.entries[1].name == L"folder with spaces");
    assert(result.entries[1].isDirectory == true);

    wprintf(L"  PASS: TestUrlDecodedNames\n");
}

void TestEmptyPage() {
    ParseResult result = ParseDirectoryListing(L"<html><body>No files here</body></html>", L"example.com");
    assert(result.entries.empty());
    wprintf(L"  PASS: TestEmptyPage\n");
}

void TestSameHostAbsoluteLinks() {
    const wchar_t* html =
        L"<html><body>\n"
        L"<a href=\"https://example.com/files/game.zip\">game.zip</a>\n"
        L"<a href=\"https://other.com/files/hack.zip\">hack.zip</a>\n"
        L"</body></html>";

    ParseResult result = ParseDirectoryListing(html, L"example.com");

    assert(result.entries.size() == 1);
    assert(result.entries[0].name == L"game.zip");

    wprintf(L"  PASS: TestSameHostAbsoluteLinks\n");
}

}  // namespace

int main() {
    wprintf(L"Running HtmlDirectoryParser tests...\n");

    TestApacheAutoindex();
    TestNginxAutoindex();
    TestFilteredLinks();
    TestUrlDecodedNames();
    TestEmptyPage();
    TestSameHostAbsoluteLinks();

    wprintf(L"\nAll HtmlDirectoryParser tests passed!\n");
    return 0;
}
