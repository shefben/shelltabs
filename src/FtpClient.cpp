#include "FtpClient.h"

#include <wincred.h>
#include <winhttp.h>
#include <wininet.h>
#include <ws2tcpip.h>

// Define missing WinINet constants if not available in SDK
#ifndef INTERNET_OPTION_PASSIVE
#define INTERNET_OPTION_PASSIVE 52
#endif
#ifndef INTERNET_OPTION_TRANSFER_TYPE
#define INTERNET_OPTION_TRANSFER_TYPE 53
#endif

#include <algorithm>
#include <array>
#include <cassert>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cwchar>
#include <cwctype>
#include <errno.h>
#include <map>
#include <mutex>
#include <sstream>
#include <unordered_map>
#include <vector>

#include "Logging.h"

#pragma comment(lib, "credui.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "ws2_32.lib")

#define RETURN_IF_FAILED(hr_expr)                 \
    do {                                          \
        const HRESULT _hr_temp = (hr_expr);       \
        if (FAILED(_hr_temp)) {                   \
            return _hr_temp;                      \
        }                                         \
    } while (false)

namespace shelltabs {

namespace {

auto MakeWide(const wchar_t* value) {
    return value ? std::wstring(value) : std::wstring();
}

auto NarrowToWide(const std::string& value) {
    if (value.empty()) {
        return std::wstring();
    }
    int length = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
    if (length <= 0) {
        return std::wstring();
    }
    std::wstring result;
    result.resize(static_cast<size_t>(length));
    MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), result.data(), length);
    return result;
}

bool ParseMlsdTimestamp(std::wstring_view value, FILETIME* fileTime) {
    if (!fileTime || value.size() < 14) {
        return false;
    }
    SYSTEMTIME systemTime{};
    auto ParsePart = [](std::wstring_view part) -> WORD {
        if (part.empty()) {
            return 0;
        }
        return static_cast<WORD>(std::wcstol(std::wstring(part).c_str(), nullptr, 10));
    };

    systemTime.wYear = ParsePart(value.substr(0, 4));
    systemTime.wMonth = ParsePart(value.substr(4, 2));
    systemTime.wDay = ParsePart(value.substr(6, 2));
    systemTime.wHour = ParsePart(value.substr(8, 2));
    systemTime.wMinute = ParsePart(value.substr(10, 2));
    systemTime.wSecond = ParsePart(value.substr(12, 2));
    if (systemTime.wYear == 0 || systemTime.wMonth == 0 || systemTime.wDay == 0) {
        return false;
    }
    return SystemTimeToFileTime(&systemTime, fileTime) != FALSE;
}

bool HasWritePermission(std::wstring_view permFacts) {
    for (wchar_t ch : permFacts) {
        switch (ch) {
            case L'w':  // write file contents
            case L'm':  // create directory
            case L'a':  // append
            case L'c':  // create file
            case L'd':  // delete
            case L'f':  // rename
                return true;
        }
    }
    return false;
}

DWORD BuildAttributes(bool isDirectory, bool canWrite) {
    DWORD attributes = isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_ARCHIVE;
    if (!canWrite) {
        attributes |= FILE_ATTRIBUTE_READONLY;
    }
    return attributes;
}

bool EqualsIgnoreCase(const std::wstring& left, const std::wstring& right) {
    if (left.size() != right.size()) {
        return false;
    }
    for (size_t index = 0; index < left.size(); ++index) {
        if (towlower(left[index]) != towlower(right[index])) {
            return false;
        }
    }
    return true;
}

class ScopedCredBuffer {
public:
    ScopedCredBuffer() = default;
    ~ScopedCredBuffer() { Reset(); }

    ScopedCredBuffer(const ScopedCredBuffer&) = delete;
    ScopedCredBuffer& operator=(const ScopedCredBuffer&) = delete;

    wchar_t* Data() { return buffer_.data(); }
    size_t Size() const { return buffer_.size(); }
    void Reset() { SecureZeroMemory(buffer_.data(), buffer_.size() * sizeof(wchar_t)); }

private:
    std::array<wchar_t, CREDUI_MAX_USERNAME_LENGTH + 1> buffer_{};
};

class ScopedPasswordBuffer {
public:
    ScopedPasswordBuffer() = default;
    ~ScopedPasswordBuffer() { Reset(); }

    ScopedPasswordBuffer(const ScopedPasswordBuffer&) = delete;
    ScopedPasswordBuffer& operator=(const ScopedPasswordBuffer&) = delete;

    wchar_t* Data() { return buffer_.data(); }
    size_t Size() const { return buffer_.size(); }
    void Reset() { SecureZeroMemory(buffer_.data(), buffer_.size() * sizeof(wchar_t)); }

private:
    std::array<wchar_t, CREDUI_MAX_PASSWORD_LENGTH + 1> buffer_{};
};

class WinInetHandle {
public:
    WinInetHandle() = default;
    explicit WinInetHandle(HINTERNET handle) : handle_(handle) {}
    ~WinInetHandle() { Reset(); }

    WinInetHandle(const WinInetHandle&) = delete;
    WinInetHandle& operator=(const WinInetHandle&) = delete;

    WinInetHandle(WinInetHandle&& other) noexcept : handle_(other.handle_) { other.handle_ = nullptr; }
    WinInetHandle& operator=(WinInetHandle&& other) noexcept {
        if (this != &other) {
            Reset();
            handle_ = other.handle_;
            other.handle_ = nullptr;
        }
        return *this;
    }

    HINTERNET Get() const { return handle_; }
    explicit operator bool() const { return handle_ != nullptr; }

    void Reset(HINTERNET handle = nullptr) {
        if (handle_) {
            InternetCloseHandle(handle_);
        }
        handle_ = handle;
    }

private:
    HINTERNET handle_ = nullptr;
};

class SocketHandle {
public:
    SocketHandle() = default;
    explicit SocketHandle(SOCKET socket) : socket_(socket) {}
    ~SocketHandle() { Reset(); }

    SocketHandle(const SocketHandle&) = delete;
    SocketHandle& operator=(const SocketHandle&) = delete;

    SocketHandle(SocketHandle&& other) noexcept : socket_(other.socket_) { other.socket_ = INVALID_SOCKET; }
    SocketHandle& operator=(SocketHandle&& other) noexcept {
        if (this != &other) {
            Reset();
            socket_ = other.socket_;
            other.socket_ = INVALID_SOCKET;
        }
        return *this;
    }

    SOCKET Get() const { return socket_; }
    explicit operator bool() const { return socket_ != INVALID_SOCKET; }

    SOCKET Release() {
        SOCKET socket = socket_;
        socket_ = INVALID_SOCKET;
        return socket;
    }

    void Reset(SOCKET socket = INVALID_SOCKET) {
        if (socket_ != INVALID_SOCKET) {
            closesocket(socket_);
        }
        socket_ = socket;
    }

private:
    SOCKET socket_ = INVALID_SOCKET;
};

bool IsWinInetAvailable() {
    static const HMODULE module = LoadLibraryW(L"wininet.dll");
    return module != nullptr;
}

bool IsWinHttpAvailable() {
    static const HMODULE module = LoadLibraryW(L"winhttp.dll");
    return module != nullptr;
}

int EnsureWinsockInitialized() {
    static std::once_flag once;
    static int error = 0;
    std::call_once(once, []() {
        WSADATA data;
        error = WSAStartup(MAKEWORD(2, 2), &data);
    });
    return error;
}

std::wstring FormatHostPort(const std::wstring& host, INTERNET_PORT port) {
    std::wstringstream stream;
    stream << host;
    if (port != INTERNET_DEFAULT_FTP_PORT && port != 0) {
        stream << L":" << port;
    }
    return stream.str();
}

HRESULT HResultFromSocketError(int error) {
    if (error == 0) {
        return S_OK;
    }
    return HRESULT_FROM_WIN32(error);
}

HRESULT HResultFromFtpReply(int statusCode) {
    if (statusCode >= 200 && statusCode < 300) {
        return S_OK;
    }
    switch (statusCode) {
        case 421:
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        case 425:
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
        case 426:
            return HRESULT_FROM_WIN32(ERROR_NETNAME_DELETED);
        case 450:
        case 451:
            return HRESULT_FROM_WIN32(ERROR_BUSY);
        case 452:
            return HRESULT_FROM_WIN32(ERROR_DISK_FULL);
        case 500:
        case 501:
        case 502:
        case 503:
        case 504:
            return HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION);
        case 530:
            return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
        case 550:
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        case 551:
            return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
        case 552:
            return HRESULT_FROM_WIN32(ERROR_DISK_FULL);
        case 553:
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        default:
            return HRESULT_FROM_WIN32(ERROR_INTERNET_EXTENDED_ERROR);
    }
}

HRESULT HResultFromInternetError(DWORD error) {
    if (error == ERROR_SUCCESS) {
        return S_OK;
    }
    switch (error) {
        case ERROR_INTERNET_TIMEOUT:
            return HRESULT_FROM_WIN32(WAIT_TIMEOUT);
        case ERROR_INTERNET_NAME_NOT_RESOLVED:
            return HRESULT_FROM_WIN32(WSAHOST_NOT_FOUND);
        case ERROR_INTERNET_CANNOT_CONNECT:
        case ERROR_INTERNET_CONNECTION_ABORTED:
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
        case ERROR_INTERNET_CONNECTION_RESET:
            return HRESULT_FROM_WIN32(ERROR_NETNAME_DELETED);
        case ERROR_INTERNET_LOGIN_FAILURE:
            return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
        case ERROR_INTERNET_EXTENDED_ERROR:
            return HRESULT_FROM_WIN32(ERROR_INTERNET_EXTENDED_ERROR);
        default:
            return HRESULT_FROM_WIN32(error);
    }
}

DWORD ParsePassivePort(const std::wstring& response) {
    size_t start = response.find(L"(");
    size_t end = response.find(L")", start != std::wstring::npos ? start : 0);
    if (start == std::wstring::npos || end == std::wstring::npos || end <= start + 1) {
        return 0;
    }
    std::wstring inside = response.substr(start + 1, end - start - 1);
    std::vector<int> parts;
    std::wstringstream ss(inside);
    while (ss.good()) {
        int value = 0;
        wchar_t comma = L'\0';
        ss >> value;
        if (ss.fail()) {
            break;
        }
        parts.push_back(value);
        if (!(ss >> comma)) {
            break;
        }
        if (comma != L',') {
            break;
        }
    }
    if (parts.size() < 6) {
        return 0;
    }
    return static_cast<DWORD>((parts[4] << 8) | parts[5]);
}

DWORD ParseExtendedPassivePort(const std::wstring& response) {
    size_t start = response.find(L"(");
    size_t end = response.find(L")", start != std::wstring::npos ? start : 0);
    if (start == std::wstring::npos || end == std::wstring::npos || end <= start + 1) {
        return 0;
    }
    std::wstring inside = response.substr(start + 1, end - start - 1);
    size_t lastPipe = inside.find_last_of(L'|');
    if (lastPipe == std::wstring::npos) {
        return 0;
    }
    std::wstring portString = inside.substr(lastPipe + 1);
    return static_cast<DWORD>(wcstoul(portString.c_str(), nullptr, 10));
}

struct SocketResponse {
    int status = 0;
    std::wstring message;
};

HRESULT ReadSocketLine(SOCKET socket, SocketResponse* response, std::string* rawBuffer) {
    if (!response) {
        return E_POINTER;
    }
    std::string buffer;
    buffer.reserve(512);
    char chunk[256];
    while (true) {
        int received = recv(socket, chunk, sizeof(chunk), 0);
        if (received == 0) {
            break;
        }
        if (received == SOCKET_ERROR) {
            return HResultFromSocketError(WSAGetLastError());
        }
        buffer.append(chunk, static_cast<size_t>(received));
        if (buffer.size() >= 2 && buffer[buffer.size() - 2] == '\r' && buffer.back() == '\n') {
            break;
        }
    }
    if (rawBuffer) {
        *rawBuffer = buffer;
    }
    if (buffer.empty()) {
        return HRESULT_FROM_WIN32(ERROR_CONNECTION_ABORTED);
    }
    if (buffer.size() < 4) {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    int status = 0;
    for (size_t i = 0; i < 3 && i < buffer.size(); ++i) {
        if (buffer[i] < '0' || buffer[i] > '9') {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        status = status * 10 + (buffer[i] - '0');
    }
    response->status = status;
    response->message = NarrowToWide(buffer.substr(4));
    if (!response->message.empty() && response->message.back() == L'\n') {
        response->message.pop_back();
    }
    if (!response->message.empty() && response->message.back() == L'\r') {
        response->message.pop_back();
    }
    return S_OK;
}

HRESULT SendSocketCommand(SOCKET socket, const std::wstring& command) {
    std::string narrow;
    int required = WideCharToMultiByte(CP_UTF8, 0, command.c_str(), static_cast<int>(command.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    narrow.resize(static_cast<size_t>(required));
    WideCharToMultiByte(CP_UTF8, 0, command.c_str(), static_cast<int>(command.size()), narrow.data(), required, nullptr, nullptr);
    narrow.append("\r\n");
    const char* data = narrow.data();
    size_t remaining = narrow.size();
    while (remaining > 0) {
        int sent = send(socket, data, static_cast<int>(remaining), 0);
        if (sent == SOCKET_ERROR) {
            return HResultFromSocketError(WSAGetLastError());
        }
        remaining -= static_cast<size_t>(sent);
        data += sent;
    }
    return S_OK;
}

std::wstring NormalizeRemotePath(const std::wstring& path) {
    if (path.empty()) {
        return L"";
    }
    std::wstring normalized = path;
    std::replace(normalized.begin(), normalized.end(), L'\\', L'/');
    return normalized;
}

HRESULT ExtractDirectoryListing(const std::string& raw, std::vector<FtpDirectoryEntry>* entries) {
    if (!entries) {
        return E_POINTER;
    }
    entries->clear();
    std::wistringstream stream(NarrowToWide(raw));
    std::wstring line;
    while (std::getline(stream, line)) {
        if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
        }
        if (line.empty()) {
            continue;
        }

        size_t separator = line.find(L' ');
        if (separator == std::wstring::npos) {
            // Fallback for unexpected formats; keep legacy behavior.
            FtpDirectoryEntry entry;
            entry.name = line;
            entry.isDirectory = false;
            entry.attributes = FILE_ATTRIBUTE_ARCHIVE;
            entries->push_back(std::move(entry));
            continue;
        }

        std::wstring facts = line.substr(0, separator);
        std::wstring name = line.substr(separator + 1);
        size_t firstChar = name.find_first_not_of(L" ");
        if (firstChar != std::wstring::npos) {
            name.erase(0, firstChar);
        }
        if (name.empty() || name == L"." || name == L"..") {
            continue;
        }

        bool isDirectory = false;
        bool canWrite = true;
        ULONGLONG size = 0;
        FILETIME modified{};

        size_t position = 0;
        while (position < facts.size()) {
            size_t next = facts.find(L';', position);
            std::wstring_view fact(&facts[position], (next == std::wstring::npos ? facts.size() : next) - position);
            size_t equals = fact.find(L'=');
            if (equals != std::wstring_view::npos) {
                std::wstring_view key = fact.substr(0, equals);
                std::wstring_view value = fact.substr(equals + 1);
                if (key == L"type") {
                    if (value == L"dir" || value == L"cdir") {
                        isDirectory = true;
                    } else if (value == L"pdir") {
                        // Parent directory entry should be skipped.
                        name.clear();
                        break;
                    } else {
                        isDirectory = false;
                    }
                } else if (key == L"size") {
                    size = _wcstoui64(std::wstring(value).c_str(), nullptr, 10);
                } else if (key == L"modify") {
                    ParseMlsdTimestamp(value, &modified);
                } else if (key == L"perm") {
                    canWrite = HasWritePermission(value);
                }
            }
            if (next == std::wstring::npos) {
                break;
            }
            position = next + 1;
        }

        if (name.empty()) {
            continue;
        }

        FtpDirectoryEntry entry;
        entry.name = std::move(name);
        entry.isDirectory = isDirectory;
        entry.size = size;
        entry.lastWriteTime = modified;
        entry.attributes = BuildAttributes(isDirectory, canWrite);
        entries->push_back(std::move(entry));
    }
    return S_OK;
}

}  // namespace

class FtpClient::FtpCommandStateMachine {
public:
    enum class Command {
        None,
        SendUser,
        SendPassword,
        SetBinaryMode,
        ChangeDirectory,
        EnterPassive,
        List,
        Mlsd,
        Retrieve,
        Store,
        Complete,
    };

    struct Step {
        Command command = Command::None;
        std::wstring argument;
    };

    FtpCommandStateMachine(const FtpConnectionOptions& options, const FtpCredential& credentials,
                           const FtpOperationContext& context)
        : options_(options), credentials_(credentials), context_(context) {
        BuildSteps();
    }

    bool HasNext() const { return currentIndex_ < steps_.size(); }

    Step Next() {
        if (!HasNext()) {
            return {};
        }
        return steps_[currentIndex_++];
    }

    void ResetForRetry() { currentIndex_ = 0; }

private:
    void BuildSteps() {
        steps_.clear();
        steps_.push_back({Command::SendUser, credentials_.userName});
        steps_.push_back({Command::SendPassword, credentials_.password});
        steps_.push_back({Command::SetBinaryMode, L""});
        if (!options_.initialPath.empty()) {
            steps_.push_back({Command::ChangeDirectory, NormalizeRemotePath(options_.initialPath)});
        }
        if (!context_.remotePath.empty() && context_.kind == FtpOperationKind::DirectoryListing) {
            steps_.push_back({Command::ChangeDirectory, NormalizeRemotePath(context_.remotePath)});
        }
        if (options_.passiveMode) {
            steps_.push_back({Command::EnterPassive, L""});
        }
        switch (context_.kind) {
            case FtpOperationKind::DirectoryListing:
                if (context_.useMlsd && options_.preferMlsd) {
                    steps_.push_back({Command::Mlsd, L""});
                } else {
                    steps_.push_back({Command::List, L""});
                }
                break;
            case FtpOperationKind::Download:
                steps_.push_back({Command::Retrieve, NormalizeRemotePath(context_.remotePath)});
                break;
            case FtpOperationKind::Upload:
                steps_.push_back({Command::Store, NormalizeRemotePath(context_.remotePath)});
                break;
        }
        steps_.push_back({Command::Complete, L""});
    }

    const FtpConnectionOptions& options_;
    const FtpCredential& credentials_;
    const FtpOperationContext& context_;
    std::vector<Step> steps_;
    size_t currentIndex_ = 0;
};

class FtpClient::FtpTransportSession {
public:
    virtual ~FtpTransportSession() = default;
    virtual HRESULT Open(const FtpConnectionOptions& options, const FtpCredential& credentials) = 0;
    virtual HRESULT SendUser(const std::wstring& userName) = 0;
    virtual HRESULT SendPassword(const std::wstring& password) = 0;
    virtual HRESULT SetBinaryMode() = 0;
    virtual HRESULT ChangeDirectory(const std::wstring& remotePath) = 0;
    virtual HRESULT EnterPassiveMode(bool preferExtended, DWORD* portOut) = 0;
    virtual HRESULT ListDirectory(bool useMlsd, std::vector<FtpDirectoryEntry>* entries) = 0;
    virtual HRESULT RetrieveFile(const std::wstring& remotePath, const std::wstring& localPath,
                                 FtpTransferResult* result) = 0;
    virtual HRESULT StoreFile(const std::wstring& localPath, const std::wstring& remotePath,
                              FtpTransferResult* result) = 0;
    virtual HRESULT CompleteOperation() = 0;
    virtual bool IsAlive() const = 0;
};

class WinInetFtpSession : public FtpClient::FtpTransportSession {
public:
    WinInetFtpSession() = default;
    ~WinInetFtpSession() override { Close(); }

    HRESULT Open(const FtpConnectionOptions& options, const FtpCredential& credentials) override {
        options_ = options;
        credentials_ = credentials;
        if (!internet_) {
            internet_.Reset(InternetOpenW(L"ShellTabs", INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0));
            if (!internet_) {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
        if (!connection_) {
            connection_.Reset(InternetConnectW(internet_.Get(), options.host.c_str(), options.port, credentials.userName.c_str(),
                                               credentials.password.c_str(), INTERNET_SERVICE_FTP, 0, 0));
            if (!connection_) {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
        DWORD passive = options.passiveMode ? 1u : 0u;
        if (!InternetSetOptionW(connection_.Get(), INTERNET_OPTION_PASSIVE, &passive, sizeof(passive))) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }

    HRESULT SendUser(const std::wstring&) override { return S_OK; }

    HRESULT SendPassword(const std::wstring&) override { return S_OK; }

    HRESULT SetBinaryMode() override {
        DWORD type = FTP_TRANSFER_TYPE_BINARY;
        if (!InternetSetOptionW(connection_.Get(), INTERNET_OPTION_TRANSFER_TYPE, &type, sizeof(type))) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }

    HRESULT ChangeDirectory(const std::wstring& remotePath) override {
        if (remotePath.empty()) {
            return S_OK;
        }
        if (!FtpSetCurrentDirectoryW(connection_.Get(), remotePath.c_str())) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }

    HRESULT EnterPassiveMode(bool preferExtended, DWORD* portOut) override {
        if (portOut) {
            *portOut = 0;
        }
        if (!preferExtended) {
            return S_OK;
        }
        // WinInet handles passive mode internally.
        return S_OK;
    }

    HRESULT ListDirectory(bool useMlsd, std::vector<FtpDirectoryEntry>* entries) override {
        if (!entries) {
            return E_POINTER;
        }
        entries->clear();
        if (useMlsd) {
            HINTERNET commandHandle = nullptr;
            if (FtpCommandW(connection_.Get(), TRUE, FTP_TRANSFER_TYPE_ASCII, L"MLSD", 0, &commandHandle)) {
                WinInetHandle handle(commandHandle);
                if (!handle) {
                    return HRESULT_FROM_WIN32(GetLastError());
                }
                std::string buffer;
                buffer.resize(4096);
                std::string aggregate;
                DWORD read = 0;
                while (InternetReadFile(handle.Get(), buffer.data(), static_cast<DWORD>(buffer.size()), &read) && read != 0) {
                    aggregate.append(buffer.data(), read);
                }
                return ExtractDirectoryListing(aggregate, entries);
            } else {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
        WIN32_FIND_DATAW findData;
        ZeroMemory(&findData, sizeof(findData));
        WinInetHandle handle(FtpFindFirstFileW(connection_.Get(), L"*", &findData, INTERNET_FLAG_RELOAD, 0));
        if (!handle) {
            DWORD error = GetLastError();
            if (error == ERROR_NO_MORE_FILES) {
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error);
        }
        do {
            if (wcscmp(findData.cFileName, L".") == 0 || wcscmp(findData.cFileName, L"..") == 0) {
                continue;
            }
            FtpDirectoryEntry entry;
            entry.name = findData.cFileName;
            entry.lastWriteTime = findData.ftLastWriteTime;
            entry.size = (static_cast<ULONGLONG>(findData.nFileSizeHigh) << 32) | findData.nFileSizeLow;
            entry.isDirectory = (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            entry.attributes = findData.dwFileAttributes;
            entries->push_back(std::move(entry));
        } while (InternetFindNextFileW(handle.Get(), &findData));
        return S_OK;
    }

    HRESULT RetrieveFile(const std::wstring& remotePath, const std::wstring& localPath,
                         FtpTransferResult* result) override {
        if (!FtpGetFileW(connection_.Get(), remotePath.c_str(), localPath.c_str(), FALSE, FILE_ATTRIBUTE_NORMAL,
                         FTP_TRANSFER_TYPE_BINARY, 0)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (result) {
            WIN32_FILE_ATTRIBUTE_DATA data;
            if (GetFileAttributesExW(localPath.c_str(), GetFileExInfoStandard, &data)) {
                result->bytesTransferred = (static_cast<ULONGLONG>(data.nFileSizeHigh) << 32) | data.nFileSizeLow;
            }
        }
        return S_OK;
    }

    HRESULT StoreFile(const std::wstring& localPath, const std::wstring& remotePath,
                      FtpTransferResult* result) override {
        if (!FtpPutFileW(connection_.Get(), localPath.c_str(), remotePath.c_str(), FTP_TRANSFER_TYPE_BINARY, 0)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (result) {
            WIN32_FILE_ATTRIBUTE_DATA data;
            if (GetFileAttributesExW(localPath.c_str(), GetFileExInfoStandard, &data)) {
                result->bytesTransferred = (static_cast<ULONGLONG>(data.nFileSizeHigh) << 32) | data.nFileSizeLow;
            }
        }
        return S_OK;
    }

    HRESULT CompleteOperation() override { return S_OK; }

    bool IsAlive() const override { return connection_.Get() != nullptr; }

private:
    void Close() {
        connection_.Reset();
        internet_.Reset();
    }

    FtpConnectionOptions options_;
    FtpCredential credentials_;
    WinInetHandle internet_;
    WinInetHandle connection_;
};

class WinHttpFtpSession : public FtpClient::FtpTransportSession {
public:
    WinHttpFtpSession() = default;
    ~WinHttpFtpSession() override { Close(); }

    HRESULT Open(const FtpConnectionOptions& options, const FtpCredential&) override {
        UNREFERENCED_PARAMETER(options);
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT SendUser(const std::wstring&) override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    HRESULT SendPassword(const std::wstring&) override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    HRESULT SetBinaryMode() override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    HRESULT ChangeDirectory(const std::wstring&) override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    HRESULT EnterPassiveMode(bool, DWORD*) override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    HRESULT ListDirectory(bool, std::vector<FtpDirectoryEntry>*) override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    HRESULT RetrieveFile(const std::wstring&, const std::wstring&, FtpTransferResult*) override {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    HRESULT StoreFile(const std::wstring&, const std::wstring&, FtpTransferResult*) override {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    HRESULT CompleteOperation() override { return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED); }
    bool IsAlive() const override { return false; }

private:
    void Close() {}
};

class SocketFtpSession : public FtpClient::FtpTransportSession {
public:
    SocketFtpSession() = default;
    ~SocketFtpSession() override { Close(); }

    HRESULT Open(const FtpConnectionOptions& options, const FtpCredential& credentials) override {
        options_ = options;
        credentials_ = credentials;
        int winsockError = EnsureWinsockInitialized();
        if (winsockError != 0) {
            return HResultFromSocketError(winsockError);
        }
        hostName_ = options.host;
        controlSocket_.Reset();
        struct addrinfoW hints = {};
        hints.ai_family = AF_UNSPEC;
        hints.ai_socktype = SOCK_STREAM;
        hints.ai_protocol = IPPROTO_TCP;
        std::wstring port = std::to_wstring(options.port);
        PADDRINFOW result = nullptr;
        int error = GetAddrInfoW(options.host.c_str(), port.c_str(), &hints, &result);
        if (error != 0) {
            return HResultFromSocketError(error);
        }
        SocketHandle connected;
        for (PADDRINFOW current = result; current != nullptr; current = current->ai_next) {
            SOCKET candidate = ::socket(current->ai_family, current->ai_socktype, current->ai_protocol);
            if (candidate == INVALID_SOCKET) {
                continue;
            }
            if (connect(candidate, current->ai_addr, static_cast<int>(current->ai_addrlen)) == 0) {
                connected.Reset(candidate);
                break;
            }
            closesocket(candidate);
        }
        FreeAddrInfoW(result);
        if (!connected) {
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
        }
        controlSocket_ = std::move(connected);
        SocketResponse response;
        RETURN_IF_FAILED(ReadResponse(response));
        RETURN_IF_FAILED(Login(credentials));
        return S_OK;
    }

    HRESULT SendUser(const std::wstring& userName) override {
        currentUser_ = userName;
        return S_OK;
    }

    HRESULT SendPassword(const std::wstring& password) override {
        currentPassword_ = password;
        return S_OK;
    }

    HRESULT SetBinaryMode() override {
        return ExecuteSimpleCommand(L"TYPE I");
    }

    HRESULT ChangeDirectory(const std::wstring& remotePath) override {
        if (remotePath.empty()) {
            return S_OK;
        }
        std::wstring command = L"CWD " + remotePath;
        return ExecuteSimpleCommand(command);
    }

    HRESULT EnterPassiveMode(bool preferExtended, DWORD* portOut) override {
        if (portOut) {
            *portOut = 0;
        }
        std::wstring command = preferExtended ? L"EPSV" : L"PASV";
        SocketResponse response;
        RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), command));
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status >= 200 && response.status < 300) {
            DWORD port = preferExtended ? ParseExtendedPassivePort(response.message) : ParsePassivePort(response.message);
            if (portOut) {
                *portOut = port;
            }
            passivePort_ = port;
            return S_OK;
        }
        if (preferExtended && response.status >= 500) {
            return EnterPassiveMode(false, portOut);
        }
        return HResultFromFtpReply(response.status);
    }

    HRESULT ListDirectory(bool useMlsd, std::vector<FtpDirectoryEntry>* entries) override {
        if (!entries) {
            return E_POINTER;
        }
        entries->clear();
        SocketHandle dataSocket;
        RETURN_IF_FAILED(OpenDataSocket(&dataSocket));
        std::wstring command = useMlsd ? L"MLSD" : L"LIST";
        SocketResponse response;
        RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), command));
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status != 150 && response.status != 125) {
            return HResultFromFtpReply(response.status);
        }
        std::string buffer;
        char chunk[4096];
        while (true) {
            int received = recv(dataSocket.Get(), chunk, sizeof(chunk), 0);
            if (received == 0) {
                break;
            }
            if (received == SOCKET_ERROR) {
                return HResultFromSocketError(WSAGetLastError());
            }
            buffer.append(chunk, static_cast<size_t>(received));
        }
        dataSocket.Reset();
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status != 226 && response.status != 250) {
            return HResultFromFtpReply(response.status);
        }
        return ExtractDirectoryListing(buffer, entries);
    }

    HRESULT RetrieveFile(const std::wstring& remotePath, const std::wstring& localPath,
                         FtpTransferResult* result) override {
        SocketHandle dataSocket;
        RETURN_IF_FAILED(OpenDataSocket(&dataSocket));
        std::wstring command = L"RETR " + remotePath;
        SocketResponse response;
        RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), command));
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status != 150 && response.status != 125) {
            return HResultFromFtpReply(response.status);
        }
        FILE* file = nullptr;
        const errno_t openResult = _wfopen_s(&file, localPath.c_str(), L"wb");
        if (openResult != 0 || !file) {
            int lastDosError = 0;
            _get_doserrno(&lastDosError);
            const DWORD error = lastDosError != 0 ? static_cast<DWORD>(lastDosError) : ERROR_OPEN_FAILED;
            return HRESULT_FROM_WIN32(error);
        }
        std::unique_ptr<FILE, decltype(&fclose)> fileCloser(file, &fclose);
        char chunk[4096];
        ULONGLONG total = 0;
        while (true) {
            int received = recv(dataSocket.Get(), chunk, sizeof(chunk), 0);
            if (received == 0) {
                break;
            }
            if (received == SOCKET_ERROR) {
                return HResultFromSocketError(WSAGetLastError());
            }
            size_t written = fwrite(chunk, 1, static_cast<size_t>(received), file);
            if (written != static_cast<size_t>(received)) {
                return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
            }
            total += static_cast<ULONGLONG>(received);
        }
        dataSocket.Reset();
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status != 226 && response.status != 250) {
            return HResultFromFtpReply(response.status);
        }
        if (result) {
            result->bytesTransferred = total;
        }
        return S_OK;
    }

    HRESULT StoreFile(const std::wstring& localPath, const std::wstring& remotePath,
                      FtpTransferResult* result) override {
        SocketHandle dataSocket;
        RETURN_IF_FAILED(OpenDataSocket(&dataSocket));
        std::wstring command = L"STOR " + remotePath;
        SocketResponse response;
        RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), command));
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status != 150 && response.status != 125) {
            return HResultFromFtpReply(response.status);
        }
        FILE* file = nullptr;
        const errno_t openResult = _wfopen_s(&file, localPath.c_str(), L"rb");
        if (openResult != 0 || !file) {
            int lastDosError = 0;
            _get_doserrno(&lastDosError);
            const DWORD error = lastDosError != 0 ? static_cast<DWORD>(lastDosError) : ERROR_OPEN_FAILED;
            return HRESULT_FROM_WIN32(error);
        }
        std::unique_ptr<FILE, decltype(&fclose)> fileCloser(file, &fclose);
        char chunk[4096];
        ULONGLONG total = 0;
        while (true) {
            size_t read = fread(chunk, 1, sizeof(chunk), file);
            if (read == 0) {
                if (ferror(file)) {
                    return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
                }
                break;
            }
            int sent = send(dataSocket.Get(), chunk, static_cast<int>(read), 0);
            if (sent == SOCKET_ERROR) {
                return HResultFromSocketError(WSAGetLastError());
            }
            total += static_cast<ULONGLONG>(sent);
        }
        dataSocket.Reset();
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status != 226 && response.status != 250) {
            return HResultFromFtpReply(response.status);
        }
        if (result) {
            result->bytesTransferred = total;
        }
        return S_OK;
    }

    HRESULT CompleteOperation() override { return S_OK; }

    bool IsAlive() const override { return controlSocket_.Get() != INVALID_SOCKET; }

private:
    HRESULT Login(const FtpCredential& credentials) {
        SocketResponse response;
        RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), L"USER " + credentials.userName));
        RETURN_IF_FAILED(ReadResponse(response));
        if (response.status == 331) {
            RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), L"PASS " + credentials.password));
            RETURN_IF_FAILED(ReadResponse(response));
        }
        if (response.status >= 200 && response.status < 300) {
            return S_OK;
        }
        return HResultFromFtpReply(response.status);
    }

    HRESULT ExecuteSimpleCommand(const std::wstring& command) {
        SocketResponse response;
        RETURN_IF_FAILED(SendSocketCommand(controlSocket_.Get(), command));
        RETURN_IF_FAILED(ReadResponse(response));
        return HResultFromFtpReply(response.status);
    }

    HRESULT ReadResponse(SocketResponse& response) {
        std::string raw;
        return ReadSocketLine(controlSocket_.Get(), &response, &raw);
    }

    HRESULT OpenDataSocket(SocketHandle* socket) {
        if (!socket) {
            return E_POINTER;
        }
        if (passivePort_ == 0) {
            RETURN_IF_FAILED(EnterPassiveMode(true, &passivePort_));
        }
        if (passivePort_ == 0) {
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
        }
        struct addrinfoW hints = {};
        hints.ai_family = AF_UNSPEC;
        hints.ai_socktype = SOCK_STREAM;
        hints.ai_protocol = IPPROTO_TCP;
        std::wstring port = std::to_wstring(passivePort_);
        PADDRINFOW result = nullptr;
        int error = GetAddrInfoW(hostName_.c_str(), port.c_str(), &hints, &result);
        if (error != 0) {
            return HResultFromSocketError(error);
        }
        SocketHandle connected;
        for (PADDRINFOW current = result; current != nullptr; current = current->ai_next) {
            SOCKET candidate = ::socket(current->ai_family, current->ai_socktype, current->ai_protocol);
            if (candidate == INVALID_SOCKET) {
                continue;
            }
            if (connect(candidate, current->ai_addr, static_cast<int>(current->ai_addrlen)) == 0) {
                connected.Reset(candidate);
                break;
            }
            closesocket(candidate);
        }
        FreeAddrInfoW(result);
        if (!connected) {
            return HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
        }
        *socket = std::move(connected);
        passivePort_ = 0;
        return S_OK;
    }

    void Close() { controlSocket_.Reset(); }

    FtpConnectionOptions options_;
    FtpCredential credentials_;
    std::wstring hostName_;
    std::wstring currentUser_;
    std::wstring currentPassword_;
    DWORD passivePort_ = 0;
    SocketHandle controlSocket_;
};

struct FtpClient::Impl {
    std::mutex mutex;
    std::unordered_map<SessionKey, std::vector<PooledSession>, SessionKeyHasher> pool;
};

FtpClient::FtpClient() : impl_(std::make_unique<Impl>()) {}
FtpClient::~FtpClient() = default;

bool FtpClient::SessionKey::operator==(const SessionKey& other) const noexcept {
    return passiveMode == other.passiveMode && useExplicitFtps == other.useExplicitFtps && port == other.port &&
           transport == other.transport && EqualsIgnoreCase(host, other.host);
}

size_t FtpClient::SessionKeyHasher::operator()(const SessionKey& key) const noexcept {
    std::wstring lowerHost = key.host;
    std::transform(lowerHost.begin(), lowerHost.end(), lowerHost.begin(), towlower);
    std::hash<std::wstring> hasher;
    size_t value = hasher(lowerHost);
    value ^= static_cast<size_t>(key.port) << 1;
    value ^= static_cast<size_t>(key.useExplicitFtps) << 2;
    value ^= static_cast<size_t>(key.passiveMode) << 3;
    value ^= static_cast<size_t>(key.transport) << 4;
    return value;
}

HRESULT FtpClient::ListDirectory(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                                 std::vector<FtpDirectoryEntry>* results, HWND credentialParent) {
    if (!results) {
        return E_POINTER;
    }
    FtpOperationContext context;
    context.kind = FtpOperationKind::DirectoryListing;
    context.remotePath.clear();
    context.useMlsd = options.preferMlsd;
    context.directoryResults = results;
    return ExecuteOperation(options, explicitCredential, credentialParent, &context);
}

HRESULT FtpClient::DownloadFile(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                                const std::wstring& remotePath, const std::wstring& localPath,
                                FtpTransferResult* transferResult, HWND credentialParent) {
    FtpOperationContext context;
    context.kind = FtpOperationKind::Download;
    context.remotePath = remotePath;
    context.localPath = localPath;
    context.transferResult = transferResult;
    return ExecuteOperation(options, explicitCredential, credentialParent, &context);
}

HRESULT FtpClient::UploadFile(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                              const std::wstring& localPath, const std::wstring& remotePath,
                              FtpTransferResult* transferResult, HWND credentialParent) {
    FtpOperationContext context;
    context.kind = FtpOperationKind::Upload;
    context.remotePath = remotePath;
    context.localPath = localPath;
    context.transferResult = transferResult;
    return ExecuteOperation(options, explicitCredential, credentialParent, &context);
}

void FtpClient::ClearConnectionPool() {
    std::lock_guard<std::mutex> guard(impl_->mutex);
    impl_->pool.clear();
}

std::wstring FtpClient::BuildCommandDescription(FtpOperationContext& context) const {
    switch (context.kind) {
        case FtpOperationKind::DirectoryListing:
            return L"directory listing";
        case FtpOperationKind::Download:
            return L"download";
        case FtpOperationKind::Upload:
            return L"upload";
    }
    return L"operation";
}

HRESULT FtpClient::ExecuteOperation(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                                    HWND credentialParent, FtpOperationContext* context) {
    if (!context) {
        return E_POINTER;
    }
    HRESULT lastError = E_FAIL;
    auto transports = BuildTransportPriority(options);
    for (auto transport : transports) {
        SessionKey key;
        key.host = options.host;
        key.port = options.port;
        key.useExplicitFtps = options.useExplicitFtps;
        key.passiveMode = options.passiveMode;
        key.transport = transport;

        for (unsigned int attempt = 0; attempt < std::max(1u, options.maxRetries); ++attempt) {
            FtpCredential credentials;
            bool persist = false;
            if (explicitCredential) {
                credentials = *explicitCredential;
            } else {
                RETURN_IF_FAILED(AcquireCredentials(options, credentialParent, &credentials, &persist));
            }
            HRESULT openResult = S_OK;
            auto session = AcquireSession(key, options, credentials, &openResult);
            if (!session) {
                lastError = openResult;
                if (FAILED(openResult) && openResult == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)) {
                    break;
                }
                continue;
            }
            FtpCommandStateMachine machine(options, credentials, *context);
            ExecuteResult result = ExecuteStateMachine(*session, machine, options, credentials, *context);
            ReleaseSession(key, session, SUCCEEDED(result.hr));
            if (SUCCEEDED(result.hr)) {
                if (persist && !credentials.password.empty()) {
                    PersistCredentials(options, credentials);
                }
                return result.hr;
            }
            lastError = result.hr;
            if (!result.retryable) {
                if (!explicitCredential && persist) {
                    RemovePersistedCredentials(options);
                }
                break;
            }
            if (options.retryDelayMilliseconds > 0) {
                Sleep(options.retryDelayMilliseconds);
            }
        }
        if (SUCCEEDED(lastError)) {
            break;
        }
    }
    return lastError;
}

HRESULT FtpClient::AcquireCredentials(const FtpConnectionOptions& options, HWND credentialParent,
                                      FtpCredential* credentials, bool* shouldPersist) {
    if (!credentials || !shouldPersist) {
        return E_POINTER;
    }
    *shouldPersist = false;
    credentials->userName.clear();
    credentials->password.clear();

    const std::wstring target = BuildCredentialTargetName(options);
    CREDENTIALW* stored = nullptr;
    if (!options.alwaysPromptForCredentials && options.allowCredentialPersistence &&
        CredReadW(target.c_str(), CRED_TYPE_GENERIC, 0, &stored)) {
        credentials->userName = MakeWide(stored->UserName);
        credentials->password.assign(reinterpret_cast<const wchar_t*>(stored->CredentialBlob),
                                      stored->CredentialBlobSize / sizeof(wchar_t));
        credentials->persisted = true;
        CredFree(stored);
        return S_OK;
    }

    ScopedCredBuffer userName;
    ScopedPasswordBuffer password;
    BOOL save = options.allowCredentialPersistence ? TRUE : FALSE;
    CREDUI_INFOW info = {};
    info.cbSize = sizeof(info);
    info.hwndParent = credentialParent;
    info.pszMessageText = L"Enter FTP credentials";
    info.pszCaptionText = L"ShellTabs";
    const std::wstring server = FormatHostPort(options.host, options.port);
    DWORD status = CredUIPromptForCredentialsW(&info, server.c_str(), nullptr, 0, userName.Data(),
                                              static_cast<ULONG>(userName.Size()), password.Data(),
                                              static_cast<ULONG>(password.Size()), &save,
                                              CREDUI_FLAGS_GENERIC_CREDENTIALS |
                                                  (options.allowCredentialPersistence ? 0 : CREDUI_FLAGS_DO_NOT_PERSIST) |
                                                  CREDUI_FLAGS_ALWAYS_SHOW_UI |
                                                  (options.alwaysPromptForCredentials ? CREDUI_FLAGS_EXCLUDE_CERTIFICATES : 0));
    if (status != NO_ERROR) {
        return HRESULT_FROM_WIN32(status);
    }
    credentials->userName.assign(userName.Data());
    credentials->password.assign(password.Data());
    credentials->persisted = save == TRUE;
    *shouldPersist = credentials->persisted && options.allowCredentialPersistence;
    password.Reset();
    return S_OK;
}

HRESULT FtpClient::PersistCredentials(const FtpConnectionOptions& options, const FtpCredential& credentials) {
    if (!options.allowCredentialPersistence) {
        return S_FALSE;
    }
    if (credentials.userName.empty() || credentials.password.empty()) {
        return S_FALSE;
    }
    const std::wstring target = BuildCredentialTargetName(options);
    CREDENTIALW cred = {};
    cred.Type = CRED_TYPE_GENERIC;
    cred.TargetName = const_cast<LPWSTR>(target.c_str());
    cred.UserName = const_cast<LPWSTR>(credentials.userName.c_str());
    cred.CredentialBlobSize = static_cast<DWORD>((credentials.password.size() + 1) * sizeof(wchar_t));
    cred.CredentialBlob = reinterpret_cast<LPBYTE>(const_cast<wchar_t*>(credentials.password.c_str()));
    cred.Persist = CRED_PERSIST_LOCAL_MACHINE;
    if (!CredWriteW(&cred, 0)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}

HRESULT FtpClient::RemovePersistedCredentials(const FtpConnectionOptions& options) {
    const std::wstring target = BuildCredentialTargetName(options);
    if (!CredDeleteW(target.c_str(), CRED_TYPE_GENERIC, 0)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    return S_OK;
}

std::wstring FtpClient::BuildCredentialTargetName(const FtpConnectionOptions& options) const {
    std::wstring target = L"ShellTabs/FTP/";
    target += FormatHostPort(options.host, options.port);
    return target;
}

FtpTransportType FtpClient::DetermineTransport(const FtpConnectionOptions& options) const {
    if (options.forcedTransport) {
        return *options.forcedTransport;
    }
    if (IsWinInetAvailable()) {
        return FtpTransportType::WinInet;
    }
    if (options.allowFallbackTransports && IsWinHttpAvailable()) {
        return FtpTransportType::WinHttp;
    }
    return FtpTransportType::CustomSockets;
}

std::vector<FtpTransportType> FtpClient::BuildTransportPriority(const FtpConnectionOptions& options) const {
    std::vector<FtpTransportType> transports;
    if (options.forcedTransport) {
        transports.push_back(*options.forcedTransport);
        return transports;
    }
    if (IsWinInetAvailable()) {
        transports.push_back(FtpTransportType::WinInet);
    }
    if (options.allowFallbackTransports && IsWinHttpAvailable()) {
        transports.push_back(FtpTransportType::WinHttp);
    }
    transports.push_back(FtpTransportType::CustomSockets);
    transports.erase(std::unique(transports.begin(), transports.end()), transports.end());
    return transports;
}

std::shared_ptr<FtpClient::FtpTransportSession> FtpClient::AcquireSession(const SessionKey& key,
                                                                           const FtpConnectionOptions& options,
                                                                           const FtpCredential& credentials,
                                                                           HRESULT* openResult) {
    std::shared_ptr<FtpTransportSession> session;
    {
        std::lock_guard<std::mutex> guard(impl_->mutex);
        auto& entries = impl_->pool[key];
        auto now = std::chrono::steady_clock::now();
        while (!entries.empty()) {
            auto entry = entries.back();
            entries.pop_back();
            if (!entry.session || !entry.session->IsAlive()) {
                continue;
            }
            if (now - entry.lastUsed > options.poolIdleTimeout) {
                continue;
            }
            session = entry.session;
            if (openResult) {
                *openResult = S_OK;
            }
            break;
        }
    }
    if (session) {
        return session;
    }
    switch (key.transport) {
        case FtpTransportType::WinInet:
            session = std::make_shared<WinInetFtpSession>();
            break;
        case FtpTransportType::WinHttp:
            session = std::make_shared<WinHttpFtpSession>();
            break;
        case FtpTransportType::CustomSockets:
            session = std::make_shared<SocketFtpSession>();
            break;
    }
    if (!session) {
        if (openResult) {
            *openResult = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        return nullptr;
    }
    const HRESULT hr = session->Open(options, credentials);
    if (FAILED(hr)) {
        if (openResult) {
            *openResult = hr;
        }
        return nullptr;
    }
    if (openResult) {
        *openResult = hr;
    }
    return session;
}

void FtpClient::ReleaseSession(const SessionKey& key, const std::shared_ptr<FtpTransportSession>& session, bool keepAlive) {
    if (!session) {
        return;
    }
    if (!keepAlive || !session->IsAlive()) {
        return;
    }
    std::lock_guard<std::mutex> guard(impl_->mutex);
    auto& entries = impl_->pool[key];
    entries.push_back({session, std::chrono::steady_clock::now()});
}

FtpClient::ExecuteResult FtpClient::ExecuteStateMachine(FtpTransportSession& session, FtpCommandStateMachine& machine,
                                                        const FtpConnectionOptions& options,
                                                        const FtpCredential& credentials, FtpOperationContext& context) {
    ExecuteResult result;
    result.hr = S_OK;
    while (machine.HasNext()) {
        auto step = machine.Next();
        switch (step.command) {
            case FtpCommandStateMachine::Command::SendUser:
                result.hr = session.SendUser(credentials.userName);
                break;
            case FtpCommandStateMachine::Command::SendPassword:
                result.hr = session.SendPassword(credentials.password);
                break;
            case FtpCommandStateMachine::Command::SetBinaryMode:
                result.hr = session.SetBinaryMode();
                break;
            case FtpCommandStateMachine::Command::ChangeDirectory:
                result.hr = session.ChangeDirectory(step.argument);
                break;
            case FtpCommandStateMachine::Command::EnterPassive:
                result.hr = options.passiveMode ? session.EnterPassiveMode(true, nullptr) : S_OK;
                break;
            case FtpCommandStateMachine::Command::List:
                result.hr = session.ListDirectory(false, context.directoryResults);
                break;
            case FtpCommandStateMachine::Command::Mlsd:
                result.hr = session.ListDirectory(true, context.directoryResults);
                break;
            case FtpCommandStateMachine::Command::Retrieve:
                result.hr = session.RetrieveFile(step.argument, context.localPath, context.transferResult);
                break;
            case FtpCommandStateMachine::Command::Store:
                result.hr = session.StoreFile(context.localPath, step.argument, context.transferResult);
                break;
            case FtpCommandStateMachine::Command::Complete:
                result.hr = session.CompleteOperation();
                break;
            case FtpCommandStateMachine::Command::None:
            default:
                result.hr = S_OK;
                break;
        }
        if (FAILED(result.hr)) {
            DWORD error = HRESULT_CODE(result.hr);
            result.retryable = error == WAIT_TIMEOUT || error == ERROR_INTERNET_TIMEOUT ||
                               error == ERROR_CONNECTION_ABORTED || error == ERROR_CONNECTION_REFUSED ||
                               error == ERROR_NETNAME_DELETED;
            return result;
        }
    }
    return result;
}

HRESULT FtpClient::MapInternetErrorToHresult(DWORD error, bool isSocketError) const {
    if (isSocketError) {
        return HResultFromSocketError(static_cast<int>(error));
    }
    return HResultFromInternetError(error);
}

void FtpClient::ClearExpiredSessions(const SessionKey& key, const FtpConnectionOptions& options) {
    std::lock_guard<std::mutex> guard(impl_->mutex);
    auto iter = impl_->pool.find(key);
    if (iter == impl_->pool.end()) {
        return;
    }
    auto now = std::chrono::steady_clock::now();
    auto& vector = iter->second;
    vector.erase(std::remove_if(vector.begin(), vector.end(), [&](const PooledSession& session) {
                      return !session.session || !session.session->IsAlive() ||
                             (now - session.lastUsed) > options.poolIdleTimeout;
                  }),
                 vector.end());
}

}  // namespace shelltabs

namespace shelltabs::ftp::testhooks {

HRESULT ParseDirectoryListing(const std::string& raw, std::vector<FtpDirectoryEntry>* entries) {
    return ExtractDirectoryListing(raw, entries);
}

}  // namespace shelltabs::ftp::testhooks

