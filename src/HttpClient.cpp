#include "HttpClient.h"
#include "HtmlDirectoryParser.h"
#include "Logging.h"
#include "Utilities.h"

#include <algorithm>
#include <cstring>
#include <string>

namespace shelltabs::http {

namespace {

constexpr DWORD kReadBufferSize = 65536;

std::wstring BuildFullPath(const std::wstring& basePath, const std::wstring& relativePath) {
    std::wstring path = basePath;
    if (!path.empty() && path.back() != L'/') {
        path.push_back(L'/');
    }
    if (!relativePath.empty() && relativePath.front() == L'/') {
        path += relativePath.substr(1);
    } else {
        path += relativePath;
    }
    return path;
}

}  // namespace

HttpClient::HttpClient() = default;
HttpClient::~HttpClient() = default;

HRESULT HttpClient::AcquireSession(const HttpConnectionOptions& options, HINTERNET* session, HINTERNET* connection) {
    if (!session || !connection) {
        return E_POINTER;
    }

    SessionKey key{options.host, options.port, options.useHttps};

    {
        std::lock_guard<std::mutex> lock(poolMutex_);
        auto it = pool_.find(key);
        if (it != pool_.end()) {
            auto elapsed = std::chrono::steady_clock::now() - it->second.lastUsed;
            if (elapsed < options.poolIdleTimeout) {
                *session = it->second.session.handle;
                *connection = it->second.connection.handle;
                // Take ownership out of pool without closing
                it->second.session.handle = nullptr;
                it->second.connection.handle = nullptr;
                pool_.erase(it);
                return S_OK;
            }
            pool_.erase(it);
        }
    }

    HINTERNET hSession = WinHttpOpen(L"ShellTabs/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) {
        DWORD err = GetLastError();
        LogMessage(LogLevel::Error, L"[HttpClient] AcquireSession: WinHttpOpen failed err=%lu", err);
        return HRESULT_FROM_WIN32(err);
    }

    // Set timeouts: 15s connect, 30s send, 30s receive
    WinHttpSetTimeouts(hSession, 0, 15000, 30000, 30000);

    LogMessage(LogLevel::Info, L"[HttpClient] AcquireSession: connecting to %ls:%d",
               options.host.c_str(), static_cast<int>(options.port));
    HINTERNET hConnection = WinHttpConnect(hSession, options.host.c_str(),
                                           static_cast<INTERNET_PORT>(options.port), 0);
    if (!hConnection) {
        DWORD err = GetLastError();
        LogMessage(LogLevel::Error, L"[HttpClient] AcquireSession: WinHttpConnect failed err=%lu for %ls:%d",
                   err, options.host.c_str(), static_cast<int>(options.port));
        WinHttpCloseHandle(hSession);
        return HRESULT_FROM_WIN32(err);
    }

    *session = hSession;
    *connection = hConnection;
    return S_OK;
}

void HttpClient::ReturnSession(const HttpConnectionOptions& options, WinHttpHandle session, WinHttpHandle connection) {
    SessionKey key{options.host, options.port, options.useHttps};
    std::lock_guard<std::mutex> lock(poolMutex_);
    PooledSession pooled;
    pooled.session = std::move(session);
    pooled.connection = std::move(connection);
    pooled.lastUsed = std::chrono::steady_clock::now();
    pool_[key] = std::move(pooled);
}

HRESULT HttpClient::SendRequest(HINTERNET connection, const wchar_t* verb, const std::wstring& path,
                                bool useHttps, WinHttpHandle* request) {
    DWORD flags = useHttps ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hRequest = WinHttpOpenRequest(connection, verb, path.c_str(), nullptr,
                                            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!hRequest) {
        DWORD err = GetLastError();
        LogMessage(LogLevel::Error, L"[HttpClient] SendRequest: WinHttpOpenRequest failed err=%lu for %ls %ls",
                   err, verb, path.c_str());
        return HRESULT_FROM_WIN32(err);
    }

    // Follow redirects automatically
    DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS;
    WinHttpSetOption(hRequest, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

    if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
        DWORD err = GetLastError();
        LogMessage(LogLevel::Error, L"[HttpClient] SendRequest: WinHttpSendRequest failed err=%lu for %ls %ls",
                   err, verb, path.c_str());
        WinHttpCloseHandle(hRequest);
        return HRESULT_FROM_WIN32(err);
    }

    if (!WinHttpReceiveResponse(hRequest, nullptr)) {
        DWORD err = GetLastError();
        LogMessage(LogLevel::Error, L"[HttpClient] SendRequest: WinHttpReceiveResponse failed err=%lu", err);
        WinHttpCloseHandle(hRequest);
        return HRESULT_FROM_WIN32(err);
    }

    // Check status code
    DWORD statusCode = 0;
    DWORD statusSize = sizeof(statusCode);
    WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                        WINHTTP_HEADER_NAME_BY_INDEX, &statusCode, &statusSize, WINHTTP_NO_HEADER_INDEX);

    LogMessage(LogLevel::Info, L"[HttpClient] SendRequest: %ls %ls -> status %lu", verb, path.c_str(), statusCode);

    if (statusCode >= 400) {
        LogMessage(LogLevel::Error, L"[HttpClient] SendRequest: HTTP error %lu for %ls %ls", statusCode, verb, path.c_str());
        WinHttpCloseHandle(hRequest);
        if (statusCode == 404) {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }
        if (statusCode == 403) {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        return HRESULT_FROM_WIN32(ERROR_WINHTTP_INVALID_SERVER_RESPONSE);
    }

    *request = WinHttpHandle(hRequest);
    return S_OK;
}

HRESULT HttpClient::ReadResponseBody(HINTERNET request, std::vector<BYTE>* data) {
    if (!data) {
        return E_POINTER;
    }
    data->clear();

    std::vector<BYTE> buffer(kReadBufferSize);
    DWORD bytesRead = 0;
    while (WinHttpReadData(request, buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead)) {
        if (bytesRead == 0) {
            break;
        }
        data->insert(data->end(), buffer.data(), buffer.data() + bytesRead);
    }
    return S_OK;
}

HRESULT HttpClient::ListDirectory(const HttpConnectionOptions& options, const std::wstring& path,
                                  std::vector<DirectoryEntry>* results) {
    if (!results) {
        return E_POINTER;
    }
    results->clear();

    LogMessage(LogLevel::Info, L"[HttpClient] ListDirectory: host=%ls basePath=%ls path=%ls https=%d port=%d",
               options.host.c_str(), options.basePath.c_str(), path.c_str(),
               options.useHttps ? 1 : 0, static_cast<int>(options.port));

    HINTERNET hSession = nullptr;
    HINTERNET hConnection = nullptr;
    HRESULT hr = AcquireSession(options, &hSession, &hConnection);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpClient] ListDirectory: AcquireSession failed hr=0x%08X", hr);
        return hr;
    }

    WinHttpHandle session(hSession);
    WinHttpHandle connection(hConnection);

    std::wstring fullPath = BuildFullPath(options.basePath, path);
    if (!fullPath.empty() && fullPath.back() != L'/') {
        fullPath.push_back(L'/');
    }
    LogMessage(LogLevel::Info, L"[HttpClient] ListDirectory: GET %ls", fullPath.c_str());

    WinHttpHandle request;
    hr = SendRequest(connection.get(), L"GET", fullPath, options.useHttps, &request);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpClient] ListDirectory: SendRequest failed hr=0x%08X", hr);
        return hr;
    }

    std::vector<BYTE> body;
    hr = ReadResponseBody(request.get(), &body);
    if (FAILED(hr)) {
        LogMessage(LogLevel::Error, L"[HttpClient] ListDirectory: ReadResponseBody failed hr=0x%08X", hr);
        return hr;
    }

    LogMessage(LogLevel::Info, L"[HttpClient] ListDirectory: received %zu bytes", body.size());

    // Convert UTF-8 to wide string
    std::string_view utf8(reinterpret_cast<const char*>(body.data()), body.size());
    std::wstring html = Utf8ToWide(utf8);

    // Parse the HTML
    ParseResult parsed = ParseDirectoryListing(html, options.host);
    LogMessage(LogLevel::Info, L"[HttpClient] ListDirectory: parsed %zu entries, format=%d",
               parsed.entries.size(), static_cast<int>(parsed.format));
    *results = std::move(parsed.entries);

    // Return session to pool
    ReturnSession(options, std::move(session), std::move(connection));
    return S_OK;
}

HRESULT HttpClient::DownloadFile(const HttpConnectionOptions& options, const std::wstring& remotePath,
                                 const std::wstring& localPath, HttpProgressCallback callback) {
    HINTERNET hSession = nullptr;
    HINTERNET hConnection = nullptr;
    HRESULT hr = AcquireSession(options, &hSession, &hConnection);
    if (FAILED(hr)) {
        return hr;
    }

    WinHttpHandle session(hSession);
    WinHttpHandle connection(hConnection);

    std::wstring fullPath = BuildFullPath(options.basePath, remotePath);

    WinHttpHandle request;
    hr = SendRequest(connection.get(), L"GET", fullPath, options.useHttps, &request);
    if (FAILED(hr)) {
        return hr;
    }

    // Get content length
    ULONGLONG totalSize = 0;
    wchar_t contentLengthStr[64] = {};
    DWORD contentLengthSize = sizeof(contentLengthStr);
    if (WinHttpQueryHeaders(request.get(), WINHTTP_QUERY_CONTENT_LENGTH, WINHTTP_HEADER_NAME_BY_INDEX,
                            contentLengthStr, &contentLengthSize, WINHTTP_NO_HEADER_INDEX)) {
        totalSize = _wtoi64(contentLengthStr);
    }

    // Create output file
    HANDLE hFile = CreateFileW(localPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    std::vector<BYTE> buffer(kReadBufferSize);
    ULONGLONG downloaded = 0;
    DWORD bytesRead = 0;
    bool cancelled = false;

    while (WinHttpReadData(request.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead)) {
        if (bytesRead == 0) {
            break;
        }

        DWORD bytesWritten = 0;
        if (!WriteFile(hFile, buffer.data(), bytesRead, &bytesWritten, nullptr)) {
            DWORD err = GetLastError();
            CloseHandle(hFile);
            DeleteFileW(localPath.c_str());
            return HRESULT_FROM_WIN32(err);
        }

        downloaded += bytesRead;

        if (callback) {
            if (!callback(downloaded, totalSize)) {
                cancelled = true;
                break;
            }
        }
    }

    CloseHandle(hFile);

    if (cancelled) {
        DeleteFileW(localPath.c_str());
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    ReturnSession(options, std::move(session), std::move(connection));
    return S_OK;
}

HRESULT HttpClient::DownloadToMemory(const HttpConnectionOptions& options, const std::wstring& remotePath,
                                     std::vector<BYTE>* data) {
    if (!data) {
        return E_POINTER;
    }

    HINTERNET hSession = nullptr;
    HINTERNET hConnection = nullptr;
    HRESULT hr = AcquireSession(options, &hSession, &hConnection);
    if (FAILED(hr)) {
        return hr;
    }

    WinHttpHandle session(hSession);
    WinHttpHandle connection(hConnection);

    std::wstring fullPath = BuildFullPath(options.basePath, remotePath);

    WinHttpHandle request;
    hr = SendRequest(connection.get(), L"GET", fullPath, options.useHttps, &request);
    if (FAILED(hr)) {
        return hr;
    }

    hr = ReadResponseBody(request.get(), data);
    if (FAILED(hr)) {
        return hr;
    }

    ReturnSession(options, std::move(session), std::move(connection));
    return S_OK;
}

HRESULT HttpClient::HeadRequest(const HttpConnectionOptions& options, const std::wstring& remotePath,
                                std::wstring* contentType, ULONGLONG* contentLength) {
    HINTERNET hSession = nullptr;
    HINTERNET hConnection = nullptr;
    HRESULT hr = AcquireSession(options, &hSession, &hConnection);
    if (FAILED(hr)) {
        return hr;
    }

    WinHttpHandle session(hSession);
    WinHttpHandle connection(hConnection);

    std::wstring fullPath = BuildFullPath(options.basePath, remotePath);

    WinHttpHandle request;
    hr = SendRequest(connection.get(), L"HEAD", fullPath, options.useHttps, &request);
    if (FAILED(hr)) {
        return hr;
    }

    if (contentType) {
        wchar_t buffer[256] = {};
        DWORD bufferSize = sizeof(buffer);
        if (WinHttpQueryHeaders(request.get(), WINHTTP_QUERY_CONTENT_TYPE, WINHTTP_HEADER_NAME_BY_INDEX,
                                buffer, &bufferSize, WINHTTP_NO_HEADER_INDEX)) {
            *contentType = buffer;
        }
    }

    if (contentLength) {
        wchar_t buffer[64] = {};
        DWORD bufferSize = sizeof(buffer);
        if (WinHttpQueryHeaders(request.get(), WINHTTP_QUERY_CONTENT_LENGTH, WINHTTP_HEADER_NAME_BY_INDEX,
                                buffer, &bufferSize, WINHTTP_NO_HEADER_INDEX)) {
            *contentLength = _wtoi64(buffer);
        }
    }

    ReturnSession(options, std::move(session), std::move(connection));
    return S_OK;
}

void HttpClient::ClearConnectionPool() {
    std::lock_guard<std::mutex> lock(poolMutex_);
    pool_.clear();
}

// ============================================================================
// HttpDownloadQueue
// ============================================================================

HttpDownloadQueue::~HttpDownloadQueue() {
    CancelAll();
}

void HttpDownloadQueue::Configure(int maxConcurrent, int speedLimitKBps) {
    std::lock_guard<std::mutex> lock(m_mutex);
    m_maxConcurrent = std::clamp(maxConcurrent, 1, 16);
    m_speedLimitKBps = std::max(0, speedLimitKBps);
}

HRESULT HttpDownloadQueue::Enqueue(DownloadTask task) {
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_queue.push_back(std::move(task));
    }
    EnsureWorkers();
    m_cv.notify_one();
    return S_OK;
}

void HttpDownloadQueue::CancelAll() {
    {
        std::lock_guard<std::mutex> lock(m_mutex);
        m_shutdown.store(true, std::memory_order_release);
        m_queue.clear();
    }
    m_cv.notify_all();
    for (auto& t : m_workers) {
        if (t.joinable()) t.join();
    }
    m_workers.clear();
    m_workersStarted = false;
    m_shutdown.store(false, std::memory_order_release);
}

size_t HttpDownloadQueue::PendingCount() const {
    std::lock_guard<std::mutex> lock(const_cast<std::mutex&>(m_mutex));
    return m_queue.size();
}

void HttpDownloadQueue::EnsureWorkers() {
    std::lock_guard<std::mutex> lock(m_mutex);
    if (m_workersStarted) return;
    m_workersStarted = true;
    m_shutdown.store(false, std::memory_order_release);
    int count = m_maxConcurrent;
    for (int i = 0; i < count; ++i) {
        m_workers.emplace_back([this]() { WorkerLoop(); });
    }
}

void HttpDownloadQueue::WorkerLoop() {
    HttpClient client;
    while (true) {
        DownloadTask task;
        {
            std::unique_lock<std::mutex> lock(m_mutex);
            m_cv.wait(lock, [this]() {
                return m_shutdown.load(std::memory_order_acquire) || !m_queue.empty();
            });
            if (m_shutdown.load(std::memory_order_acquire) && m_queue.empty()) break;
            if (m_queue.empty()) continue;
            task = std::move(m_queue.front());
            m_queue.pop_front();
        }

        // Wrap callback with throttling
        HttpProgressCallback throttledCallback;
        if (m_speedLimitKBps > 0 || task.callback) {
            throttledCallback = [this, &task](ULONGLONG downloaded, ULONGLONG total) -> bool {
                if (m_speedLimitKBps > 0) {
                    ThrottleIfNeeded(static_cast<size_t>(downloaded));
                }
                if (task.callback) {
                    return task.callback(downloaded, total);
                }
                return !m_shutdown.load(std::memory_order_acquire);
            };
        }

        client.DownloadFile(task.options, task.remotePath, task.localPath, throttledCallback);
    }
}

void HttpDownloadQueue::ThrottleIfNeeded(size_t bytesJustRead) {
    if (m_speedLimitKBps <= 0) return;

    std::lock_guard<std::mutex> lock(m_speedMutex);
    auto now = std::chrono::steady_clock::now();
    auto elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(now - m_windowStart);

    if (elapsed >= std::chrono::seconds(1)) {
        m_bytesThisSecond.store(0, std::memory_order_relaxed);
        m_windowStart = now;
    }

    m_bytesThisSecond.fetch_add(bytesJustRead, std::memory_order_relaxed);
    size_t limit = static_cast<size_t>(m_speedLimitKBps) * 1024;
    if (m_bytesThisSecond.load(std::memory_order_relaxed) > limit) {
        auto remaining = std::chrono::seconds(1) - elapsed;
        if (remaining > std::chrono::milliseconds(0)) {
            std::this_thread::sleep_for(remaining);
            m_bytesThisSecond.store(0, std::memory_order_relaxed);
            m_windowStart = std::chrono::steady_clock::now();
        }
    }
}

}  // namespace shelltabs::http
