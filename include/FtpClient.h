#pragma once

#include <windows.h>
#include <wininet.h>

#include <chrono>
#include <memory>
#include <optional>
#include <string>
#include <vector>

namespace shelltabs {

struct FtpUrlParts;

enum class FtpTransportType {
    WinInet,
    WinHttp,
    CustomSockets,
};

struct FtpConnectionOptions {
    std::wstring host;
    INTERNET_PORT port = INTERNET_DEFAULT_FTP_PORT;
    bool useExplicitFtps = false;
    bool passiveMode = true;
    bool preferMlsd = true;
    bool allowFallbackTransports = true;
    bool alwaysPromptForCredentials = false;
    bool allowCredentialPersistence = true;
    std::wstring initialPath;
    std::wstring serviceName = L"ftp";
    std::chrono::seconds poolIdleTimeout{180};
    unsigned int maxRetries = 3;
    unsigned int retryDelayMilliseconds = 500;
    std::optional<FtpTransportType> forcedTransport;
};

struct FtpCredential {
    std::wstring userName;
    std::wstring password;
    bool persisted = false;
};

struct FtpDirectoryEntry {
    std::wstring name;
    FILETIME lastWriteTime{};
    ULONGLONG size = 0;
    bool isDirectory = false;
};

struct FtpTransferResult {
    ULONGLONG bytesTransferred = 0;
};

enum class FtpOperationKind {
    DirectoryListing,
    Download,
    Upload,
};

struct FtpOperationContext {
    FtpOperationKind kind = FtpOperationKind::DirectoryListing;
    std::wstring remotePath;
    std::wstring localPath;
    bool useMlsd = true;
    std::vector<FtpDirectoryEntry>* directoryResults = nullptr;
    FtpTransferResult* transferResult = nullptr;
};

class FtpClient {
public:
    FtpClient();
    ~FtpClient();

    FtpClient(const FtpClient&) = delete;
    FtpClient& operator=(const FtpClient&) = delete;

    HRESULT ListDirectory(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                          std::vector<FtpDirectoryEntry>* results, HWND credentialParent = nullptr);

    HRESULT DownloadFile(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                         const std::wstring& remotePath, const std::wstring& localPath,
                         FtpTransferResult* transferResult, HWND credentialParent = nullptr);

    HRESULT UploadFile(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                       const std::wstring& localPath, const std::wstring& remotePath,
                       FtpTransferResult* transferResult, HWND credentialParent = nullptr);

    void ClearConnectionPool();

private:
    class FtpTransportSession;

    struct SessionKey {
        std::wstring host;
        INTERNET_PORT port = 0;
        bool useExplicitFtps = false;
        bool passiveMode = true;
        FtpTransportType transport = FtpTransportType::WinInet;

        bool operator==(const SessionKey& other) const noexcept;
    };

    struct SessionKeyHasher {
        size_t operator()(const SessionKey& key) const noexcept;
    };

    struct PooledSession {
        std::shared_ptr<FtpTransportSession> session;
        std::chrono::steady_clock::time_point lastUsed;
    };

    class FtpCommandStateMachine;

    struct ExecuteResult {
        HRESULT hr = E_FAIL;
        bool retryable = false;
    };

    HRESULT ExecuteOperation(const FtpConnectionOptions& options, const FtpCredential* explicitCredential,
                             HWND credentialParent, FtpOperationContext* context);

    HRESULT AcquireCredentials(const FtpConnectionOptions& options, HWND credentialParent,
                               FtpCredential* credentials, bool* shouldPersist);

    HRESULT PersistCredentials(const FtpConnectionOptions& options, const FtpCredential& credentials);
    HRESULT RemovePersistedCredentials(const FtpConnectionOptions& options);
    std::wstring BuildCredentialTargetName(const FtpConnectionOptions& options) const;

    std::shared_ptr<FtpTransportSession> AcquireSession(const SessionKey& key,
                                                        const FtpConnectionOptions& options,
                                                        const FtpCredential& credentials,
                                                        HRESULT* openResult);
    void ReleaseSession(const SessionKey& key, const std::shared_ptr<FtpTransportSession>& session, bool keepAlive,
                        const FtpConnectionOptions& options);

    FtpTransportType DetermineTransport(const FtpConnectionOptions& options) const;
    std::vector<FtpTransportType> BuildTransportPriority(const FtpConnectionOptions& options) const;

    ExecuteResult ExecuteStateMachine(FtpTransportSession& session, FtpCommandStateMachine& machine,
                                      const FtpConnectionOptions& options, const FtpCredential& credentials,
                                      FtpOperationContext& context);

    HRESULT MapInternetErrorToHresult(DWORD error, bool isSocketError) const;

    void ClearExpiredSessions(const SessionKey& key, const FtpConnectionOptions& options);

    std::wstring BuildCommandDescription(FtpOperationContext& context) const;

    struct Impl;
    std::unique_ptr<Impl> impl_;
};

}  // namespace shelltabs

