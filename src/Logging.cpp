#include "Logging.h"

#include <windows.h>

#include <chrono>
#include <cwchar>
#include <string>
#include <string_view>

namespace shelltabs {

namespace {
INIT_ONCE g_loggingInitOnce = INIT_ONCE_STATIC_INIT;
CRITICAL_SECTION g_logLock;
HANDLE g_logFile = INVALID_HANDLE_VALUE;
bool g_logLockInitialized = false;
bool g_loggingReady = false;
bool g_loggingShutdown = false;
HMODULE g_hostModule = nullptr;
LPTOP_LEVEL_EXCEPTION_FILTER g_previousFilter = nullptr;

constexpr wchar_t kLogDirectory[] = L"ShellTabs\\Logs";

std::wstring_view LevelToString(LogLevel level) {
    switch (level) {
        case LogLevel::Info:
            return L"INFO";
        case LogLevel::Warning:
            return L"WARN";
        case LogLevel::Error:
        default:
            return L"ERROR";
    }
}

std::wstring FormatWideString(const wchar_t* format, va_list args) {
    if (!format) {
        return {};
    }

    va_list copy;
    va_copy(copy, args);
    int required = _vscwprintf(format, copy);
    va_end(copy);
    if (required <= 0) {
        return {};
    }

    std::wstring buffer(static_cast<size_t>(required) + 1, L'\0');
    va_copy(copy, args);
    int written = vswprintf(buffer.data(), buffer.size(), format, copy);
    va_end(copy);
    if (written <= 0) {
        return {};
    }

    buffer.resize(static_cast<size_t>(written));
    return buffer;
}

std::wstring GetEnvironmentValue(const wchar_t* name) {
    DWORD required = GetEnvironmentVariableW(name, nullptr, 0);
    if (required == 0) {
        return {};
    }

    std::wstring value(static_cast<size_t>(required), L'\0');
    DWORD written = GetEnvironmentVariableW(name, value.data(), required);
    if (written == 0 || written >= required) {
        return {};
    }
    value.resize(written);
    return value;
}

bool EnsureDirectoryExists(const std::wstring& path) {
    if (path.empty()) {
        return false;
    }

    if (CreateDirectoryW(path.c_str(), nullptr)) {
        return true;
    }

    const DWORD error = GetLastError();
    if (error == ERROR_ALREADY_EXISTS) {
        return true;
    }
    if (error != ERROR_PATH_NOT_FOUND) {
        return false;
    }

    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring::npos) {
        return false;
    }

    const std::wstring parent = path.substr(0, separator);
    if (!EnsureDirectoryExists(parent)) {
        return false;
    }

    if (CreateDirectoryW(path.c_str(), nullptr)) {
        return true;
    }
    return GetLastError() == ERROR_ALREADY_EXISTS;
}

std::wstring BuildLogDirectory() {
    std::wstring base = GetEnvironmentValue(L"LOCALAPPDATA");
    if (base.empty()) {
        wchar_t tempPath[MAX_PATH];
        DWORD written = GetTempPathW(ARRAYSIZE(tempPath), tempPath);
        if (written == 0 || written >= ARRAYSIZE(tempPath)) {
            return {};
        }
        base.assign(tempPath, written);
    }

    if (!base.empty()) {
        if (base.back() != L'\\' && base.back() != L'/') {
            base.push_back(L'\\');
        }
    }

    base.append(kLogDirectory);
    return base;
}

std::wstring BuildProcessDescription() {
    wchar_t path[MAX_PATH] = {};
    DWORD written = GetModuleFileNameW(nullptr, path, ARRAYSIZE(path));
    if (written == 0 || written >= ARRAYSIZE(path)) {
        return L"(unknown)";
    }
    return std::wstring(path, written);
}

std::wstring BuildLogFilePath() {
    std::wstring directory = BuildLogDirectory();
    if (directory.empty()) {
        return {};
    }

    if (!EnsureDirectoryExists(directory)) {
        return {};
    }

    SYSTEMTIME st;
    GetSystemTime(&st);

    wchar_t fileName[128];
    swprintf(fileName, ARRAYSIZE(fileName), L"shelltabs-%04u%02u%02u-%02u%02u%02u-%03u-%lu.log", st.wYear, st.wMonth,
             st.wDay, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, GetCurrentProcessId());

    if (directory.back() != L'\\' && directory.back() != L'/') {
        directory.push_back(L'\\');
    }
    directory.append(fileName);
    return directory;
}

std::string WideToUtf8(const std::wstring& value) {
    if (value.empty()) {
        return {};
    }

    int required = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0) {
        return {};
    }
    std::string buffer(static_cast<size_t>(required), '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), buffer.data(), required, nullptr, nullptr);
    return buffer;
}

void WriteLineToFile(const std::wstring& line) {
    if (!g_loggingReady || g_logFile == INVALID_HANDLE_VALUE || !g_logLockInitialized) {
        return;
    }

    std::string utf8 = WideToUtf8(line);
    if (utf8.empty()) {
        return;
    }

    EnterCriticalSection(&g_logLock);
    DWORD written = 0;
    WriteFile(g_logFile, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
    LeaveCriticalSection(&g_logLock);
}

std::wstring Trimmed(std::wstring value) {
    while (!value.empty() && (value.back() == L'\r' || value.back() == L'\n')) {
        value.pop_back();
    }
    return value;
}

std::wstring DescribeSystemMessage(DWORD code) {
    wchar_t buffer[512];
    DWORD flags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    DWORD written = FormatMessageW(flags, nullptr, code, 0, buffer, ARRAYSIZE(buffer), nullptr);
    if (written == 0) {
        return {};
    }
    return Trimmed(std::wstring(buffer, written));
}

LONG CALLBACK UnhandledExceptionFilterCallback(_In_ struct _EXCEPTION_POINTERS* info) {
    if (!info || !info->ExceptionRecord) {
        return EXCEPTION_CONTINUE_SEARCH;
    }

    const auto code = info->ExceptionRecord->ExceptionCode;
    const void* address = info->ExceptionRecord->ExceptionAddress;
    LogMessage(LogLevel::Error, L"Unhandled exception 0x%08X at %p", code, address);
    return EXCEPTION_CONTINUE_SEARCH;
}

BOOL CALLBACK InitializeLoggingOnce(PINIT_ONCE, PVOID parameter, PVOID*) {
    g_hostModule = static_cast<HMODULE>(parameter);

    if (InitializeCriticalSectionEx(&g_logLock, 4000, 0)) {
        g_logLockInitialized = true;
    } else {
        g_logLockInitialized = false;
    }

    const std::wstring path = BuildLogFilePath();
    if (path.empty()) {
        g_loggingReady = false;
        return TRUE;
    }

    HANDLE file = CreateFileW(path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
                              OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        g_loggingReady = false;
        return TRUE;
    }

    g_logFile = file;
    g_loggingReady = true;

    const std::wstring process = BuildProcessDescription();
    std::wstring header = L"=== ShellTabs logging started for " + process + L" ===\r\n";
    WriteLineToFile(header);
    OutputDebugStringW(header.c_str());

    g_previousFilter = SetUnhandledExceptionFilter(&UnhandledExceptionFilterCallback);

    return TRUE;
}

void EnsureLoggingInitialized() {
    if (g_loggingShutdown) {
        return;
    }
    InitOnceExecuteOnce(&g_loggingInitOnce, InitializeLoggingOnce, g_hostModule, nullptr);
}

std::wstring BuildLogLine(LogLevel level, const std::wstring& message) {
    SYSTEMTIME st;
    GetSystemTime(&st);
    wchar_t timestamp[64];
    swprintf(timestamp, ARRAYSIZE(timestamp), L"%04u-%02u-%02uT%02u:%02u:%02u.%03uZ", st.wYear, st.wMonth, st.wDay,
             st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);

    wchar_t prefix[128];
    swprintf(prefix, ARRAYSIZE(prefix), L"[%ls][pid %lu][tid %lu] ", timestamp, GetCurrentProcessId(), GetCurrentThreadId());

    std::wstring line(prefix);
    line.append(LevelToString(level));
    line.append(L" ");
    line.append(message);
    line.append(L"\r\n");
    return line;
}

}  // namespace

void InitializeLoggingEarly(HMODULE module) noexcept {
    g_hostModule = module;
    EnsureLoggingInitialized();
}

void ShutdownLogging() noexcept {
    if (g_loggingShutdown) {
        return;
    }

    LogMessage(LogLevel::Info, L"Logging shutting down");
    g_loggingShutdown = true;

    if (g_previousFilter) {
        SetUnhandledExceptionFilter(g_previousFilter);
        g_previousFilter = nullptr;
    }

    if (g_logFile != INVALID_HANDLE_VALUE) {
        HANDLE file = g_logFile;
        if (g_logLockInitialized) {
            EnterCriticalSection(&g_logLock);
            file = g_logFile;
            g_logFile = INVALID_HANDLE_VALUE;
            LeaveCriticalSection(&g_logLock);
        } else {
            g_logFile = INVALID_HANDLE_VALUE;
        }
        if (file != INVALID_HANDLE_VALUE) {
            CloseHandle(file);
        }
    }

    if (g_logLockInitialized) {
        DeleteCriticalSection(&g_logLock);
        g_logLockInitialized = false;
    }
}

void LogMessage(LogLevel level, const wchar_t* format, ...) noexcept {
    va_list args;
    va_start(args, format);
    LogMessageV(level, format, args);
    va_end(args);
}

void LogMessageV(LogLevel level, const wchar_t* format, va_list args) noexcept {
    EnsureLoggingInitialized();

    std::wstring message = FormatWideString(format, args);
    if (message.empty()) {
        return;
    }

    std::wstring line = BuildLogLine(level, message);
    OutputDebugStringW(line.c_str());
    WriteLineToFile(line);
}

void LogLastError(const wchar_t* operation, DWORD error) noexcept {
    std::wstring description = DescribeSystemMessage(error);
    if (description.empty()) {
        LogMessage(LogLevel::Error, L"%ls failed with Win32 error %lu", operation ? operation : L"(unknown operation)", error);
    } else {
        LogMessage(LogLevel::Error, L"%ls failed with Win32 error %lu: %ls", operation ? operation : L"(unknown operation)",
                   error, description.c_str());
    }
}

void LogHrFailure(const wchar_t* operation, HRESULT hr) noexcept {
    if (HRESULT_FACILITY(hr) == FACILITY_WIN32) {
        const DWORD win32 = HRESULT_CODE(hr);
        LogLastError(operation, win32);
        return;
    }

    std::wstring description = DescribeSystemMessage(static_cast<DWORD>(hr));
    if (description.empty()) {
        LogMessage(LogLevel::Error, L"%ls failed with HRESULT 0x%08X", operation ? operation : L"(unknown operation)", hr);
    } else {
        LogMessage(LogLevel::Error, L"%ls failed with HRESULT 0x%08X: %ls", operation ? operation : L"(unknown operation)", hr,
                   description.c_str());
    }
}

LogScope::LogScope(const wchar_t* scope) noexcept : m_scope(scope) {
    if (m_scope && *m_scope) {
        LogMessage(LogLevel::Info, L"%ls - begin", m_scope);
    }
}

LogScope::~LogScope() {
    if (m_scope && *m_scope) {
        LogMessage(LogLevel::Info, L"%ls - end", m_scope);
    }
}

}  // namespace shelltabs

