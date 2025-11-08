#include "Logging.h"

#include "Utilities.h"

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

// Prevent winsock.h from being included by windows.h to avoid conflicts with winsock2.h
#ifndef _WINSOCKAPI_
#define _WINSOCKAPI_
#endif

#include <windows.h>
#include <psapi.h>

#include <atomic>
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
PVOID g_vectoredExceptionHandler = nullptr;
std::atomic<DWORD> g_accessViolationCount{0};

constexpr wchar_t kLogDirectory[] = L"ShellTabs\\Logs";
constexpr DWORD kMaxAccessViolationsSuppressed = 100;

std::wstring_view LevelToString(LogLevel level) {
    switch (level) {
        case LogLevel::Verbose:
            return L"VERBOSE";
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

bool IsAddressInShellTabsModule(const void* address) {
    if (!address || !g_hostModule) {
        return false;
    }

    MODULEINFO moduleInfo{};
    if (!GetModuleInformation(GetCurrentProcess(), g_hostModule, &moduleInfo, sizeof(moduleInfo))) {
        return false;
    }

    const auto moduleStart = reinterpret_cast<BYTE*>(moduleInfo.lpBaseOfDll);
    const auto moduleEnd = moduleStart + moduleInfo.SizeOfImage;
    const auto testAddress = reinterpret_cast<const BYTE*>(address);

    return testAddress >= moduleStart && testAddress < moduleEnd;
}

LONG CALLBACK VectoredExceptionHandler(_In_ struct _EXCEPTION_POINTERS* info) {
    if (!info || !info->ExceptionRecord || !info->ContextRecord) {
        return EXCEPTION_CONTINUE_SEARCH;
    }

    const auto code = info->ExceptionRecord->ExceptionCode;
    const void* address = info->ExceptionRecord->ExceptionAddress;

    // Log all exceptions from our module with detailed information
    if (IsAddressInShellTabsModule(address)) {
        // Get more details about the exception
        if (code == EXCEPTION_ACCESS_VIOLATION) {
            const DWORD count = ++g_accessViolationCount;
            const ULONG_PTR* exceptionInfo = info->ExceptionRecord->ExceptionInformation;
            const ULONG_PTR isWrite = (info->ExceptionRecord->NumberParameters > 0) ? exceptionInfo[0] : 0;
            const ULONG_PTR faultAddress = (info->ExceptionRecord->NumberParameters > 1) ? exceptionInfo[1] : 0;

            LogMessage(LogLevel::Error,
                       L"Access violation #%lu in ShellTabs: %ls at instruction %p, fault address %p",
                       count,
                       isWrite ? L"write" : L"read",
                       address,
                       reinterpret_cast<void*>(faultAddress));

            // For null pointer dereferences (fault address is near 0), suppress the crash
            // This handles cases where we accidentally dereference null pointers
            if (faultAddress < 0x10000 && count <= kMaxAccessViolationsSuppressed) {
                LogMessage(LogLevel::Warning,
                           L"Null pointer dereference detected and suppressed to prevent Explorer crash (count: %lu/%lu)",
                           count, kMaxAccessViolationsSuppressed);

                // Return from the faulting function safely by unwinding the stack
                // We simulate a function return by setting the instruction pointer to the return address
                // and adjusting the stack pointer
#if defined(_M_X64) || defined(_M_AMD64)
                // On x64, read the return address from the stack and set up for return
                // Use __try to safely read from the stack in case of corruption
                __try {
                    ULONG_PTR* stackPtr = reinterpret_cast<ULONG_PTR*>(info->ContextRecord->Rsp);
                    ULONG_PTR returnAddress = *stackPtr;
                    info->ContextRecord->Rip = returnAddress;  // Set instruction pointer to return address
                    info->ContextRecord->Rsp += sizeof(ULONG_PTR);  // Pop return address from stack
                    info->ContextRecord->Rax = 0;  // Set return value to 0/NULL
                    return EXCEPTION_CONTINUE_EXECUTION;  // Continue execution from the return address
                }
                __except (EXCEPTION_EXECUTE_HANDLER) {
                    LogMessage(LogLevel::Error,
                               L"Failed to read return address from stack (stack corruption?), allowing exception to propagate");
                }
#elif defined(_M_IX86)
                // On x86, read the return address from the stack and set up for return
                __try {
                    ULONG_PTR* stackPtr = reinterpret_cast<ULONG_PTR*>(info->ContextRecord->Esp);
                    ULONG_PTR returnAddress = *stackPtr;
                    info->ContextRecord->Eip = returnAddress;  // Set instruction pointer to return address
                    info->ContextRecord->Esp += sizeof(ULONG_PTR);  // Pop return address from stack
                    info->ContextRecord->Eax = 0;  // Set return value to 0/NULL
                    return EXCEPTION_CONTINUE_EXECUTION;  // Continue execution from the return address
                }
                __except (EXCEPTION_EXECUTE_HANDLER) {
                    LogMessage(LogLevel::Error,
                               L"Failed to read return address from stack (stack corruption?), allowing exception to propagate");
                }
#endif
            }
        } else {
            LogMessage(LogLevel::Error,
                       L"Exception 0x%08X in ShellTabs at %p",
                       code,
                       address);
        }
    }

    return EXCEPTION_CONTINUE_SEARCH;
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

    // Install vectored exception handler first (called before SEH handlers)
    // This allows us to catch and suppress access violations to prevent Explorer crashes
    g_vectoredExceptionHandler = AddVectoredExceptionHandler(1, &VectoredExceptionHandler);
    if (!g_vectoredExceptionHandler) {
        LogMessage(LogLevel::Warning, L"Failed to install vectored exception handler");
    }

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

    if (g_vectoredExceptionHandler) {
        RemoveVectoredExceptionHandler(g_vectoredExceptionHandler);
        g_vectoredExceptionHandler = nullptr;
    }

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

