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
#include <DbgHelp.h>
#include <psapi.h>

#include <atomic>
#include <chrono>
#include <cwchar>
#include <cwctype>
#include <string>
#include <string_view>
#include <algorithm>
#include <vector>

namespace shelltabs {

namespace {
INIT_ONCE g_loggingInitOnce = INIT_ONCE_STATIC_INIT;
INIT_ONCE g_symbolInitOnce = INIT_ONCE_STATIC_INIT;
CRITICAL_SECTION g_logLock;
HANDLE g_logFile = INVALID_HANDLE_VALUE;
bool g_logLockInitialized = false;
bool g_loggingReady = false;
bool g_loggingShutdown = false;
HMODULE g_hostModule = nullptr;
LPTOP_LEVEL_EXCEPTION_FILTER g_previousFilter = nullptr;
PVOID g_vectoredExceptionHandler = nullptr;
std::atomic<uint64_t> g_faultSequence{0};
std::atomic<bool> g_faultMitigationTriggered{false};

struct FaultMitigationHandlerEntry {
    FaultMitigationCallback callback = nullptr;
    void* context = nullptr;
};

SRWLOCK g_faultHandlersLock = SRWLOCK_INIT;
std::vector<FaultMitigationHandlerEntry> g_faultHandlers;
bool g_symbolHandlerReady = false;

constexpr wchar_t kMinidumpOptInEnvironmentVariable[] = L"SHELLTABS_WRITE_MINIDUMPS";

constexpr wchar_t kLogDirectory[] = L"ShellTabs\\Logs";
constexpr USHORT kMaxStackFrames = 64;

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

bool ParseBooleanEnvironmentValue(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(), [](wchar_t ch) { return std::towlower(ch); });
    return value == L"1" || value == L"true" || value == L"yes" || value == L"on";
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

bool ShouldWriteMinidumps() {
    const std::wstring value = GetEnvironmentValue(kMinidumpOptInEnvironmentVariable);
    if (value.empty()) {
        return false;
    }
    return ParseBooleanEnvironmentValue(value);
}

bool IsMinidumpOptInEnabled() {
    static const bool enabled = ShouldWriteMinidumps();
    return enabled;
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

BOOL CALLBACK InitializeSymbolsOnce(PINIT_ONCE, PVOID, PVOID*) {
    HANDLE process = GetCurrentProcess();
    SymSetOptions(SYMOPT_DEFERRED_LOADS | SYMOPT_UNDNAME | SYMOPT_LOAD_LINES | SYMOPT_FAIL_CRITICAL_ERRORS);
    g_symbolHandlerReady = SymInitializeW(process, nullptr, TRUE) != FALSE;
    if (!g_symbolHandlerReady) {
        const DWORD error = GetLastError();
        LogMessage(LogLevel::Warning, L"Failed to initialize symbol handler (error=%lu)", error);
    }
    return TRUE;
}

void EnsureSymbolHandlerInitialized() {
    InitOnceExecuteOnce(&g_symbolInitOnce, InitializeSymbolsOnce, nullptr, nullptr);
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

std::wstring DescribeModuleForAddress(const void* address) {
    if (!address) {
        return L"(unknown module)";
    }

    HMODULE module = nullptr;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                            reinterpret_cast<LPCWSTR>(address), &module) || !module) {
        return L"(unknown module)";
    }

    wchar_t path[MAX_PATH] = {};
    DWORD written = GetModuleFileNameW(module, path, ARRAYSIZE(path));
    if (written == 0 || written >= ARRAYSIZE(path)) {
        return L"(unknown module)";
    }

    MODULEINFO info{};
    if (!GetModuleInformation(GetCurrentProcess(), module, &info, sizeof(info))) {
        return std::wstring(path, written);
    }

    const auto moduleBase = reinterpret_cast<const BYTE*>(info.lpBaseOfDll);
    const auto target = reinterpret_cast<const BYTE*>(address);
    const size_t offset = static_cast<size_t>(target - moduleBase);

    wchar_t description[MAX_PATH + 32] = {};
    swprintf(description, ARRAYSIZE(description), L"%ls+0x%zX", path, offset);
    return description;
}

std::wstring DescribeSymbolForAddress(const void* address, std::wstring* lineInfo) {
    if (!address) {
        return L"(unknown symbol)";
    }

    EnsureSymbolHandlerInitialized();
    if (!g_symbolHandlerReady) {
        return L"(symbols unavailable)";
    }

    HANDLE process = GetCurrentProcess();
    BYTE buffer[sizeof(SYMBOL_INFOW) + (MAX_SYM_NAME * sizeof(wchar_t))];
    auto* symbol = reinterpret_cast<SYMBOL_INFOW*>(buffer);
    symbol->SizeOfStruct = sizeof(SYMBOL_INFOW);
    symbol->MaxNameLen = MAX_SYM_NAME;

    DWORD64 displacement = 0;
    if (!SymFromAddrW(process, reinterpret_cast<DWORD64>(address), &displacement, symbol)) {
        return L"(symbol lookup failed)";
    }

    if (lineInfo) {
        IMAGEHLP_LINEW64 line = {};
        line.SizeOfStruct = sizeof(line);
        DWORD lineDisplacement = 0;
        if (SymGetLineFromAddrW64(process, reinterpret_cast<DWORD64>(address), &lineDisplacement, &line)) {
            wchar_t lineBuffer[MAX_PATH + 32] = {};
            swprintf(lineBuffer, ARRAYSIZE(lineBuffer), L"%ls:%lu", line.FileName ? line.FileName : L"?", line.LineNumber);
            lineInfo->assign(lineBuffer);
        }
    }

    wchar_t symbolBuffer[MAX_SYM_NAME + 32] = {};
    swprintf(symbolBuffer, ARRAYSIZE(symbolBuffer), L"%ls+0x%llX", symbol->Name, displacement);
    return symbolBuffer;
}

void LogStackTrace(uint64_t faultId) {
    void* frames[kMaxStackFrames] = {};
    const USHORT captured = RtlCaptureStackBackTrace(0, kMaxStackFrames, frames, nullptr);
    if (captured == 0) {
        LogMessage(LogLevel::Warning, L"[fault %llu] Unable to capture stack trace", faultId);
        return;
    }

    LogMessage(LogLevel::Error, L"[fault %llu] Stack trace (%hu frames)", faultId, captured);
    const USHORT skip = captured > 2 ? 2 : 0;
    for (USHORT i = skip; i < captured; ++i) {
        std::wstring lineInfo;
        const std::wstring module = DescribeModuleForAddress(frames[i]);
        const std::wstring symbol = DescribeSymbolForAddress(frames[i], &lineInfo);
        if (lineInfo.empty()) {
            LogMessage(LogLevel::Error, L"[fault %llu]   #%02hu %p %ls | %ls", faultId, static_cast<USHORT>(i - skip),
                       frames[i], module.c_str(), symbol.c_str());
        } else {
            LogMessage(LogLevel::Error, L"[fault %llu]   #%02hu %p %ls | %ls (%ls)", faultId,
                       static_cast<USHORT>(i - skip), frames[i], module.c_str(), symbol.c_str(), lineInfo.c_str());
        }
    }
}

std::wstring BuildMinidumpFilePath(uint64_t faultId) {
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
    swprintf(fileName, ARRAYSIZE(fileName), L"shelltabs-fault-%04u%02u%02u-%02u%02u%02u-%03u-%llu.dmp", st.wYear, st.wMonth,
             st.wDay, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, faultId);

    if (directory.back() != L'\\' && directory.back() != L'/') {
        directory.push_back(L'\\');
    }
    directory.append(fileName);
    return directory;
}

bool WriteMinidumpToPath(const std::wstring& path, const EXCEPTION_POINTERS* info) {
    if (path.empty() || !info) {
        return false;
    }

    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }

    MINIDUMP_EXCEPTION_INFORMATION exceptionInfo{};
    exceptionInfo.ThreadId = GetCurrentThreadId();
    exceptionInfo.ExceptionPointers = const_cast<EXCEPTION_POINTERS*>(info);
    exceptionInfo.ClientPointers = FALSE;

    const BOOL result = MiniDumpWriteDump(GetCurrentProcess(), GetCurrentProcessId(), file,
                                          MiniDumpWithDataSegs | MiniDumpWithFullMemoryInfo |
                                              MiniDumpWithHandleData | MiniDumpScanMemory,
                                          info ? &exceptionInfo : nullptr, nullptr, nullptr);
    CloseHandle(file);
    return result != FALSE;
}

void MaybeWriteMinidump(uint64_t faultId, const EXCEPTION_POINTERS* info) {
    if (!IsMinidumpOptInEnabled() || !info) {
        return;
    }

    const std::wstring path = BuildMinidumpFilePath(faultId);
    if (path.empty()) {
        LogMessage(LogLevel::Warning, L"[fault %llu] Failed to build minidump path", faultId);
        return;
    }

    if (WriteMinidumpToPath(path, info)) {
        LogMessage(LogLevel::Info, L"[fault %llu] Minidump written to %ls", faultId, path.c_str());
    } else {
        LogMessage(LogLevel::Warning, L"[fault %llu] Failed to write minidump to %ls (error=%lu)", faultId, path.c_str(),
                   GetLastError());
    }
}

void InvokeFaultMitigationHandlers(const FaultMitigationDetails& details) {
    if (g_faultMitigationTriggered.exchange(true)) {
        return;
    }

    LogMessage(LogLevel::Warning, L"[fault %llu] Triggering ShellTabs fault mitigation (code=0x%08X)", details.faultId,
               details.exceptionCode);

    std::vector<FaultMitigationHandlerEntry> handlers;
    AcquireSRWLockShared(&g_faultHandlersLock);
    handlers = g_faultHandlers;
    ReleaseSRWLockShared(&g_faultHandlersLock);

    for (const FaultMitigationHandlerEntry& entry : handlers) {
        if (entry.callback) {
            entry.callback(details, entry.context);
        }
    }
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
    if (!info || !info->ExceptionRecord) {
        return EXCEPTION_CONTINUE_SEARCH;
    }

    const auto code = info->ExceptionRecord->ExceptionCode;
    const void* address = info->ExceptionRecord->ExceptionAddress;

    if (!IsAddressInShellTabsModule(address)) {
        return EXCEPTION_CONTINUE_SEARCH;
    }

    const uint64_t faultId = ++g_faultSequence;
    const std::wstring moduleDescription = DescribeModuleForAddress(address);

    if (code == EXCEPTION_ACCESS_VIOLATION && info->ExceptionRecord->NumberParameters >= 2) {
        const ULONG_PTR* exceptionInfo = info->ExceptionRecord->ExceptionInformation;
        const ULONG_PTR isWrite = exceptionInfo ? exceptionInfo[0] : 0;
        const ULONG_PTR faultAddress = exceptionInfo ? exceptionInfo[1] : 0;
        LogMessage(LogLevel::Error,
                   L"ShellTabs access violation [fault %llu]: %ls at %p (%ls) targeting %p on thread %lu",
                   faultId,
                   isWrite ? L"write" : L"read",
                   address,
                   moduleDescription.c_str(),
                   reinterpret_cast<void*>(faultAddress),
                   GetCurrentThreadId());
    } else {
        LogMessage(LogLevel::Error,
                   L"ShellTabs exception [fault %llu]: code=0x%08X at %p (%ls) on thread %lu",
                   faultId,
                   code,
                   address,
                   moduleDescription.c_str(),
                   GetCurrentThreadId());
    }

    LogStackTrace(faultId);
    MaybeWriteMinidump(faultId, info);

    FaultMitigationDetails details{};
    details.faultId = faultId;
    details.exceptionCode = code;
    details.exceptionPointers = info;
    InvokeFaultMitigationHandlers(details);

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

    if (g_symbolHandlerReady) {
        SymCleanup(GetCurrentProcess());
        g_symbolHandlerReady = false;
    }
}

void RegisterFaultMitigationHandler(FaultMitigationCallback callback, void* context) noexcept {
    if (!callback) {
        return;
    }

    AcquireSRWLockExclusive(&g_faultHandlersLock);
    g_faultHandlers.push_back({callback, context});
    ReleaseSRWLockExclusive(&g_faultHandlersLock);
}

bool HasFaultMitigationTriggered() noexcept {
    return g_faultMitigationTriggered.load();
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

