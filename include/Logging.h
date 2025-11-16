#pragma once

#include <windows.h>

#include <cstdint>
#include <cstdarg>

namespace shelltabs {

enum class LogLevel {
    Verbose,
    Info,
    Warning,
    Error,
};

struct FaultMitigationDetails {
    uint64_t faultId = 0;
    DWORD exceptionCode = 0;
    const struct _EXCEPTION_POINTERS* exceptionPointers = nullptr;
};

using FaultMitigationCallback = void (*)(const FaultMitigationDetails& details, void* context);

void InitializeLoggingEarly(HMODULE module) noexcept;
void ShutdownLogging() noexcept;

void LogMessage(LogLevel level, const wchar_t* format, ...) noexcept;
void LogMessageV(LogLevel level, const wchar_t* format, va_list args) noexcept;
void LogLastError(const wchar_t* operation, DWORD error) noexcept;
void LogHrFailure(const wchar_t* operation, HRESULT hr) noexcept;

void RegisterFaultMitigationHandler(FaultMitigationCallback callback, void* context) noexcept;
bool HasFaultMitigationTriggered() noexcept;

class LogScope {
public:
    explicit LogScope(const wchar_t* scope) noexcept;
    ~LogScope();

    LogScope(const LogScope&) = delete;
    LogScope& operator=(const LogScope&) = delete;

private:
    const wchar_t* m_scope;
};

}  // namespace shelltabs

