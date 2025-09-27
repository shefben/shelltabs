#pragma once

#include <windows.h>

#include <cstdarg>

namespace shelltabs {

enum class LogLevel {
    Info,
    Warning,
    Error,
};

void InitializeLoggingEarly(HMODULE module) noexcept;
void ShutdownLogging() noexcept;

void LogMessage(LogLevel level, const wchar_t* format, ...) noexcept;
void LogMessageV(LogLevel level, const wchar_t* format, va_list args) noexcept;
void LogLastError(const wchar_t* operation, DWORD error) noexcept;
void LogHrFailure(const wchar_t* operation, HRESULT hr) noexcept;

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

