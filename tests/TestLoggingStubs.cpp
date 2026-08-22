#include "Logging.h"

namespace shelltabs {

void LogMessage(LogLevel, const wchar_t*, ...) noexcept {}

void LogMessageV(LogLevel, const wchar_t*, va_list) noexcept {}

void LogLastError(const wchar_t*, DWORD) noexcept {}

void LogHrFailure(const wchar_t*, HRESULT) noexcept {}

}  // namespace shelltabs
