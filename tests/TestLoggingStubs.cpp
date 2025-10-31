#include "Logging.h"

namespace shelltabs {

void LogMessage(LogLevel, const wchar_t*, ...) {}

void LogLastError(const wchar_t*, DWORD) {}

}  // namespace shelltabs

