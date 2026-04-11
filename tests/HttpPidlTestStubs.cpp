// Minimal stubs for HttpPidlTests so we don't drag in all of Utilities.cpp/Logging.cpp
#include "Utilities.h"

#include <objbase.h>

namespace shelltabs {

void PidlDeleter::operator()(AbsolutePidl* pidl) const noexcept {
    if (pidl) {
        CoTaskMemFree(pidl);
    }
}

}  // namespace shelltabs
