#include "Utilities.h"

#include <string>

namespace shelltabs {

std::wstring NormalizeFileSystemPath(const std::wstring& path) { return path; }

bool IsLikelyFileSystemPath(const std::wstring&) { return false; }

std::wstring GetCanonicalParsingName(PCIDLIST_ABSOLUTE) { return {}; }

std::wstring GetParsingName(PCIDLIST_ABSOLUTE) { return {}; }

}  // namespace shelltabs
