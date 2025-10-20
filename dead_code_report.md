# Unreferenced Functions

- `TryGetFileSystemPath(IShellItem*, std::wstring*)` — Declared in `include/Utilities.h` and defined in `src/Utilities.cpp`; repository searches show only those declarations with no call sites.
- `NormalizeFileSystemPath(const std::wstring&)` — Declared in `include/Utilities.h` and defined in `src/Utilities.cpp`; it is only used by `TryGetFileSystemPath`, leaving no external callers.
