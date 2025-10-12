#pragma once
#include <windows.h>
#include <string>
#include <unordered_map>
#include <vector>
#include <mutex>

namespace shelltabs {

	// Persists per-path text color overrides as JSON in %APPDATA%\ShellTabs\namecolors.json
	class FileColorOverrides {
	public:
		static FileColorOverrides& Instance();

                bool TryGetColor(const std::wstring& path, COLORREF* out) const;
                void SetColor(const std::vector<std::wstring>& paths, COLORREF color);
                void ClearColor(const std::vector<std::wstring>& paths);

                // Ephemeral overrides are kept in-memory only. They are ideal for transient
                // visualisations such as folder comparisons where persisting colours to disk
                // would be undesirable.
                void SetEphemeralColor(const std::vector<std::wstring>& paths, COLORREF color);
                void ClearEphemeral();

	private:
		FileColorOverrides() = default;
		FileColorOverrides(const FileColorOverrides&) = delete;
		FileColorOverrides& operator=(const FileColorOverrides&) = delete;

		void Load() const;   // lazy
		void Save() const;

		static std::wstring StoragePath();      // %APPDATA%\ShellTabs\namecolors.json
		static std::wstring ToLowerCopy(std::wstring s);

		mutable std::mutex mtx_;
		mutable bool loaded_ = false;
                mutable std::unordered_map<std::wstring, COLORREF> map_;       // persistent colours
                mutable std::unordered_map<std::wstring, COLORREF> transient_; // in-memory only
	};

} // namespace shelltabs
