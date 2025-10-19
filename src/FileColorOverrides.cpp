#include "FileColorOverrides.h"

#include "CommonDialogColorizer.h"
#include <ShlObj.h>
#include <Shlwapi.h>
#include <vector>
#include <fstream>
#include <sstream>
#include <codecvt>
#include <locale>

#pragma comment(lib, "Shlwapi.lib")

namespace shelltabs {
	namespace {
		static std::wstring ColorToHex(COLORREF c) {
			wchar_t buf[8];
			// #RRGGBB
			swprintf(buf, _countof(buf), L"#%02X%02X%02X",
				GetRValue(c), GetGValue(c), GetBValue(c));
			return buf;
		}

		static bool HexToColor(const std::wstring& s, COLORREF* out) {
			if (!out || s.size() != 7 || s[0] != L'#') return false;
			unsigned r = 0, g = 0, b = 0;
			if (swscanf(s.c_str() + 1, L"%02x%02x%02x", &r, &g, &b) != 3) return false;
			*out = RGB(r, g, b);
			return true;
		}
		static std::wstring JsonEscape(const std::wstring& in) {
			std::wstring out; out.reserve(in.size() + 8);
			for (wchar_t ch : in) {
				if (ch == L'\\' || ch == L'"') { out.push_back(L'\\'); out.push_back(ch); }
				else out.push_back(ch);
			}
			return out;
		}
		static std::wstring CombinePath(PCWSTR a, PCWSTR b) {
			wchar_t buf[MAX_PATH * 2]{};
			lstrcpynW(buf, a, _countof(buf));
			PathAppendW(buf, b);
			return buf;
		}
	}

	FileColorOverrides& FileColorOverrides::Instance() {
		static FileColorOverrides inst;
		return inst;
	}

	std::wstring FileColorOverrides::ToLowerCopy(std::wstring s) {
		if (!s.empty()) CharLowerBuffW(s.data(), static_cast<DWORD>(s.size()));
		return s;
	}

	std::wstring FileColorOverrides::StoragePath() {
		PWSTR roaming = nullptr;
		if (SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, nullptr, &roaming) != S_OK) return L"";
		std::wstring dir = CombinePath(roaming, L"ShellTabs");
		CoTaskMemFree(roaming);
		CreateDirectoryW(dir.c_str(), nullptr);
		return CombinePath(dir.c_str(), L"namecolors.json");
	}

	void FileColorOverrides::Load() const {
		std::lock_guard<std::mutex> lock(mtx_);
		if (loaded_) return;
		loaded_ = true;

		const auto path = StoragePath();
		if (path.empty()) return;

		std::wifstream fin(path);
		if (!fin) return;
		fin.imbue(std::locale(fin.getloc(), new std::codecvt_utf8_utf16<wchar_t>));
		std::wstringstream ss; ss << fin.rdbuf();
		const std::wstring json = ss.str();

		std::unordered_map<std::wstring, COLORREF> tmp;
		size_t i = 0;
		auto skipws = [&]() { while (i < json.size() && iswspace(json[i])) ++i; };
		auto readstr = [&]() -> std::wstring {
			std::wstring s;
			if (i >= json.size() || json[i] != L'"') return s;
			++i;
			for (; i < json.size(); ++i) {
				wchar_t ch = json[i];
				if (ch == L'\\') {
					if (i + 1 < json.size()) { s.push_back(json[i + 1]); ++i; }
				}
				else if (ch == L'"') { ++i; break; }
				else s.push_back(ch);
			}
			return s;
			};

		skipws();
		if (i >= json.size() || json[i] != L'{') return;
		++i;
		while (i < json.size()) {
			skipws();
			if (i < json.size() && json[i] == L'}') { ++i; break; }
			auto key = readstr();
			skipws(); if (i >= json.size() || json[i] != L':') break; ++i;
			skipws();
			auto val = readstr();
			COLORREF c;
			if (!key.empty() && HexToColor(val, &c)) tmp.emplace(ToLowerCopy(key), c);
			skipws();
			if (i < json.size() && json[i] == L',') { ++i; continue; }
			if (i < json.size() && json[i] == L'}') { ++i; break; }
		}
		map_.swap(tmp);
	}

	void FileColorOverrides::Save() const {
		std::lock_guard<std::mutex> lock(mtx_);
		const auto path = StoragePath();
		if (path.empty()) return;

		std::wostringstream ss;
		ss << L"{";
		bool first = true;
		for (const auto& kv : map_) {
			if (!first) ss << L",";
			first = false;
			ss << L"\"" << JsonEscape(kv.first) << L"\":\"" << ColorToHex(kv.second) << L"\"";
		}
		ss << L"}";

		std::wofstream fout(path, std::ios::trunc);
		fout.imbue(std::locale(fout.getloc(), new std::codecvt_utf8_utf16<wchar_t>));
		fout << ss.str();
	}

        bool FileColorOverrides::TryGetColor(const std::wstring& path, COLORREF* out) const {
                Load();
                std::lock_guard<std::mutex> lock(mtx_);
                const auto key = ToLowerCopy(path);

                auto transientIt = transient_.find(key);
                if (transientIt != transient_.end()) {
                        if (out) *out = transientIt->second;
                        return true;
                }

                auto it = map_.find(key);
                if (it == map_.end()) return false;
                if (out) *out = it->second;
                return true;
        }

        void FileColorOverrides::SetColor(const std::vector<std::wstring>& paths, COLORREF color) {
                Load();
                {
                        std::lock_guard<std::mutex> lock(mtx_);
                        for (const auto& p : paths) map_[ToLowerCopy(p)] = color;
                }
                Save();
                CommonDialogColorizer::NotifyColorDataChanged();
        }

        void FileColorOverrides::ClearColor(const std::vector<std::wstring>& paths) {
                Load();
                {
                        std::lock_guard<std::mutex> lock(mtx_);
                        for (const auto& p : paths) map_.erase(ToLowerCopy(p));
                }
                Save();
                CommonDialogColorizer::NotifyColorDataChanged();
        }

        void FileColorOverrides::SetEphemeralColor(const std::vector<std::wstring>& paths, COLORREF color) {
                std::lock_guard<std::mutex> lock(mtx_);
                for (const auto& p : paths) transient_[ToLowerCopy(p)] = color;
                CommonDialogColorizer::NotifyColorDataChanged();
        }

        void FileColorOverrides::ClearEphemeral() {
                std::lock_guard<std::mutex> lock(mtx_);
                transient_.clear();
                CommonDialogColorizer::NotifyColorDataChanged();
        }

        void FileColorOverrides::TransferColor(const std::wstring& fromPath, const std::wstring& toPath) {
                if (fromPath.empty() || toPath.empty()) {
                        return;
                }

                const auto fromKey = ToLowerCopy(fromPath);
                const auto toKey = ToLowerCopy(toPath);
                if (fromKey == toKey) {
                        return;
                }

                Load();

                bool persistentChanged = false;
                bool transientChanged = false;
                {
                        std::lock_guard<std::mutex> lock(mtx_);

                        auto mapIt = map_.find(fromKey);
                        if (mapIt != map_.end()) {
                                COLORREF color = mapIt->second;
                                map_.erase(mapIt);
                                map_[toKey] = color;
                                persistentChanged = true;
                        }

                        auto transientIt = transient_.find(fromKey);
                        if (transientIt != transient_.end()) {
                                COLORREF color = transientIt->second;
                                transient_.erase(transientIt);
                                transient_[toKey] = color;
                                transientChanged = true;
                        }
                }

                if (persistentChanged) {
                        Save();
                }
                if (persistentChanged || transientChanged) {
                        CommonDialogColorizer::NotifyColorDataChanged();
                }
        }

} // namespace shelltabs
