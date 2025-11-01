#pragma once

#include <cstdint>
#include <string>
#include <string_view>
#include <type_traits>
#include <vector>

namespace shelltabs {

enum class TabBandDockMode;
enum class NewTabTemplate;

bool EqualsIgnoreCase(std::wstring_view lhs, std::wstring_view rhs);
std::wstring_view TrimView(std::wstring_view value);
std::wstring Trim(std::wstring_view value);
std::vector<std::wstring_view> Split(std::wstring_view value, wchar_t delimiter);
bool ParseBool(std::wstring_view token);
TabBandDockMode ParseDockMode(std::wstring_view token);
std::wstring DockModeToString(TabBandDockMode mode);
NewTabTemplate ParseNewTabTemplate(std::wstring_view token);
std::wstring NewTabTemplateToString(NewTabTemplate value);

template <typename Callback>
bool ParseConfigLines(std::wstring_view content, wchar_t commentChar, wchar_t delimiter, Callback&& callback) {
    size_t lineStart = 0;
    while (lineStart < content.size()) {
        const size_t lineEnd = content.find(L'\n', lineStart);
        std::wstring_view line =
            content.substr(lineStart, lineEnd == std::wstring_view::npos ? std::wstring_view::npos : lineEnd - lineStart);
        if (lineEnd == std::wstring_view::npos) {
            lineStart = content.size();
        } else {
            lineStart = lineEnd + 1;
        }

        line = TrimView(line);
        if (line.empty() || line.front() == commentChar) {
            continue;
        }

        auto tokens = Split(line, delimiter);
        for (auto& token : tokens) {
            token = TrimView(token);
        }

        using CallbackResult = std::invoke_result_t<Callback&, const std::vector<std::wstring_view>&>;
        if constexpr (std::is_same_v<std::decay_t<CallbackResult>, bool>) {
            if (!callback(tokens)) {
                return false;
            }
        } else {
            static_assert(std::is_void_v<CallbackResult>,
                          "ParseConfigLines callback must return void or bool");
            callback(tokens);
        }
    }

    return true;
}

}  // namespace shelltabs
int ParseInt(std::wstring_view token);
bool TryParseUint64(std::wstring_view token, uint64_t* valueOut);

