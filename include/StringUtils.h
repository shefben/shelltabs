#pragma once

#include <string>
#include <type_traits>
#include <vector>

namespace shelltabs {

enum class TabBandDockMode;
enum class NewTabTemplate;

std::wstring Trim(const std::wstring& value);
std::vector<std::wstring> Split(const std::wstring& value, wchar_t delimiter);
bool ParseBool(const std::wstring& token);
TabBandDockMode ParseDockMode(const std::wstring& token);
std::wstring DockModeToString(TabBandDockMode mode);
NewTabTemplate ParseNewTabTemplate(const std::wstring& token);
std::wstring NewTabTemplateToString(NewTabTemplate value);

template <typename Callback>
bool ParseConfigLines(const std::wstring& content, wchar_t commentChar, wchar_t delimiter, Callback&& callback) {
    size_t lineStart = 0;
    while (lineStart < content.size()) {
        const size_t lineEnd = content.find(L'\n', lineStart);
        std::wstring line =
            content.substr(lineStart, lineEnd == std::wstring::npos ? std::wstring::npos : lineEnd - lineStart);
        if (lineEnd == std::wstring::npos) {
            lineStart = content.size();
        } else {
            lineStart = lineEnd + 1;
        }

        line = Trim(line);
        if (line.empty() || line.front() == commentChar) {
            continue;
        }

        auto tokens = Split(line, delimiter);
        for (auto& token : tokens) {
            token = Trim(token);
        }

        using CallbackResult = std::invoke_result_t<Callback&, const std::vector<std::wstring>&>;
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
