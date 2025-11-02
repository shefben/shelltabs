#include "EditGradientRenderer.h"

#include <CommCtrl.h>

#include <algorithm>
#include <cmath>
#include <string>
#include <vector>

#include "DpiUtils.h"

namespace shelltabs {

namespace {

RECT GetClientRectSafe(HWND hwnd) {
    RECT rect{0, 0, 0, 0};
    if (hwnd && IsWindow(hwnd)) {
        GetClientRect(hwnd, &rect);
    }
    return rect;
}

}  // namespace

bool RenderGradientEditContent(HWND hwnd, HDC dc, const BreadcrumbGradientConfig& gradientConfig,
                               const GradientEditRenderOptions& options) {
    if (!hwnd || !IsWindow(hwnd) || !dc || !gradientConfig.enabled) {
        return false;
    }

    RECT client = GetClientRectSafe(hwnd);
    if (client.right <= client.left || client.bottom <= client.top) {
        return false;
    }

    const UINT windowDpi = GetWindowDpi(hwnd);
    const UINT initialDcDpiX = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSX)));
    const UINT initialDcDpiY = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSY)));
    const bool dpiMismatch = (windowDpi != 0) &&
                             (initialDcDpiX != windowDpi || initialDcDpiY != windowDpi);

#ifdef DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    const ScopedThreadDpiAwarenessContext dpiScope(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, dpiMismatch);
#else
    const ScopedThreadDpiAwarenessContext dpiScope(nullptr, false);
#endif

    const BOOL caretHidden = options.hideCaret ? HideCaret(hwnd) : FALSE;

    if (options.requestEraseBackground) {
        LRESULT eraseResult = SendMessageW(hwnd, WM_ERASEBKGND, reinterpret_cast<WPARAM>(dc), 0);
        if (eraseResult == 0) {
            HBRUSH backgroundBrush = reinterpret_cast<HBRUSH>(GetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND));
            if (!backgroundBrush) {
                backgroundBrush = GetSysColorBrush(COLOR_WINDOW);
            }
            FillRect(dc, &client, backgroundBrush);
        }
    }

    std::wstring text;
    const int length = GetWindowTextLengthW(hwnd);
    if (length > 0) {
        text.resize(static_cast<size_t>(length) + 1);
        int copied = GetWindowTextW(hwnd, text.data(), length + 1);
        if (copied < 0) {
            copied = 0;
        }
        text.resize(static_cast<size_t>(std::clamp<int>(copied, 0, length)));
    }

    RECT formatRect = client;
    SendMessageW(hwnd, EM_GETRECT, 0, reinterpret_cast<LPARAM>(&formatRect));
    if (formatRect.right <= formatRect.left) {
        formatRect = client;
    }

    const UINT effectiveDcDpiX = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSX)));
    const UINT effectiveDcDpiY = static_cast<UINT>(std::max(0, GetDeviceCaps(dc, LOGPIXELSY)));
    const double scaleX = (windowDpi != 0) ? static_cast<double>(effectiveDcDpiX) /
                                                static_cast<double>(windowDpi)
                                          : 1.0;
    const double scaleY = (windowDpi != 0) ? static_cast<double>(effectiveDcDpiY) /
                                                static_cast<double>(windowDpi)
                                          : 1.0;

    auto scaleCoordinateX = [&](int value) -> double {
        return static_cast<double>(value) * scaleX;
    };
    auto scaleCoordinateY = [&](int value) -> double {
        return static_cast<double>(value) * scaleY;
    };

    auto scaleRectToDevice = [&](const RECT& rect) -> RECT {
        const double left = scaleCoordinateX(rect.left);
        const double right = scaleCoordinateX(rect.right);
        const double top = scaleCoordinateY(rect.top);
        const double bottom = scaleCoordinateY(rect.bottom);

        RECT result{};
        result.left = static_cast<LONG>(std::lround(std::min(left, right)));
        result.right = static_cast<LONG>(std::lround(std::max(left, right)));
        result.top = static_cast<LONG>(std::lround(std::min(top, bottom)));
        result.bottom = static_cast<LONG>(std::lround(std::max(top, bottom)));
        return result;
    };

    RECT scaledFormatRect = scaleRectToDevice(formatRect);
    RECT clipRect = scaledFormatRect;
    if (options.clipRect.has_value()) {
        RECT scaledClip = scaleRectToDevice(options.clipRect.value());
        RECT intersect{};
        if (IntersectRect(&intersect, &clipRect, &scaledClip)) {
            clipRect = intersect;
        } else {
            clipRect = RECT{0, 0, 0, 0};
        }
    }

    if (clipRect.right <= clipRect.left || clipRect.bottom <= clipRect.top) {
        if (caretHidden) {
            ShowCaret(hwnd);
        }
        return true;
    }

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    if (!font) {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    HFONT oldFont = nullptr;
    if (font) {
        oldFont = static_cast<HFONT>(SelectObject(dc, font));
    }

    const BreadcrumbGradientPalette gradientPalette = ResolveBreadcrumbGradientPalette(gradientConfig);

    const COLORREF previousTextColor = GetTextColor(dc);
    const int previousBkMode = SetBkMode(dc, TRANSPARENT);

    struct CharacterMetrics {
        int index;
        double x;
        double y;
        double width;
    };

    std::vector<CharacterMetrics> characters;
    characters.reserve(text.size());

    double gradientLeft = static_cast<double>(scaledFormatRect.right);
    double gradientRight = static_cast<double>(scaledFormatRect.left);

    for (int i = 0; i < static_cast<int>(text.size()); ++i) {
        LRESULT pos = SendMessageW(hwnd, EM_POSFROMCHAR, i, 0);
        if (pos == -1) {
            continue;
        }

        const int rawCharX = static_cast<SHORT>(LOWORD(static_cast<DWORD_PTR>(pos)));
        const int rawCharY = static_cast<SHORT>(HIWORD(static_cast<DWORD_PTR>(pos)));
        const double charX = scaleCoordinateX(rawCharX);
        const double charY = scaleCoordinateY(rawCharY);

        LRESULT nextPos = SendMessageW(hwnd, EM_POSFROMCHAR, i + 1, 0);
        double nextX = charX;
        if (nextPos != -1) {
            const int rawNextX = static_cast<SHORT>(LOWORD(static_cast<DWORD_PTR>(nextPos)));
            nextX = scaleCoordinateX(rawNextX);
        }
        if (nextPos == -1 || nextX <= charX) {
            SIZE extent{};
            if (GetTextExtentPoint32W(dc, &text[i], 1, &extent) && extent.cx > 0) {
                nextX = charX + static_cast<double>(extent.cx);
            } else {
                nextX = charX + 1.0;
            }
        }

        const double charWidth = std::max(1.0, nextX - charX);
        characters.push_back(CharacterMetrics{i, charX, charY, charWidth});

        gradientLeft = std::min(gradientLeft, charX);
        gradientRight = std::max(gradientRight, charX + charWidth);
    }

    if (!characters.empty()) {
        gradientLeft = std::min(gradientLeft, characters.front().x - characters.front().width * 0.5);
        gradientRight = std::max(gradientRight, characters.back().x + characters.back().width * 0.5);
    }

    const double scaledLeftBound = static_cast<double>(scaledFormatRect.left);
    const double scaledRightBound = static_cast<double>(scaledFormatRect.right);

    gradientLeft = std::clamp<double>(gradientLeft, scaledLeftBound, scaledRightBound);
    gradientRight = std::clamp<double>(gradientRight, scaledLeftBound, scaledRightBound);

    if (gradientRight <= gradientLeft) {
        gradientLeft = scaledLeftBound;
        gradientRight = std::max(scaledLeftBound + 1.0, scaledRightBound);
    }

    const double gradientWidth = std::max(1.0, gradientRight - gradientLeft);

    DWORD selectionStart = 0;
    DWORD selectionEnd = 0;
    if (SendMessageW(hwnd, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart),
                     reinterpret_cast<LPARAM>(&selectionEnd)) == 0) {
        selectionStart = selectionEnd = 0;
    }
    if (selectionEnd < selectionStart) {
        std::swap(selectionStart, selectionEnd);
    }
    const bool hasSelection = (selectionEnd > selectionStart);
    const COLORREF highlightTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT);

    if (hasSelection) {
        const COLORREF highlightColor = GetSysColor(COLOR_HIGHLIGHT);
        HBRUSH selectionBrush = CreateSolidBrush(highlightColor);
        bool deleteSelectionBrush = true;
        if (!selectionBrush) {
            selectionBrush = GetSysColorBrush(COLOR_HIGHLIGHT);
            deleteSelectionBrush = false;
        }

        TEXTMETRICW metrics{};
        int lineHeight = 0;
        if (GetTextMetricsW(dc, &metrics)) {
            lineHeight = metrics.tmHeight + metrics.tmExternalLeading;
        }
        if (lineHeight <= 0) {
            lineHeight = std::max<LONG>(1, scaledFormatRect.bottom - scaledFormatRect.top);
        }

        auto paintSelectionRun = [&](double startX, double endX, int top, int bottom) {
            if (!selectionBrush) {
                return;
            }
            RECT selectionRect{};
            selectionRect.left = static_cast<LONG>(std::floor(startX));
            selectionRect.right = static_cast<LONG>(std::ceil(endX));
            selectionRect.top = top;
            selectionRect.bottom = bottom;

            RECT intersected{};
            if (IntersectRect(&intersected, &selectionRect, &clipRect)) {
                FillRect(dc, &intersected, selectionBrush);
            }
        };

        bool runActive = false;
        double runStartX = 0.0;
        double runEndX = 0.0;
        int runTop = clipRect.top;
        int runBottom = clipRect.top;

        auto flushRun = [&]() {
            if (!runActive) {
                return;
            }
            paintSelectionRun(runStartX, runEndX, runTop, runBottom);
            runActive = false;
        };

        const LONG clipTop = clipRect.top;
        const LONG clipBottom = clipRect.bottom;

        for (const CharacterMetrics& character : characters) {
            const bool isSelected = character.index >= static_cast<int>(selectionStart) &&
                                   character.index < static_cast<int>(selectionEnd);
            if (!isSelected) {
                flushRun();
                continue;
            }

            const double charLeft = character.x;
            const double charRight = character.x + character.width;
            int charTop = static_cast<int>(std::floor(character.y));
            int charBottom = charTop + lineHeight;
            charTop = static_cast<int>(
                std::clamp<LONG>(static_cast<LONG>(charTop), clipTop, clipBottom));
            charBottom = static_cast<int>(
                std::clamp<LONG>(static_cast<LONG>(charBottom), clipTop, clipBottom));
            if (charBottom <= charTop) {
                charBottom = std::min<int>(static_cast<int>(clipBottom), charTop + lineHeight);
            }
            if (charBottom <= charTop) {
                charBottom = std::min<int>(static_cast<int>(clipBottom), charTop + 1);
            }

            if (!runActive) {
                runActive = true;
                runStartX = charLeft;
                runTop = charTop;
                runBottom = charBottom;
            }

            runEndX = charRight;
            runTop = std::min(runTop, charTop);
            runBottom = std::max(runBottom, charBottom);
        }

        flushRun();

        if (deleteSelectionBrush && selectionBrush) {
            DeleteObject(selectionBrush);
        }
    }

    for (const CharacterMetrics& character : characters) {
        const bool isSelected = hasSelection && character.index >= static_cast<int>(selectionStart) &&
                                character.index < static_cast<int>(selectionEnd);

        if (isSelected) {
            SetTextColor(dc, highlightTextColor);
        } else {
            const double centerX = character.x + character.width / 2.0;
            double position = (centerX - gradientLeft) / gradientWidth;
            position = std::clamp<double>(position, 0.0, 1.0);

            const COLORREF gradientColor = EvaluateBreadcrumbGradientColor(gradientPalette, position);
            SetTextColor(dc, gradientColor);
        }

        const int drawX = static_cast<int>(std::lround(character.x));
        const int drawY = static_cast<int>(std::lround(character.y));

        ExtTextOutW(dc, drawX, drawY, ETO_CLIPPED, &clipRect, text.data() + character.index, 1, nullptr);
    }

    SetBkMode(dc, previousBkMode);
    SetTextColor(dc, previousTextColor);

    if (oldFont) {
        SelectObject(dc, oldFont);
    }

    if (caretHidden) {
        ShowCaret(hwnd);
    }

    return true;
}

}  // namespace shelltabs

