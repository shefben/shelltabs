#include <windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <shobjidl.h>
#include <wrl/client.h>

// ShellTabs Out-of-Process Preview Worker
int wmain(int argc, wchar_t* argv[]) {
    // Prevent Windows Error Reporting / crash dialogs
    SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOGPFAULTERRORBOX | SEM_NOOPENFILEERRORBOX);

    if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
        return 1;
    }

    HANDLE hStdIn = GetStdHandle(STD_INPUT_HANDLE);
    HANDLE hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);

    while (true) {
        // Read Request Header
        struct Request {
            uint64_t id;
            int32_t cx;
            int32_t cy;
            int32_t pidlLen;
        } req;

        DWORD bytesRead = 0;
        if (!ReadFile(hStdIn, &req, sizeof(req), &bytesRead, nullptr) || bytesRead != sizeof(req)) {
            break;
        }

        std::vector<BYTE> pidlBytes(req.pidlLen);
        if (req.pidlLen > 0) {
            if (!ReadFile(hStdIn, pidlBytes.data(), req.pidlLen, &bytesRead, nullptr) || bytesRead != req.pidlLen) {
                break;
            }
        }

        HBITMAP hBitmap = nullptr;
        SIZE requestSize = {req.cx, req.cy};

        if (req.pidlLen > 0) {
            Microsoft::WRL::ComPtr<IShellItem> item;
            PCIDLIST_ABSOLUTE pidl = reinterpret_cast<PCIDLIST_ABSOLUTE>(pidlBytes.data());
            if (SUCCEEDED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&item))) && item) {
                Microsoft::WRL::ComPtr<IShellItemImageFactory> factory;
                if (SUCCEEDED(item.As(&factory)) && factory) {
                    HRESULT hr = factory->GetImage(requestSize, SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK | SIIGBF_THUMBNAILONLY, &hBitmap);
                    if (FAILED(hr)) {
                        hr = factory->GetImage(requestSize, SIIGBF_RESIZETOFIT | SIIGBF_BIGGERSIZEOK, &hBitmap);
                    }
                    if (FAILED(hr)) {
                        factory->GetImage(requestSize, SIIGBF_ICONONLY, &hBitmap);
                    }
                }
            }
        }

        struct Response {
            uint64_t id;
            int32_t width;
            int32_t height;
            uint32_t dataSize;
        } resp = { req.id, 0, 0, 0 };

        std::vector<BYTE> pixelData;

        if (hBitmap) {
            BITMAP bmp;
            if (GetObject(hBitmap, sizeof(bmp), &bmp) > 0) {
                resp.width = bmp.bmWidth;
                resp.height = bmp.bmHeight;

                HDC hdc = GetDC(nullptr);
                BITMAPINFO bmi = {};
                bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
                bmi.bmiHeader.biWidth = bmp.bmWidth;
                bmi.bmiHeader.biHeight = -bmp.bmHeight; // top-down
                bmi.bmiHeader.biPlanes = 1;
                bmi.bmiHeader.biBitCount = 32;
                bmi.bmiHeader.biCompression = BI_RGB;

                resp.dataSize = bmp.bmWidth * bmp.bmHeight * 4;
                pixelData.resize(resp.dataSize);

                GetDIBits(hdc, hBitmap, 0, bmp.bmHeight, pixelData.data(), &bmi, DIB_RGB_COLORS);
                ReleaseDC(nullptr, hdc);
            }
            DeleteObject(hBitmap);
        }

        DWORD bytesWritten = 0;
        WriteFile(hStdOut, &resp, sizeof(resp), &bytesWritten, nullptr);
        if (resp.dataSize > 0 && !pixelData.empty()) {
            WriteFile(hStdOut, pixelData.data(), resp.dataSize, &bytesWritten, nullptr);
        }
    }

    CoUninitialize();
    return 0;
}
