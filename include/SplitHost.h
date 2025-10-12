#pragma once
#include <windows.h>
#include <wrl/client.h>
#include <atomic>
#include <memory>
#include <string>
#include <thread>
#include <vector>

#include "ExplorerPane.h"

namespace shelltabs {

class SplitHost {
public:
    static SplitHost* CreateAndAttach(HWND contentParent);
    static SplitHost* FromHwnd(HWND hwnd);
    static void DestroyIfExistsOn(HWND contentParent);

    struct CompareResult {
        std::vector<std::wstring> leftOnly;
        std::vector<std::wstring> rightOnly;
        std::vector<std::wstring> differingLeft;
        std::vector<std::wstring> differingRight;
        std::vector<std::wstring> folderDiffLeft;
        std::vector<std::wstring> folderDiffRight;
    };

    HWND Hwnd() const { return m_hwnd; }
    void SetFolders(PCIDLIST_ABSOLUTE left, PCIDLIST_ABSOLUTE right);
    void Swap();

private:
    struct AddressBarContext {
        SplitHost* host = nullptr;
        bool isLeft = false;
    };

    SplitHost() = default;
    ~SplitHost();

    static LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
    LRESULT OnCreate(HWND hwnd);
    void OnSize();
    void OnLButtonDown(int x);
    void OnMouseMove(int x);
    void OnLButtonUp();
    void LayoutChildren();
    void OnPaneNavigated(bool isLeft, const std::wstring& path);
    void UpdateAddressBar(bool isLeft, const std::wstring& path);
    bool NavigateFromAddressBar(bool isLeft);
    void ScheduleComparison();
    void ResetComparison();
    void ApplyComparisonResult(uint64_t token, std::unique_ptr<CompareResult> result);
    void InvalidatePanes();
    static LRESULT CALLBACK AddressBarSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                                   UINT_PTR id, DWORD_PTR refData);

    HWND m_hwnd = nullptr;
    HWND m_splitter = nullptr;
    HWND m_leftAddress = nullptr;
    HWND m_rightAddress = nullptr;
    int m_splitX = 400;  // default position
    bool m_drag = false;

    ExplorerPane m_left;
    ExplorerPane m_right;

    AddressBarContext m_leftAddressContext{};
    AddressBarContext m_rightAddressContext{};
    std::wstring m_leftPath;
    std::wstring m_rightPath;
    std::jthread m_compareThread;
    std::atomic<uint64_t> m_compareToken{0};
    std::atomic<uint64_t> m_latestToken{0};
};

}  // namespace shelltabs
