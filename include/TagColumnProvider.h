#pragma once

#include <ShlObj.h>
#include <windows.h>

#include <atomic>
#include <string>

namespace shelltabs {

class TagColumnProvider : public IColumnProvider {
public:
    TagColumnProvider();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override;
    IFACEMETHODIMP_(ULONG) AddRef() override;
    IFACEMETHODIMP_(ULONG) Release() override;

    // IColumnProvider
    IFACEMETHODIMP Initialize(LPCSHCOLUMNINIT init) override;
    IFACEMETHODIMP GetColumnInfo(DWORD index, SHCOLUMNINFO* info) override;
    IFACEMETHODIMP GetItemData(LPCSHCOLUMNID columnId, LPCSHCOLUMNDATA data, VARIANT* value) override;

private:
    std::atomic<long> m_refCount;
    std::wstring m_folderPath;
};

}  // namespace shelltabs

