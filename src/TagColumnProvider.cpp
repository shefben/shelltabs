#include "TagColumnProvider.h"

#include <Shlwapi.h>

#include <new>

#include "Guids.h"
#include "Tagging.h"

namespace shelltabs {

TagColumnProvider::TagColumnProvider() : m_refCount(1) {}

IFACEMETHODIMP TagColumnProvider::QueryInterface(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }
    if (riid == IID_IUnknown || riid == IID_IColumnProvider) {
        *object = static_cast<IColumnProvider*>(this);
        AddRef();
        return S_OK;
    }
    *object = nullptr;
    return E_NOINTERFACE;
}

IFACEMETHODIMP_(ULONG) TagColumnProvider::AddRef() { return static_cast<ULONG>(++m_refCount); }

IFACEMETHODIMP_(ULONG) TagColumnProvider::Release() {
    const ULONG count = static_cast<ULONG>(--m_refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}

IFACEMETHODIMP TagColumnProvider::Initialize(LPCSHCOLUMNINIT init) {
    if (init && init->wszFolder[0] != L'\0') {
        m_folderPath = init->wszFolder;
    } else {
        m_folderPath.clear();
    }
    return S_OK;
}

IFACEMETHODIMP TagColumnProvider::GetColumnInfo(DWORD index, SHCOLUMNINFO* info) {
    if (!info) {
        return E_POINTER;
    }
    if (index > 0) {
        return S_FALSE;
    }

    info->scid.fmtid = FMTID_ShellTabsTags;
    info->scid.pid = 0;
    info->vt = VT_BSTR;
    info->fmt = LVCFMT_LEFT;
    info->cChars = 64;
    info->csFlags = SHCOLSTATE_TYPE_STR | SHCOLSTATE_SLOW | SHCOLSTATE_SECONDARYUI;

    lstrcpynW(info->wszTitle, L"Tags", ARRAYSIZE(info->wszTitle));
    lstrcpynW(info->wszDescription, L"ShellTabs tag assignments", ARRAYSIZE(info->wszDescription));

    return S_OK;
}

namespace {
std::wstring CombinePath(const std::wstring& base, LPCWSTR file) {
    if (!file || file[0] == L'\0') {
        return {};
    }
    if (!PathIsRelativeW(file)) {
        return file;
    }
    if (base.empty()) {
        return {};
    }
    wchar_t buffer[MAX_PATH];
    if (PathCombineW(buffer, base.c_str(), file)) {
        return buffer;
    }
    std::wstring combined = base;
    if (!combined.empty() && combined.back() != L'\\' && combined.back() != L'/') {
        combined += L'\\';
    }
    combined += file;
    return combined;
}
}  // namespace

IFACEMETHODIMP TagColumnProvider::GetItemData(LPCSHCOLUMNID columnId, LPCSHCOLUMNDATA data, VARIANT* value) {
    if (!columnId || !data || !value) {
        return E_POINTER;
    }

    VariantInit(value);

    if (columnId->fmtid != FMTID_ShellTabsTags) {
        return S_FALSE;
    }

    const std::wstring path = CombinePath(m_folderPath, data->wszFile);
    if (path.empty()) {
        return S_FALSE;
    }

    const std::wstring tags = TagStore::Instance().GetTagListForPath(path);
    if (tags.empty()) {
        return S_FALSE;
    }

    value->bstrVal = SysAllocString(tags.c_str());
    if (!value->bstrVal) {
        return E_OUTOFMEMORY;
    }
    value->vt = VT_BSTR;
    return S_OK;
}

}  // namespace shelltabs

