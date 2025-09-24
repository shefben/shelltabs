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

IFACEMETHODIMP TagColumnProvider::Initialize(LPCSHCOLUMNINIT) { return S_OK; }

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
std::wstring ResolveParsingName(LPCSHCOLUMNDATA data) {
    if (!data || !data->psf || !data->pidl) {
        return {};
    }

    STRRET strret{};
    if (FAILED(data->psf->GetDisplayNameOf(data->pidl, SHGDN_FORPARSING, &strret))) {
        return {};
    }

    PWSTR buffer = nullptr;
    if (FAILED(StrRetToStrW(&strret, data->pidl, &buffer))) {
        return {};
    }

    std::wstring result(buffer ? buffer : L"");
    if (buffer) {
        CoTaskMemFree(buffer);
    }
    return result;
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

    const std::wstring path = ResolveParsingName(data);
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

