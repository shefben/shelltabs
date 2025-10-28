#include "ClassFactory.h"

#include <atomic>
#include <memory>
#include <new>

#include "Module.h"
#include "OpenFolderCommand.h"
#include "TabBand.h"
#include "CExplorerBHO.h"
#include "FtpShellFolder.h"

namespace shelltabs {

class TabBandClassFactory : public IClassFactory {
public:
    TabBandClassFactory() : m_refCount(1) { ModuleAddRef(); }
    ~TabBandClassFactory() { ModuleRelease(); }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *object = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return static_cast<ULONG>(++m_refCount);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        const ULONG count = static_cast<ULONG>(--m_refCount);
        if (count == 0) {
            delete this;
        }
        return count;
    }

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        if (outer != nullptr) {
            return CLASS_E_NOAGGREGATION;
        }

        TabBand* band = new (std::nothrow) TabBand();
        if (!band) {
            return E_OUTOFMEMORY;
        }

        HRESULT hr = band->QueryInterface(riid, object);
        band->Release();
        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL lock) override {
        if (lock) {
            ModuleAddRef();
        } else {
            ModuleRelease();
        }
        return S_OK;
    }

private:
    std::atomic<long> m_refCount;
};

HRESULT CreateTabBandClassFactory(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }

    TabBandClassFactory* factory = new (std::nothrow) TabBandClassFactory();
    if (!factory) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, object);
    factory->Release();
    return hr;
}

class BrowserHelperClassFactory : public IClassFactory {
public:
    BrowserHelperClassFactory() : m_refCount(1) { ModuleAddRef(); }
    ~BrowserHelperClassFactory() { ModuleRelease(); }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *object = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return static_cast<ULONG>(++m_refCount); }

    IFACEMETHODIMP_(ULONG) Release() override {
        const ULONG count = static_cast<ULONG>(--m_refCount);
        if (count == 0) {
            delete this;
        }
        return count;
    }

    IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (outer) {
            return CLASS_E_NOAGGREGATION;
        }

        auto helper = std::make_unique<CExplorerBHO>();
        if (!helper) {
            return E_OUTOFMEMORY;
        }

        CExplorerBHO* raw = helper.release();
        const HRESULT hr = raw->QueryInterface(riid, object);
        raw->Release();
        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL lock) override {
        if (lock) {
            ModuleAddRef();
        } else {
            ModuleRelease();
        }
        return S_OK;
    }

private:
    std::atomic<long> m_refCount;
};

HRESULT CreateBrowserHelperClassFactory(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }

    BrowserHelperClassFactory* factory = new (std::nothrow) BrowserHelperClassFactory();
    if (!factory) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, object);
    factory->Release();
    return hr;
}

}  // namespace shelltabs

namespace {

using shelltabs::ModuleAddRef;
using shelltabs::ModuleRelease;

class OpenFolderCommandClassFactory : public IClassFactory {
public:
    OpenFolderCommandClassFactory() : m_refCount(1) { ModuleAddRef(); }
    ~OpenFolderCommandClassFactory() { ModuleRelease(); }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *object = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return static_cast<ULONG>(++m_refCount); }

    IFACEMETHODIMP_(ULONG) Release() override {
        const ULONG count = static_cast<ULONG>(--m_refCount);
        if (count == 0) {
            delete this;
        }
        return count;
    }

    IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (outer) {
            return CLASS_E_NOAGGREGATION;
        }

        auto command = std::make_unique<shelltabs::OpenFolderCommand>();
        if (!command) {
            return E_OUTOFMEMORY;
        }

        shelltabs::OpenFolderCommand* raw = command.release();
        const HRESULT hr = raw->QueryInterface(riid, object);
        raw->Release();
        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL lock) override {
        if (lock) {
            ModuleAddRef();
        } else {
            ModuleRelease();
        }
        return S_OK;
    }

private:
    std::atomic<long> m_refCount;
};

}  // namespace

namespace shelltabs {

HRESULT CreateOpenFolderCommandClassFactory(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }

    OpenFolderCommandClassFactory* factory = new (std::nothrow) OpenFolderCommandClassFactory();
    if (!factory) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, object);
    factory->Release();
    return hr;
}

class FtpFolderClassFactory : public IClassFactory {
public:
    FtpFolderClassFactory() : m_refCount(1) { ModuleAddRef(); }
    ~FtpFolderClassFactory() { ModuleRelease(); }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *object = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override { return static_cast<ULONG>(++m_refCount); }

    IFACEMETHODIMP_(ULONG) Release() override {
        const ULONG count = static_cast<ULONG>(--m_refCount);
        if (count == 0) {
            delete this;
        }
        return count;
    }

    IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** object) override {
        if (!object) {
            return E_POINTER;
        }
        if (outer) {
            return CLASS_E_NOAGGREGATION;
        }

        auto folder = std::make_unique<ftp::FtpShellFolder>();
        if (!folder) {
            return E_OUTOFMEMORY;
        }

        ftp::FtpShellFolder* raw = folder.release();
        const HRESULT hr = raw->QueryInterface(riid, object);
        raw->Release();
        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL lock) override {
        if (lock) {
            ModuleAddRef();
        } else {
            ModuleRelease();
        }
        return S_OK;
    }

private:
    std::atomic<long> m_refCount;
};

}  // namespace

namespace shelltabs {

HRESULT CreateFtpFolderClassFactory(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }

    FtpFolderClassFactory* factory = new (std::nothrow) FtpFolderClassFactory();
    if (!factory) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, object);
    factory->Release();
    return hr;
}

}  // namespace shelltabs

}  // namespace shelltabs

