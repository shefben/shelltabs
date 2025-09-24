#include "ClassFactory.h"

#include <atomic>
#include <memory>
#include <new>

#include "Module.h"
#include "TabBand.h"
#include "TagColumnProvider.h"

namespace shelltabs {

class TabBandClassFactory : public IClassFactory {
public:
    TabBandClassFactory() : m_refCount(1) { ModuleAddRef(); }
    ~TabBandClassFactory() override { ModuleRelease(); }

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

class TagColumnProviderClassFactory : public IClassFactory {
public:
    TagColumnProviderClassFactory() : m_refCount(1) { ModuleAddRef(); }
    ~TagColumnProviderClassFactory() override { ModuleRelease(); }

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

        auto provider = std::make_unique<TagColumnProvider>();
        if (!provider) {
            return E_OUTOFMEMORY;
        }

        TagColumnProvider* raw = provider.release();
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

HRESULT CreateTagColumnProviderClassFactory(REFIID riid, void** object) {
    if (!object) {
        return E_POINTER;
    }

    TagColumnProviderClassFactory* factory = new (std::nothrow) TagColumnProviderClassFactory();
    if (!factory) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, object);
    factory->Release();
    return hr;
}

}  // namespace shelltabs

