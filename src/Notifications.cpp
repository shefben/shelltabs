#include "Notifications.h"

#include "Logging.h"

#include <inspectable.h>
#include <windows.data.xml.dom.h>
#include <windows.ui.notifications.h>
#include <wrl/client.h>
#include <wrl/wrappers/corewrappers.h>
#include <roapi.h>
#include <shobjidl_core.h>

#include <atomic>
#include <string>

namespace {

using Microsoft::WRL::ComPtr;
using Microsoft::WRL::Wrappers::HStringReference;
using Microsoft::WRL::Wrappers::RoInitializeWrapper;

constexpr wchar_t kShellTabsToastAppId[] = L"ShellTabs.Automation";
constexpr wchar_t kToastTitle[] = L"ShellTabs activation blocked";

std::atomic<bool> g_toastIssued{false};

HRESULT CreateToastDocument(const std::wstring& body,
                            ABI::Windows::Data::Xml::Dom::IXmlDocument** document) noexcept {
    if (!document) {
        return E_POINTER;
    }

    *document = nullptr;

    RoInitializeWrapper initializer(RO_INIT_MULTITHREADED);
    const HRESULT initHr = initializer.HResult();
    if (FAILED(initHr) && initHr != RO_E_INITIALIZED) {
        return initHr;
    }

    ComPtr<IInspectable> inspectable;
    HRESULT hr = RoActivateInstance(HStringReference(RuntimeClass_Windows_Data_Xml_Dom_XmlDocument).Get(),
                                    &inspectable);
    if (FAILED(hr)) {
        return hr;
    }

    ComPtr<ABI::Windows::Data::Xml::Dom::IXmlDocument> xmlDocument;
    hr = inspectable.As(&xmlDocument);
    if (FAILED(hr)) {
        return hr;
    }

    ComPtr<ABI::Windows::Data::Xml::Dom::IXmlDocumentIO> xmlDocumentIo;
    hr = xmlDocument.As(&xmlDocumentIo);
    if (FAILED(hr) || !xmlDocumentIo) {
        return hr;
    }

    std::wstring xml = L"<toast scenario='reminder'><visual><binding template='ToastGeneric'><text>";
    xml.append(kToastTitle);
    xml.append(L"</text><text>");
    xml.append(body);
    xml.append(L"</text></binding></visual></toast>");

    hr = xmlDocumentIo->LoadXml(HStringReference(xml.c_str()).Get());
    if (FAILED(hr)) {
        return hr;
    }

    *document = xmlDocument.Detach();
    return S_OK;
}

std::wstring BuildToastBody(HRESULT hr) {
    wchar_t buffer[256];
    swprintf(buffer, ARRAYSIZE(buffer),
             L"Windows policies prevented ShellTabs from enabling automation (0x%08X)."
             L" Open Settings > Privacy & security > Automation to re-enable.",
             hr);
    return std::wstring(buffer);
}

}  // namespace

namespace shelltabs {

bool NotifyAutomationDisabledByPolicy(HRESULT hr) noexcept {
    if (g_toastIssued.exchange(true)) {
        return false;
    }

    try {
        std::wstring body = BuildToastBody(hr);

        HRESULT appIdHr = SetCurrentProcessExplicitAppUserModelID(kShellTabsToastAppId);
        if (FAILED(appIdHr)) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: failed to set AUMID (hr=0x%08X)", appIdHr);
        }

        ComPtr<ABI::Windows::Data::Xml::Dom::IXmlDocument> xmlDocument;
        HRESULT hrXml = CreateToastDocument(body, &xmlDocument);
        if (FAILED(hrXml) || !xmlDocument) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: toast XML creation failed (hr=0x%08X)", hrXml);
            return false;
        }

        RoInitializeWrapper initializer(RO_INIT_MULTITHREADED);
        HRESULT initHr = initializer.HResult();
        if (FAILED(initHr) && initHr != RO_E_INITIALIZED) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: RoInitialize failed (hr=0x%08X)", initHr);
            return false;
        }

        ComPtr<ABI::Windows::UI::Notifications::IToastNotificationManagerStatics> toastStatics;
        HRESULT toastHr = RoGetActivationFactory(
            HStringReference(RuntimeClass_Windows_UI_Notifications_ToastNotificationManager).Get(),
            IID_PPV_ARGS(&toastStatics));
        if (FAILED(toastHr) || !toastStatics) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: ToastNotificationManager unavailable (hr=0x%08X)",
                       toastHr);
            return false;
        }

        ComPtr<ABI::Windows::UI::Notifications::IToastNotifier> notifier;
        toastHr = toastStatics->CreateToastNotifierWithId(HStringReference(kShellTabsToastAppId).Get(),
                                                          &notifier);
        if (FAILED(toastHr) || !notifier) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: failed to create toast notifier (hr=0x%08X)", toastHr);
            return false;
        }

        ComPtr<ABI::Windows::UI::Notifications::IToastNotificationFactory> toastFactory;
        toastHr = RoGetActivationFactory(HStringReference(RuntimeClass_Windows_UI_Notifications_ToastNotification).Get(),
                                         IID_PPV_ARGS(&toastFactory));
        if (FAILED(toastHr) || !toastFactory) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: toast factory unavailable (hr=0x%08X)", toastHr);
            return false;
        }

        ComPtr<ABI::Windows::UI::Notifications::IToastNotification> toast;
        toastHr = toastFactory->CreateToastNotification(xmlDocument.Get(), &toast);
        if (FAILED(toastHr) || !toast) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: failed to create toast notification (hr=0x%08X)",
                       toastHr);
            return false;
        }

        toastHr = notifier->Show(toast.Get());
        if (FAILED(toastHr)) {
            LogMessage(LogLevel::Warning,
                       L"NotifyAutomationDisabledByPolicy: toast display failed (hr=0x%08X)", toastHr);
            return false;
        }

        LogMessage(LogLevel::Info,
                   L"NotifyAutomationDisabledByPolicy: toast displayed to explain automation policy block");
        return true;

    } catch (...) {
        LogMessage(LogLevel::Warning, L"NotifyAutomationDisabledByPolicy: toast flow aborted by exception");
        return false;
    }
}

}  // namespace shelltabs
