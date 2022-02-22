#include "pch.h"

#include "COMInterop.h"
#include "cogetserverpid.h"
#include "dcom_h.h"

namespace COMInterop {

    #define TCP_PROTOCOL_ID 6

    DWORD OfficeApp::GetProcessId2(System::IntPtr application) {
        // https://www.apriorit.com/dev-blog/724-windows-three-ways-to-get-com-server-process-id

        // OXID Resolver server listens to TCP port 135
        // https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/service-overview-and-network-port-requirements

        //RPC_WSTR OXIDResolverStringBinding = 0;

        //RpcStringBindingComposeW(
        //    NULL,
        //    RPC_WSTR(L"ncacn_ip_tcp"),
        //    RPC_WSTR(L"127.0.0.1"),
        //    RPC_WSTR(L"135"),
        //    NULL,
        //    &OXIDResolverStringBinding
        //);

        //RPC_BINDING_HANDLE OXIDResolverBinding = 0;

        //RpcBindingFromStringBindingW(
        //    OXIDResolverStringBinding,
        //    &OXIDResolverBinding
        //);

        ////Make OXID Resolver authenticate without a password

        //RpcBindingSetOption(OXIDResolverBinding, RPC_C_OPT_BINDING_NONCAUSAL, 1);

        //RPC_SECURITY_QOS securityQualityOfServiceSettings;
        //securityQualityOfServiceSettings.Version = 1;
        //securityQualityOfServiceSettings.Capabilities = RPC_C_QOS_CAPABILITIES_MUTUAL_AUTH;
        //securityQualityOfServiceSettings.IdentityTracking = RPC_C_QOS_IDENTITY_STATIC;
        //securityQualityOfServiceSettings.ImpersonationType = RPC_C_IMP_LEVEL_IMPERSONATE;

        //RpcBindingSetAuthInfoExW(
        //    OXIDResolverBinding,
        //    RPC_WSTR(L"NT Authority\\NetworkService"),
        //    RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        //    RPC_C_AUTHN_WINNT,
        //    NULL,
        //    RPC_C_AUTHZ_NONE,
        //    &securityQualityOfServiceSettings
        //);

        //unsigned short requestedProtocols[] = { TCP_PROTOCOL_ID };

        //DUALSTRINGARRAY*  COMServerStringBindings = NULL;
        //IPID            remoteUnknownIPID = GUID_NULL;
        //DWORD           authHint = 0;

        //ResolveOxid(
        //    OXIDResolverBinding,
        //    &oxid,
        //    _countof(requestedProtocols),
        //    requestedProtocols,
        //    &COMServerStringBindings,
        //    &remoteUnknownIPID,
        //    &authHint
        //);

        return 0;
    }

    DWORD OfficeApp::GetProcessId(System::IntPtr application) {

        LPUNKNOWN iunknown = reinterpret_cast<LPUNKNOWN>(application.ToPointer());
        DWORD value;
        OXID oxid;
        HRESULT hr = CoGetServerPID(iunknown, &value, &oxid);
        if (SUCCEEDED(hr))
            return value;

        return 0;
    }

}