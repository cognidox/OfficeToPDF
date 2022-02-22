// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "CoGetServerOxid.h"
#include "oxid.h"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

// See https://go.microsoft.com/fwlink/?LinkId=89824 [table I-2 Appendix I.] for protocol ids
// https://publications.opengroup.org/downloadable/download/link/id/MC4yMDI3MTUwMCAxNjQ1NTQ2MjI1MTMyNTcyMDEzNTU5MDM2OTk%2C/

#define TCP_PROTOCOL_ID 7

__declspec(dllexport) DWORD __cdecl GetCOMProcessId(LPVOID ptr) {

    // Source: https://www.apriorit.com/dev-blog/724-windows-three-ways-to-get-com-server-process-id

    OXID oxid;
    HRESULT hr = CoGetServerOXID(reinterpret_cast<LPUNKNOWN>(ptr), &oxid);
    if (FAILED(hr)) return 0;


    // OXID Resolver server listens to TCP port 135
    // https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/service-overview-and-network-port-requirements

    RPC_WSTR OXIDResolverStringBinding = 0;

    if (RpcStringBindingComposeW(
        NULL,
        RPC_WSTR(L"ncacn_ip_tcp"),
        RPC_WSTR(L"127.0.0.1"),
        RPC_WSTR(L"135"),
        NULL,
        &OXIDResolverStringBinding
    ) != RPC_S_OK) return 0;


    RPC_BINDING_HANDLE OXIDResolverBinding = 0;

    if (RpcBindingFromStringBindingW(
        OXIDResolverStringBinding,
        &OXIDResolverBinding
    ) != RPC_S_OK) return 0;


    //Make OXID Resolver authenticate without a password

    if (RpcBindingSetOption(OXIDResolverBinding, RPC_C_OPT_BINDING_NONCAUSAL, 1) != RPC_S_OK) return 0;

    RPC_SECURITY_QOS securityQualityOfServiceSettings;
    securityQualityOfServiceSettings.Version = 1;
    securityQualityOfServiceSettings.Capabilities = RPC_C_QOS_CAPABILITIES_MUTUAL_AUTH;
    securityQualityOfServiceSettings.IdentityTracking = RPC_C_QOS_IDENTITY_STATIC;
    securityQualityOfServiceSettings.ImpersonationType = RPC_C_IMP_LEVEL_IMPERSONATE;

    if (RpcBindingSetAuthInfoExW(
        OXIDResolverBinding,
        RPC_WSTR(L"NT Authority\\NetworkService"),
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_AUTHN_WINNT,
        NULL,
        RPC_C_AUTHZ_NONE,
        &securityQualityOfServiceSettings
    ) != RPC_S_OK) return 0;


    unsigned short requestedProtocols[] = { TCP_PROTOCOL_ID };

    DUALSTRINGARRAY* COMServerStringBindings = NULL;
    IPID            remoteUnknownIPID = GUID_NULL;
    DWORD           authHint = 0;


    if (ResolveOxid(
        OXIDResolverBinding,
        &oxid,
        _countof(requestedProtocols),
        requestedProtocols,
        &COMServerStringBindings,
        &remoteUnknownIPID,
        &authHint
    ) == ERROR_SUCCESS)
    {
        if (COMServerStringBindings != NULL)
        {
            unsigned short* p = COMServerStringBindings->aStringArray;
            for (size_t index = 0; index < COMServerStringBindings->wNumEntries;)
            {
                unsigned short protocol = p[index];
                if (index + 2 >= COMServerStringBindings->wNumEntries)
                    break;

                wchar_t* netaddr = reinterpret_cast<wchar_t*>(p + index + 1);
                if (protocol == TCP_PROTOCOL_ID)
                {
                    wchar_t copy[256] = { 0 };
                    wcsncpy_s(copy, netaddr, 255);
                    wchar_t* ctx = NULL;
                    wchar_t* token = wcstok_s(copy, L"[]", &ctx);
                    while (token != NULL)
                    {
                        token = wcstok_s(NULL, L"[]", &ctx);
                    }

                }
                index += wcslen(netaddr) + 2;
            }
        }
    }

    return 0;
}
