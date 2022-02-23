// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "CoGetServerObjRefInfo.h"
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

BOOL GetOXIDResolverBinding(RPC_BINDING_HANDLE& handle) {
    // OXID Resolver server listens to TCP port 135
    // https://docs.microsoft.com/en-us/troubleshoot/windows-server/networking/service-overview-and-network-port-requirements

    handle = 0;

    RPC_WSTR OXIDResolverStringBinding = 0;

    if (RpcStringBindingComposeW(
        NULL,
        RPC_WSTR(L"ncacn_ip_tcp"),
        RPC_WSTR(L"127.0.0.1"),
        RPC_WSTR(L"135"),
        NULL,
        &OXIDResolverStringBinding
    ) != RPC_S_OK) return FALSE;


    RPC_BINDING_HANDLE OXIDResolverBinding = 0;

    if (RpcBindingFromStringBindingW(
        OXIDResolverStringBinding,
        &OXIDResolverBinding
    ) != RPC_S_OK) return FALSE;


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
    ) != RPC_S_OK) return FALSE;

    handle = OXIDResolverBinding;

    return TRUE;
}

std::wstring GetPort(const wchar_t* netaddr) {

    std::list<std::wstring> ports;
    wchar_t copy[256];
    memset(copy, 0, sizeof(copy));

    size_t number = (sizeof(copy) / sizeof(copy[0])) - 1;
    wcsncpy_s(copy, netaddr, number);
    
    wchar_t* ctx = NULL;
    wchar_t* token = wcstok_s(copy, L"[]", &ctx);
    while (token != NULL)
    {
        token = wcstok_s(NULL, L"[]", &ctx);
        if (token != NULL)
            ports.push_back(std::wstring(token));
    }

    return ports.size() > 0 ? ports.front() : std::wstring();
}

#define TABLE_CLASS TCP_TABLE_CLASS::TCP_TABLE_OWNER_PID_ALL
#define TABLE_TYPE MIB_TCPTABLE_OWNER_PID
#define TABLE_ENTRY MIB_TCPROW_OWNER_PID
#define ALLOC(x) (TABLE_TYPE*)HeapAlloc(GetProcessHeap(), 0, (x))
#define FREE(x) HeapFree(GetProcessHeap(), 0, (x))


DWORD GetProcessIdFromPort(DWORD localport) {
    ULONG size = 0;
    TABLE_TYPE* tcpTable = NULL;

    DWORD result = GetExtendedTcpTable(tcpTable, &size, TRUE, AF_INET, TABLE_CLASS, 0);
    if (result != ERROR_INSUFFICIENT_BUFFER) return 0;

    tcpTable = ALLOC(size);
    if (tcpTable == NULL) return 0;

    std::list<TABLE_ENTRY> entries;

    DWORD dwNumEntries = 0;

    if ((result = GetExtendedTcpTable(tcpTable, &size, TRUE, AF_INET, TABLE_CLASS, 0)) == NO_ERROR) {
        dwNumEntries = tcpTable->dwNumEntries;
        for (size_t index = 0; index < dwNumEntries; ++index) {
            TABLE_ENTRY* value = tcpTable->table + index;
            TABLE_ENTRY entry = *value;
            // The local port number in network byte order for the TCP connection on the local computer.
            byte* values = reinterpret_cast<byte*>(&(value->dwLocalPort));
            DWORD portValue = (values[0] << 8) + values[1] + (values[2] << 24) + (values[3] << 16);
            entry.dwLocalPort = portValue;
            entries.push_back(entry);
        }
    }

    FREE(tcpTable);

    for (auto it = std::begin(entries); it != std::end(entries); ++it) {
        if ((*it).dwLocalPort == localport)
            return (*it).dwOwningPid;
    }

    return 0;
}

__declspec(dllexport) DWORD __cdecl GetCOMProcessId(LPVOID ptr) {

    // Source: https://www.apriorit.com/dev-blog/724-windows-three-ways-to-get-com-server-process-id

    OXID oxid;
    IPID ipid;
    HRESULT hr = CoGetServerObjRefInfo(reinterpret_cast<LPUNKNOWN>(ptr), &oxid, &ipid);
    if (FAILED(hr)) return 0;

    DWORD pid;
    if (GetCOMServerPID(ipid, &pid))
        return pid;

    // Fallback to second way ...

    RPC_BINDING_HANDLE OXIDResolverBinding;
    
    if(!GetOXIDResolverBinding(OXIDResolverBinding)) return 0;

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
    ) != ERROR_SUCCESS) return 0;

    if (COMServerStringBindings != NULL)
    {
        std::list<std::wstring> ports;
        size_t offset = 2; // protocol + end of string
        for (size_t index = 0; index < COMServerStringBindings->wNumEntries;)
        {
            unsigned short* values = COMServerStringBindings->aStringArray + index;
            unsigned short protocol = *values++;
            if (index + offset >= COMServerStringBindings->wNumEntries)
                break;

            wchar_t* netaddr = reinterpret_cast<wchar_t*>(values);
            if (protocol == TCP_PROTOCOL_ID)
            {
                std::wstring port = GetPort(netaddr);
                if (!port.empty())
                    ports.push_back(port);
            }
            index += wcslen(netaddr) + offset;
        }

        if (ports.size() > 0) {
            DWORD localport = std::stoul(ports.front(), nullptr, 0);
            return GetProcessIdFromPort(localport);
        }
    }

    return 0;
}
