#pragma once

#include "dcom_h.h"

// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-dcom/fe6c5e46-adf8-4e34-a8de-3f756c875f31

#define OBJREF_STANDARD 0x00000001
#define OBJREF_HANDLER 0x00000002
#define OBJREF_CUSTOM 0x00000004
#define OBJREF_EXTENDED 0x00000008
#define OBJREF_SIGNATURE 0x574f454d

#pragma pack(push, 1)
typedef struct tagOBJREF {
    unsigned long signature;
    unsigned long flags;
    GUID        iid;
    union {
        struct {
            STDOBJREF     std;
            DUALSTRINGARRAY saResAddr;
        } u_standard;
        struct {
            STDOBJREF     std;
            CLSID         clsid;
            DUALSTRINGARRAY saResAddr;
        } u_handler;
        struct {
            CLSID         clsid;
            unsigned long   cbExtension;
            unsigned long   size;
            byte* pData;
        } u_custom;
        struct {
            STDOBJREF     std;
            unsigned long   Signature1;
            DUALSTRINGARRAY saResAddr;
            unsigned long   nElms;
            unsigned long   Signature2;
            DATAELEMENT   ElmArray;
        } u_extended;
    } u_objref;
} OBJREF, * LPOBJREF;
#pragma pack(pop)


inline BOOL GetCOMServerPID(__in IPID ipid, __out DWORD* pid)
{
    static const int COM_SERVER_PID_OFFSET = 4;

    *pid = *reinterpret_cast<LPWORD>(
        (reinterpret_cast<LPBYTE>(&ipid) + COM_SERVER_PID_OFFSET)
        );

    // IPID contains only 16 - bit for PID, and if the PID > 0xffff, then it's clapped to 0xffff
    return *pid != 0xffff;
}

// Based on https://github.com/kimgr/cogetserverpid

inline HRESULT EnsureStandardProxy(LPUNKNOWN punk)
{
    /* Make sure this is a standard proxy, otherwise we can't make any
       assumptions about OBJREF wire format. */
    IUnknown* pProxyManager = NULL;
    HRESULT hr = punk->QueryInterface(IID_IProxyManager, (void**)&pProxyManager);
    if (SUCCEEDED(hr))
        pProxyManager->Release();

    return hr;
}

inline HRESULT CoGetServerObjRefInfo(LPUNKNOWN punk, OXID* oxid, IPID* ipid)
{
    if (punk == NULL) return E_INVALIDARG;
    if (oxid == NULL) return E_POINTER;
    if (ipid == NULL) return E_POINTER;

    /* Make sure this is a standard proxy, otherwise we can't make any
       assumptions about OBJREF wire format. */
    HRESULT hr = EnsureStandardProxy(punk);
    if (FAILED(hr)) return hr;

    /* Marshal the interface to get a new OBJREF. */
    IStream* pMarshalStream = NULL;
    hr = ::CreateStreamOnHGlobal(NULL, TRUE, &pMarshalStream);
    if (FAILED(hr)) return hr;

    hr = ::CoMarshalInterface(pMarshalStream, IID_IUnknown, punk, MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL);
    if (FAILED(hr)) return hr;

    /* We just created the stream so it's safe to go back to a raw pointer. */
    HGLOBAL hg = NULL;
    hr = ::GetHGlobalFromStream(pMarshalStream, &hg);
    if (SUCCEEDED(hr))
    {
        /* Start out pessimistic. */
        hr = RPC_E_INVALID_OBJREF;

        OBJREF* pObjRef = (OBJREF*)GlobalLock(hg);
        if (pObjRef != NULL)
        {
            /* Validate what we can. */
            if (pObjRef->signature == OBJREF_SIGNATURE) /* 'MEOW' */
            {
                switch (pObjRef->flags)
                {
                case OBJREF_STANDARD:
                    hr = S_OK;
                    *oxid = pObjRef->u_objref.u_standard.std.oxid;
                    *ipid = pObjRef->u_objref.u_standard.std.ipid;
                    break;
                case OBJREF_HANDLER:
                    hr = S_OK;
                    *oxid = pObjRef->u_objref.u_handler.std.oxid;
                    *ipid = pObjRef->u_objref.u_handler.std.ipid;
                    break;
                case OBJREF_EXTENDED:
                    hr = S_OK;
                    *oxid = pObjRef->u_objref.u_extended.std.oxid;
                    *ipid = pObjRef->u_objref.u_extended.std.ipid;
                    break;
                default:
                    *oxid = 0;
                    break;
                }
            }

            GlobalUnlock(hg);
        }
    }

    /* Rewind stream and release marshal data to keep refcount in order. */
    LARGE_INTEGER zero = { 0 };
    HRESULT _ = pMarshalStream->Seek(zero, SEEK_SET, NULL);
    _ = CoReleaseMarshalData(pMarshalStream);

    pMarshalStream->Release();

    return hr;
}

