/*******************************************************************************
Copyright (c) 2012, Kim Gr�sman
All rights reserved.

Released under the Modified BSD license. For details, please see LICENSE file.

https://github.com/kimgr/cogetserverpid

*******************************************************************************/
#ifndef INCLUDED_COGETSERVERPID_H__
#define INCLUDED_COGETSERVERPID_H__

#include <objbase.h>
#include "dcom_h.h"

// https://www.apriorit.com/dev-blog/724-windows-three-ways-to-get-com-server-process-id

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

inline HRESULT CoGetServerPID(IUnknown* punk, DWORD* pdwPID, OXID* oxid)
{
    HRESULT hr;
    IUnknown* pProxyManager = NULL;
    IStream* pMarshalStream = NULL;
    HGLOBAL hg = NULL;
    OBJREF* pObjRef = NULL;
    LARGE_INTEGER zero = { 0 };

    if (pdwPID == NULL) return E_POINTER;
    if (punk == NULL) return E_INVALIDARG;

    /* Make sure this is a standard proxy, otherwise we can't make any
       assumptions about OBJREF wire format. */
    hr = punk->QueryInterface(IID_IProxyManager, (void**)&pProxyManager);
    if (FAILED(hr)) return hr;

    pProxyManager->Release();

    /* Marshal the interface to get a new OBJREF. */
    hr = ::CreateStreamOnHGlobal(NULL, TRUE, &pMarshalStream);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ::CoMarshalInterface(pMarshalStream, IID_IUnknown, punk, MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL);
    if (FAILED(hr))
    {
        return hr;
    }

    /* We just created the stream so it's safe to go back to a raw pointer. */
    hr = ::GetHGlobalFromStream(pMarshalStream, &hg);
    if (SUCCEEDED(hr))
    {
        /* Start out pessimistic. */
        hr = RPC_E_INVALID_OBJREF;

        pObjRef = (OBJREF*)GlobalLock(hg);
        if (pObjRef != NULL)
        {
            /* Validate what we can. */
            if (pObjRef->signature == OBJREF_SIGNATURE) /* 'MEOW' */
            {
                IPID ipid;

                switch (pObjRef->flags)
                {
                case OBJREF_STANDARD:
                    ipid = pObjRef->u_objref.u_standard.std.ipid;
                    *oxid = pObjRef->u_objref.u_standard.std.oxid;
                    break;
                case OBJREF_HANDLER:
                    ipid = pObjRef->u_objref.u_handler.std.ipid;
                    *oxid = pObjRef->u_objref.u_handler.std.oxid;
                    break;
                case OBJREF_EXTENDED:
                    ipid = pObjRef->u_objref.u_extended.std.ipid;
                    *oxid = pObjRef->u_objref.u_extended.std.oxid;
                    break;
                default:
                    ipid = GUID_NULL;
                    *oxid = 0;
                    break;
                }

                if (GetCOMServerPID(ipid, pdwPID))
                {
                    hr = S_OK;
                }
            }

            GlobalUnlock(hg);
        }
    }

    /* Rewind stream and release marshal data to keep refcount in order. */
    HRESULT ignore;
    ignore = pMarshalStream->Seek(zero, SEEK_SET, NULL);
    ignore = CoReleaseMarshalData(pMarshalStream);

    pMarshalStream->Release();

    return hr;
}

#endif // INCLUDED_COGETSERVERPID_H__
