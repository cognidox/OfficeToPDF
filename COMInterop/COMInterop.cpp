#include "pch.h"

#include "COMInterop.h"
#include "cogetserverpid.h"

namespace COMInterop {

    #define OBJREF_STANDARD 0x00000001
    #define OBJREF_HANDLER 0x00000002
    #define OBJREF_CUSTOM 0x00000004
    #define OBJREF_EXTENDED 0x00000008

    typedef unsigned __int64 OXID;
    typedef unsigned __int64 OID;
    typedef GUID           IPID;

#pragma pack(push, 1)
    typedef struct tagSTDOBJREF {
        unsigned long flags;
        unsigned long cPublicRefs;
        OXID oxid;
        OID  oid;
        IPID ipid;
    } STDOBJREF;
#pragma pack(pop)

#pragma pack(push, 1)
    typedef struct tagDUALSTRINGARRAY {
        unsigned short wNumEntries;
        unsigned short wSecurityOffset;
        [size_is(wNumEntries)] unsigned short* aStringArray;
    } DUALSTRINGARRAY;
#pragma pack(pop)

#pragma pack(push, 1)
    typedef struct tagDATAELEMENT {
        GUID dataID;
        WORD cbSize;
        WORD cbRounded;
        [size_is(cbSize)] byte* Data;
    } DATAELEMENT;
#pragma pack(pop)

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


    DWORD OfficeApp::GetProcessId(System::IntPtr application) {

        //CComPtr<IStream> marshalStream;
        //CreateStreamOnHGlobal(NULL, TRUE, &marshalStream);

        //CoMarshalInterface(
        //    marshalStream, // Where to write the marshaled interface
        //    IID_IUnknown, // ID of the marshaled interface
        //    excelInterface, // The interface to be marshaled
        //    MSHCTX_INPROC, // Unmarshaling will be done in the same process
        //    NULL, // Reserved and must be NULL
        //    MSHLFLAGS_NORMAL // The data packet produced by the marshaling process will be unmarshaled in the destination process
        //);

        //HGLOBAL memoryHandleFromStream = NULL;
        //GetHGlobalFromStream(marshalStream, &memoryHandleFromStream);

        //LPOBJREF objef = reinterpret_cast <LPOBJREF> (GlobalLock(memoryHandleFromStream)); // It was originally published on https://www.apriorit.com/

        LPUNKNOWN iunknown = reinterpret_cast<LPUNKNOWN>(application.ToPointer());
        DWORD value;
        HRESULT hr = CoGetServerPID(iunknown, &value);
        if (SUCCEEDED(hr))
            return value;

        return 0;
    }

}