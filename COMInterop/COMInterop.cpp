#include "pch.h"

#include "COMInterop.h"
#include "cogetserverpid.h"
#include <oxid.h>

namespace COMInterop {

    #define TCP_PROTOCOL_ID 6

    DWORD OfficeApp::GetProcessId2(System::IntPtr application) {
        // https://www.apriorit.com/dev-blog/724-windows-three-ways-to-get-com-server-process-id

        LPUNKNOWN punk = reinterpret_cast<LPUNKNOWN>(application.ToPointer());

        return GetCOMProcessId(punk);
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