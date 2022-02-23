#pragma once

#include <windows.h>

#ifdef __cplusplus
extern "C" {  // only need to export C interface if
              // used by C++ source code
#endif

    extern __declspec(dllexport) DWORD __cdecl GetCOMProcessId(const LPVOID ptr);

#ifdef __cplusplus
}
#endif
