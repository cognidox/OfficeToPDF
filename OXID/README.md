# COM Interop

This project contains code that is responsible for locating the process id of the MsOffice Application COM Server.

It use the IUnknown interface to locate the process id. A single function is exposed by the assembly which can be called
from C# code via the standard pinvoke mechanism.

```csharp
using System;
using System.Runtime.InteropServices;

[DllImport("OXID.dll", SetLastError = true)]
static extern uint GetCOMProcessId(IntPtr unknown);
```

```csharp
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

var word = new Microsoft.Office.Interop.Word.Application();

IntPtr punk = Marshal.GetIUnknownForObject(word);

var processId = GetCOMProcessId(punk);

Marshal.Release(punk);

Trace.WriteLine($"Process id: {processId}");
```

The code is based on a number of sources found on the web, namely:

* [Three ways to get the com server process id](https://www.apriorit.com/dev-blog/724-windows-three-ways-to-get-com-server-process-id)
* [Stackoverflow - get process id of com server](https://stackoverflow.com/questions/5046433/get-process-id-of-com-server)
* [Github - kimgr/cogetserverpid](https://github.com/kimgr/cogetserverpid)

Other sourcee of information helped with how the OXID Resolver works:

* [The oxid resolver - Part 1](https://airbus-cyber-security.com/the-oxid-resolver-part-1-remote-enumeration-of-network-interfaces-without-any-authentication/)
* [The oxid resolver - Part 2](https://airbus-cyber-security.com/the-oxid-resolver-part-2-accessing-a-remote-object-inside-dcom/)
* [IObjectExporter](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-dcom/49aef5a4-f0ad-4478-abb5-cb9446dc13c6)
* [Build your own netstat](https://timvw.be/2007/09/09/build-your-own-netstat.exe-with-c/)

A number of documents are referenced by the MS-DCOM documentation and blog posts:

* [[MS-DCOM]-171201](../[MS-DCOM]-171201.pdf)
* [C706](../c706.pdf)
