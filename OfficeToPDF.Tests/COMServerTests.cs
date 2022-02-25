using NUnit.Framework;
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace OfficeToPDF.Tests
{
    [TestFixture]
    public class COMServerTests
    {
        [DllImport("COMServer.dll", SetLastError = true)]
        static extern uint GetCOMProcessId(IntPtr unknown);

        [Test, Explicit("Starts and stops the Word office application")]
        public void GetProcessIdReturnsTheCorrectValue()
        {
            using (var application = new DisposableApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.Word);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("WINWORD"));
            }
        }

    }
}
