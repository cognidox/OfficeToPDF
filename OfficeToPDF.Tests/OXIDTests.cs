using NUnit.Framework;
using System;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace OfficeToPDF.Tests
{
    [TestFixture]
    public class OXIDTests
    {
        [DllImport("OXID.dll", SetLastError = true)]
        static extern uint GetCOMProcessId(IntPtr unknown);

        [Test, Explicit("Starts and stops the Word office application")]
        public void GetProcessIdReturnsTheCorrectValue()
        {
            bool running = false;
            Application word = null;

            try
            {
                var result = WordConverter.StartWord(ref running, ref word);
                if (result == ExitCode.Success)
                {
                    IntPtr iunknown = Marshal.GetIUnknownForObject(word);

                    var processId = GetCOMProcessId(iunknown);

                    Marshal.Release(iunknown);

                    Trace.WriteLine($"Process id: {processId}");

                    if (processId == 0u)
                        Assert.Fail("Invalid process Id returned");

                    var process = Process.GetProcessById(Convert.ToInt32(processId));

                    Assert.That(process.ProcessName, Is.EqualTo("WINWORD"));
                }
            }
            finally
            {
                if (word != null && !running)
                {
                    WordConverter.CloseWordApplication(word);
                }
                WordConverter.ReleaseCOMObject(word);
            }
        }

    }
}
