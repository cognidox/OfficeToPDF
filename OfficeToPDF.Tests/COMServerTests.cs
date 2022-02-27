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
        public void GetProcessIdForWordReturnsTheCorrectValue()
        {
            using (var application = new DisposableWordApplication())
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

        [Test, Explicit("Starts and stops the Excel office application")]
        public void GetProcessIdForExcelReturnsTheCorrectValue()
        {
            using (var application = new DisposableExcelApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.Excel);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("EXCEL"));
            }
        }

        [Test, Explicit("Starts and stops the PowerPoint office application")]
        public void GetProcessIdForPowerPointReturnsTheCorrectValue()
        {
            using (var application = new DisposablePowerPointApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.PowerPoint);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("POWERPNT"));
            }
        }

        [Test, Explicit("Starts and stops the Outlook office application")]
        public void GetProcessIdForOutlookReturnsTheCorrectValue()
        {
            using (var application = new DisposableOutlookApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.Outlook);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("OUTLOOK"));
            }
        }

        [Test, Explicit("Starts and stops the Project office application")]
        public void GetProcessIdForProjectReturnsTheCorrectValue()
        {
            using (var application = new DisposableProjectApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.Project);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("WINPROJ"));
            }
        }

        [Test, Explicit("Starts and stops the Visio office application")]
        public void GetProcessIdForVisioReturnsTheCorrectValue()
        {
            using (var application = new DisposableVisioApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.Visio);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("VISIO"));
            }
        }

        [Test, Explicit("Starts and stops the Publisher office application")]
        public void GetProcessIdForPublisherReturnsTheCorrectValue()
        {
            using (var application = new DisposablePublisherApplication())
            {
                IntPtr iunknown = Marshal.GetIUnknownForObject(application.Publisher);

                var processId = GetCOMProcessId(iunknown);

                Marshal.Release(iunknown);

                Trace.WriteLine($"Process id: {processId}");

                if (processId == 0u)
                    Assert.Fail("Invalid process Id returned");

                var process = Process.GetProcessById(Convert.ToInt32(processId));

                Assert.That(process.ProcessName, Is.EqualTo("MSPUB"));
            }
        }
    }
}
