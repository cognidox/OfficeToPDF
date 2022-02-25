using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeToPDF
{
    internal class Watchdog : IWatchdog
    {
        private readonly TimeSpan _timeout;
        private readonly object _com;
        private ManualResetEvent _event;
        private int _triggered;

        public Watchdog(object com, int timeout) =>
            (_com, _timeout) = (com, TimeSpan.FromSeconds(timeout));

        public IWatchdog Start()
        {
            if (!IsComObject())
                return this;

            ProcessId = GetProcessIdForComServer();

            _event = new ManualResetEvent(false);

            StartBackgroundThread();

            return this;
        }

        private void StartBackgroundThread()
        {
            var background = new Thread(Work);

            background.Start();
        }

        public uint ProcessId { get; private set; }

        public bool Triggered => Interlocked.Exchange(ref _triggered, _triggered) != 0;

        private void Work()
        {
            var signalled = _event.WaitOne(_timeout); // Blocks until timeout of event signalled
            Interlocked.Exchange(ref _triggered, signalled ? 0 : 1); // No support for bool so use int.
            if (signalled)
                return; // Event was signalled so don't kill COM server

            // Timeout, so kill COM server ...

            var process = TryGetProcessById(ProcessId);

            process?.Kill();
        }

        private static Process TryGetProcessById(uint processId)
        {
            try { return Process.GetProcessById(Convert.ToInt32(processId)); }
            catch { return null; }
        }

        private bool IsComObject() => Marshal.IsComObject(_com);

        public void Stop()
        {
            if (!IsComObject())
                return; // Never started

            Cleanup();
        }

        private void Cleanup()
        {
            _event?.Set(); // Will result in background thread exiting without killing process
            _event?.Dispose();
            _event = null;
        }

        [DllImport("COMServer.dll", SetLastError = true)]
        private static extern uint GetCOMProcessId(IntPtr unknown);

        private uint GetProcessIdForComServer()
        {
            IntPtr punk = Marshal.GetIUnknownForObject(_com);

            var processId = GetCOMProcessId(punk);

            Marshal.Release(punk);

            return processId;
        }
    }
}
