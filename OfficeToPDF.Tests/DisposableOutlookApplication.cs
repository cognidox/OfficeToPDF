using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposableOutlookApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _outlook;

        public DisposableOutlookApplication()
        {
            bool running = false;
            try { OutlookConverter.StartOutlook(ref running, ref _outlook); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.Outlook.Application Outlook => _outlook;

        public void Dispose()
        {
            if (_outlook == null)
                return;

            OutlookConverter.CloseOutlookApplication(_outlook);

            Converter.ReleaseCOMObject(_outlook);
        }
    }
}
