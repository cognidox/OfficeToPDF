using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposableVisioApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.Visio.InvisibleApp _visio;

        public DisposableVisioApplication()
        {
            bool running = false;
            try { VisioConverter.StartVisio(ref running, ref _visio); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.Visio.InvisibleApp Visio => _visio;

        public void Dispose()
        {
            if (_visio == null)
                return;

            VisioConverter.CloseVisioApplication(_visio);

            Converter.ReleaseCOMObject(_visio);
        }
    }
}
