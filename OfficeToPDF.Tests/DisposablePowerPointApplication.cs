using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposablePowerPointApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.PowerPoint.Application _powerPoint;

        public DisposablePowerPointApplication()
        {
            bool running = false;
            try { PowerpointConverter.StartPowerPoint(ref running, ref _powerPoint); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.PowerPoint.Application PowerPoint => _powerPoint;

        public void Dispose()
        {
            if (_powerPoint == null)
                return;

            PowerpointConverter.ClosePowerPointApplication(_powerPoint);

            Converter.ReleaseCOMObject(_powerPoint);
        }
    }
}
