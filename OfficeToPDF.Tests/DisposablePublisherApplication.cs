using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposablePublisherApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.Publisher.Application _publisher;

        public DisposablePublisherApplication()
        {
            bool running = false;
            try { PublisherConverter.StartPublisher(ref running, ref _publisher); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.Publisher.Application Publisher => _publisher;

        public void Dispose()
        {
            if (_publisher == null)
                return;

            PublisherConverter.ClosePublisherApplication(_publisher);

            Converter.ReleaseCOMObject(_publisher);
        }
    }
}
