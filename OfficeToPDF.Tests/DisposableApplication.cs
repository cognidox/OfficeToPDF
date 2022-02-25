using Microsoft.Office.Interop.Word;
using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposableApplication : IDisposable
    {
        private readonly Application _word;

        public DisposableApplication()
        {
            bool running = false;
            try { WordConverter.StartWord(ref running, ref _word); }
            catch { /* NOOP */ }
        }

        public Application Word => _word;

        public void Dispose()
        {
            if (_word == null)
                return;

            WordConverter.CloseWordApplication(_word);

            Converter.ReleaseCOMObject(_word);
        }
    }
}
