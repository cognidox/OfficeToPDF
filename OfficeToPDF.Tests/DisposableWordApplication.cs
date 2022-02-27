using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposableWordApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.Word.Application _word;

        public DisposableWordApplication()
        {
            bool running = false;
            try { WordConverter.StartWord(ref running, ref _word); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.Word.Application Word => _word;

        public void Dispose()
        {
            if (_word == null)
                return;

            WordConverter.CloseWordApplication(_word);

            Converter.ReleaseCOMObject(_word);
        }
    }
}
