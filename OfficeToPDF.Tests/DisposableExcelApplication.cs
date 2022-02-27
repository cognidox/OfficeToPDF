using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposableExcelApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.Excel.Application _excel;

        public DisposableExcelApplication()
        {
            bool running = false;
            try { ExcelConverter.StartExcel(ref running, ref _excel); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.Excel.Application Excel => _excel;

        public void Dispose()
        {
            if (_excel == null)
                return;

            ExcelConverter.CloseExcelApplication(_excel);

            Converter.ReleaseCOMObject(_excel);
        }
    }
}
