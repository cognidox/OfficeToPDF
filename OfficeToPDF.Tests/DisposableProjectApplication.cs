using System;

namespace OfficeToPDF.Tests
{
    internal sealed class DisposableProjectApplication : IDisposable
    {
        private readonly Microsoft.Office.Interop.MSProject.Application _project;

        public DisposableProjectApplication()
        {
            bool running = false;
            try { ProjectConverter.StartProject(ref running, ref _project); }
            catch { /* NOOP */ }
        }

        public Microsoft.Office.Interop.MSProject.Application Project => _project;

        public void Dispose()
        {
            if (_project == null)
                return;

            ProjectConverter.CloseProjectApplication(_project);

            Converter.ReleaseCOMObject(_project);
        }
    }
}
