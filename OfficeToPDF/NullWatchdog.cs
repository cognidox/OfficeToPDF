namespace OfficeToPDF
{
    internal class NullWatchdog : IWatchdog
    {
        public void Start() { /* NOOP */ }
        public void Stop() { /* NOOP */ }
    }
}
