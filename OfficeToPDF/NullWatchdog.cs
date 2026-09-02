namespace OfficeToPDF
{
    internal class NullWatchdog : IWatchdog
    {
        public IWatchdog Start() => this;
        public void Stop() { /* NOOP */ }
    }
}
