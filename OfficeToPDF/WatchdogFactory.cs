namespace OfficeToPDF
{
    internal interface IWatchdog
    {
        IWatchdog Start();
        void Stop();
    }

    internal class WatchdogFactory
    {
        public IWatchdog Create(object com, int timeout) =>
            timeout > 0 ? (IWatchdog)new Watchdog(com, timeout) : new NullWatchdog();

        public IWatchdog CreateStarted(object com, int timeout) =>
            Create(com, timeout).Start();
    }
}
