namespace OfficeToPDF
{
    internal interface IWatchdog
    {
        void Start();
        void Stop();
    }

    internal class WatchdogFactory
    {
        public IWatchdog Create(object com, int timeout) =>
            timeout > 0 ? (IWatchdog)new Watchdog(com, timeout) : new NullWatchdog();
    }
}
