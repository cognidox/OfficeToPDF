using NUnit.Framework;

namespace OfficeToPDF.Tests
{
    [TestFixture]
    public class WatchdogFactoryTests
    {
        [TestCase(0)]
        [TestCase(-1)]
        [TestCase(10)]
        [TestCase(100)]
        [TestCase(1000)]
        public void WhenCreateCalledThenWatchdogReturned(int timeout)
        {
            var com = new object();

            var factory = new WatchdogFactory();

            var result = factory.Create(com, timeout);

            Assert.That(result, Is.Not.Null);
        }

        [TestCase(0)]
        [TestCase(-1)]
        [TestCase(-10)]
        public void WhenCreateCalledWithInvalidTimeoutThenNullWatchdogReturned(int timeout)
        {
            var com = new object();

            var factory = new WatchdogFactory();

            var result = factory.Create(com, timeout);

            Assert.That(result, Is.InstanceOf<NullWatchdog>());
        }

        [TestCase(1)]
        [TestCase(10)]
        [TestCase(100)]
        public void WhenCreateCalledWithValidTimeoutThenWatchdogReturned(int timeout)
        {
            var com = new object();

            var factory = new WatchdogFactory();

            var result = factory.Create(com, timeout);

            Assert.That(result, Is.InstanceOf<Watchdog>());
        }
    }
}
