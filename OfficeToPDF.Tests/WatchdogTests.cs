using NUnit.Framework;
using System.Threading;
using System;
using System.Linq;

namespace OfficeToPDF.Tests
{
    [TestFixture]
    public class WatchdogTests
    {
        [Test, Explicit("Starts and stops word")]
        public void WhenWatchdogTimesoutsThenItIsTriggered()
        {
            using (var application = new DisposableApplication())
            {
                var watchdog = new Watchdog(application.Word, timeout: 5);

                watchdog.Start();

                Thread.Sleep(TimeSpan.FromSeconds(10));

                watchdog.Stop();

                Assert.That(watchdog.Triggered, Is.True);
            }
        }

        [Test, Explicit("Starts and stops word")]
        public void WhenWatchdogDoesNotTimesoutThenItIsNotTriggered()
        {
            using (var application = new DisposableApplication())
            {
                var watchdog = new Watchdog(application.Word, timeout: 10);

                watchdog.Start();

                Thread.Sleep(TimeSpan.FromSeconds(5));

                watchdog.Stop();

                Assert.That(watchdog.Triggered, Is.False);
            }
        }

        [Test, Explicit("Starts and stops word")]
        public void WhenWatchdogStoppedThenNoExceptionsThown()
        {
            using (var application = new DisposableApplication())
            {
                var watchdog = new Watchdog(application.Word, timeout: 5);

                Enumerable.Range(0, RandomValue(3, 5))
                    .ForEach(_ => watchdog.Stop());

                Assert.That(watchdog.Triggered, Is.False);
            }
        }

        private static RandomNumberGenerator Random { get; } = new RandomNumberGenerator();

        private static int RandomValue(int min, int max) => Random.NextValue(min, max);
    }
}
