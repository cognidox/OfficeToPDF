using System;

namespace OfficeToPDF.Tests
{
    public sealed class RandomNumberGenerator
    {
        private readonly Random _random;

        public RandomNumberGenerator() =>
            _random = (DateTime.Now.Ticks & 0x7fffffff)
                .Compose(Convert.ToInt32)
                .Compose(seed => new Random(seed));

        public int NextValue(int min, int max) => _random.Next(min, max);
    }
}
