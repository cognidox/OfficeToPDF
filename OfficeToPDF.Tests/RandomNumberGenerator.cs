using System;

namespace OfficeToPDF.Tests
{
    public sealed class RandomNumberGenerator
    {
        private readonly Random _random;

        public RandomNumberGenerator() =>
            _random = Ticks()
                .Pipe(Convert.ToInt32)
                .Pipe(seed => new Random(seed));

        private static long Ticks() => DateTime.Now.Ticks & 0x7fffffff;

        public int NextValue(int min, int max) => _random.Next(min, max);
    }
}
