using System;
using System.Collections.Generic;

namespace OfficeToPDF.Tests
{
    internal static class Functional
    {
        public static Y Pipe<X, Y>(this X x, Func<X, Y> f) => f(x);

        public static void ForEach<T>(this IEnumerable<T> collection, Action<T> action)
        {
            foreach (var item in collection)
                action(item);
        }
    }
}
