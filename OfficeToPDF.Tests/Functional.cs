using System;
using System.Collections.Generic;

namespace OfficeToPDF.Tests
{
    internal static class Functional
    {
        public static B Compose<A, B>(this A a, Func<A, B> f) => f(a);

        public static void ForEach<T>(this IEnumerable<T> collection, Action<T> action)
        {
            foreach (var item in collection)
                action(item);
        }
    }
}
