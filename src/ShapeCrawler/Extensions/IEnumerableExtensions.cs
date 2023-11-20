#if NETSTANDARD2_0

using System;
using System.Collections.Generic;

namespace ShapeCrawler.Extensions;

internal static class IEnumerableExtensions
{
    internal static IEnumerable<TSource> DistinctBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
    {
        var seenKeys = new HashSet<TKey>();
        foreach (TSource element in source)
        {
            if (seenKeys.Add(keySelector(element)))
            {
                yield return element;
            }
        }
    }
}

#endif