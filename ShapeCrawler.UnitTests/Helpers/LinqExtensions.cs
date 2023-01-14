using System;
using System.Collections.Generic;

namespace ShapeCrawler.Tests.Helpers;

public static class LinqExtensions
{
    public static IEnumerable<TSource> DistinctByCustom<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
    {
        HashSet<TKey> seenKeys = new HashSet<TKey>();
        foreach (TSource element in source)
        {
            if (seenKeys.Add(keySelector(element)))
            {
                yield return element;
            }
        }
    }
}