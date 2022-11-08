using System;
using System.Collections.Generic;

namespace ShapeCrawler.Extensions;

internal static class EnumerableExtensions
{
    internal static List<TSource> ToList<TSource>(this IEnumerable<TSource> source, int capacity)
    {
        if (source == null)
        {
            throw new ArgumentNullException(nameof(source));
        }

        if (capacity < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(capacity));
        }

        var list = new List<TSource>(capacity);
        list.AddRange(source);

        return list;
    }
}