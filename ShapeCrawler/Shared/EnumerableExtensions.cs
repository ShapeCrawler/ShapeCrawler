using System;
using System.Collections.Generic;

namespace ShapeCrawler.Shared
{
    internal static class EnumerableExtensions
    {
        /// <summary>
        ///     Creates an list from a <see cref="T:System.Collections.Generic.IEnumerable`1" /> with specifying capacity.
        /// </summary>
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
}