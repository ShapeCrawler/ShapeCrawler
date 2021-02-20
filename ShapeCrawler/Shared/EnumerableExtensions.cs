using System;
using System.Collections.Generic;

namespace ShapeCrawler.Shared
{
    internal static class EnumerableExtensions
    {
        /// <summary>
        ///     Creates an array from a <see cref="T:System.Collections.Generic.IEnumerable`1" /> with specifying capacity.
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="source"></param>
        /// <param name="capacity"></param>
        /// <returns></returns>
        internal static TSource[] ToArray<TSource>(this IEnumerable<TSource> source, int capacity)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (capacity < 0) throw new ArgumentOutOfRangeException(nameof(capacity));

            var array = new TSource[capacity];
            int i = 0;
            foreach (var item in source)
            {
                array[i++] = item;
            }

            return array;
        }

        /// <summary>
        ///     Creates an list from a <see cref="T:System.Collections.Generic.IEnumerable`1" /> with specifying capacity.
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <param name="source"></param>
        /// <param name="capacity"></param>
        /// <returns></returns>
        internal static List<TSource> ToList<TSource>(this IEnumerable<TSource> source, int capacity)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (capacity < 0) throw new ArgumentOutOfRangeException(nameof(capacity));

            var list = new List<TSource>(capacity);
            list.AddRange(source);

            return list;
        }
    }
}