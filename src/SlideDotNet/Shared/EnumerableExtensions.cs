using System;
using System.Collections.Generic;

namespace SlideDotNet.Shared
{
    public static class EnumerableExtensions
    {
        public static List<T> ToList<T>(this IEnumerable<T> source, int capacity)
        {
            Check.NotNull(source, nameof(source));
            if (capacity < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(capacity));
            }

            var list = new List<T>(capacity);
            list.AddRange(source);

            return list;
        }

        public static T[] ToArray<T>(this IEnumerable<T> source, int capacity)
        {
            Check.NotNull(source, nameof(source));
            if (capacity < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(capacity));
            }

            var array = new T[capacity];
            int i = 0;
            foreach (var item in source)
            {
                array[i++] = item;
            }

            return array;
        }
    }
}
