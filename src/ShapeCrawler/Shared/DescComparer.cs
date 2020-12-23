using System.Collections.Generic;

namespace ShapeCrawler.Shared
{
    /// <summary>
    /// Represents a desc comparer.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class DescComparer<T> : IComparer<T>
    {
        public int Compare(T x, T y)
        {
            if (x == null) return -1;
            if (y == null) return 1;

            return Comparer<T>.Default.Compare(y, x);
        }
    }
}