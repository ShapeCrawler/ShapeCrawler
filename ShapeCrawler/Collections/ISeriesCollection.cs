using System.Collections.Generic;
using ShapeCrawler.Charts;

namespace ShapeCrawler.Collections
{
    public interface ISeriesCollection
    {
        /// <summary>
        ///     Gets the element at the specified index.
        /// </summary>
        Series this[int index] { get; }

        /// <summary>
        ///     Gets the number of series items in the collection.
        /// </summary>
        int Count { get; }

        /// <summary>
        ///     Gets the generic enumerator that iterates through the collection.
        /// </summary>
        IEnumerator<Series> GetEnumerator();
    }
}