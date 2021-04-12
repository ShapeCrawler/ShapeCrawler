using System.Collections.Generic;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Collections
{
    public interface IPortionCollection : IEnumerable<IPortion>
    {
        void Remove(IPortion row);
        void Remove(IList<IPortion> removingPortions);

        /// <summary>
        ///     Gets the element at the specified index.
        /// </summary>
        IPortion this[int index] { get; }

        /// <summary>
        ///     Gets the number of series items in the collection.
        /// </summary>
        int Count { get; }
    }
}