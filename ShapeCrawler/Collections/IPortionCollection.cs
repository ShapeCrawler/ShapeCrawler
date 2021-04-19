using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    public interface IPortionCollection : IEnumerable<IPortion>
    {
        /// <summary>
        ///     Gets the element at the specified index.
        /// </summary>
        IPortion this[int index] { get; }

        /// <summary>
        ///     Gets the number of series items in the collection.
        /// </summary>
        int Count { get; }

        void Remove(IPortion row);
        void Remove(IList<IPortion> removingPortions);
    }
}