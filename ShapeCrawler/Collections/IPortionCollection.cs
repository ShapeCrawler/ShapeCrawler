using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents collection of paragraph text portions.
    /// </summary>
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

        /// <summary>
        ///     Removes portion item from collection.
        /// </summary>
        void Remove(IPortion removingPortion);

        /// <summary>
        ///     Removes portion items from collection.
        /// </summary>
        void Remove(IList<IPortion> portions);
    }
}