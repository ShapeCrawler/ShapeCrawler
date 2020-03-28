using System.Collections;
using System.Collections.Generic;

namespace SlideDotNet.Collections
{
    /// <summary>
    /// An abstract library collection.
    /// </summary>
    public abstract class LibraryCollection<T> : IEnumerable<T>
    {
        #region Fields

        protected List<T> CollectionItems;

        #endregion Fields

        /// <summary>
        /// Gets the generic enumerator that iterates through the collection.
        /// </summary>
        public IEnumerator<T> GetEnumerator() => CollectionItems.GetEnumerator();

        /// <summary>
        /// Gets an enumerator that iterates through the collection.
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator() => CollectionItems.GetEnumerator();

        /// <summary>
        /// Gets the element at the specified index.
        /// </summary>
        public T this[int index] => CollectionItems[index];

        /// <summary>
        /// Gets the number of series items in the collection.
        /// </summary>
        public int Count => CollectionItems.Count;
    }
}