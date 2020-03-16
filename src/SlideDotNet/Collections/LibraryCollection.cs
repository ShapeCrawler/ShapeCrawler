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

        protected List<T> _collectionItems;

        #endregion Fields

        /// <summary>
        /// Gets a generic enumerator that iterates through the collection.
        /// </summary>
        public IEnumerator<T> GetEnumerator() => _collectionItems.GetEnumerator();

        /// <summary>
        /// Gets an enumerator that iterates through the collection.
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator() => _collectionItems.GetEnumerator();

        /// <summary>
        /// Returns the element at the specified index.
        /// </summary>
        public T this[int index] => _collectionItems[index];

        /// <summary>
        /// Gets the number of series items in the collection.
        /// </summary>
        public int Count => _collectionItems.Count;
    }
}