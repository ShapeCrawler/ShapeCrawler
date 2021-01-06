namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents a base class for editable collection.
    /// </summary>
    public abstract class EditableCollection<T> : LibraryCollection<T>
    {
        /// <summary>
        /// Removes the specific object from the collection.
        /// </summary>
        public abstract void Remove(T item);
    }
}