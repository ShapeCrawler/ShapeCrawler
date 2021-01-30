namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents a base class for editable collections.
    /// </summary>
    public abstract class EditableCollection<T> : LibraryCollection<T>
    {
        /// <summary>
        /// Removes the specific object from collection.
        /// </summary>
        public abstract void Remove(T item);
    }
}