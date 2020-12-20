namespace SlideDotNet.Collections
{
    /// <summary>
    /// Represents a base class for collections which allow item removing.
    /// </summary>
    public abstract class EditAbleCollection<T> : LibraryCollection<T>
    {
        /// <summary>
        /// Removes the specific object from the collection.
        /// </summary>
        /// <param name="item"></param>
        public abstract void Remove(T item);
    }
}