using System.Collections.Generic;

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a collection of a slides.
    /// </summary>
    public interface ISlideCollection : IEnumerable<Slide>
    {
        /// <summary>
        /// Removes slide from collection.
        /// </summary>
        /// <param name="item"></param>
        void Remove(Slide item);

        /// <summary>
        /// Returns the element at the specified index.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        Slide this[int index] { get; }
    }
}