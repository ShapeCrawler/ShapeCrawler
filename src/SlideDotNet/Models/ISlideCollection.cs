using System.Collections.Generic;

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a collection of a slides.
    /// </summary>
    public interface ISlideCollection : IEnumerable<Slide>
    {
        void Remove(Slide item);

        Slide this[int index] { get; }
    }
}