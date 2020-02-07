using System.Collections.Generic;

namespace SlideXML.Models
{
    /// <summary>
    /// Provides APIs for slide collection.
    /// </summary>
    public interface ISlideCollection : IEnumerable<Slide>
    {
        void Add(Slide item);

        void Remove(Slide item);

        Slide this[int index]
        {
            get;
        }
    }
}