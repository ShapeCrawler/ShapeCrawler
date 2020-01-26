using System.Collections.Generic;

namespace SlideXML.Models
{
    /// <summary>
    /// Provides APIs for slide collection.
    /// </summary>
    public interface ISlideCollection : IEnumerable<SlideSL>
    {
        void Add(SlideSL item);

        void Remove(SlideSL item);

        SlideSL this[int index]
        {
            get;
        }
    }
}