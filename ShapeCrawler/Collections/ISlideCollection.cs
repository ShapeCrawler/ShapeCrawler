using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    public interface ISlideCollection : IReadOnlyList<SlideSc>
    {
        void Remove(SlideSc removingSlide);
    }
}