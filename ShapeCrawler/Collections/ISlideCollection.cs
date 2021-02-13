using System.Collections.Generic;

namespace ShapeCrawler
{
    public interface ISlideCollection : IReadOnlyList<SlideSc>
    {
        void Remove(SlideSc removingSlide);
    }
}