using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    public interface ISlideCollection : IReadOnlyList<SCSlide>
    {
        void Remove(SCSlide removingSlide);
    }
}