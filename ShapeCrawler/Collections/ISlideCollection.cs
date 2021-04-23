using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    public interface ISlideCollection : IReadOnlyList<ISlide>
    {
        void Remove(ISlide removingSlide);
    }
}