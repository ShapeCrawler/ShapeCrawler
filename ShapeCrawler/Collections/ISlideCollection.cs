using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of slides.
    /// </summary>
    public interface ISlideCollection : IReadOnlyList<ISlide>
    {
        /// <summary>
        ///     Removes slide from presentation.
        /// </summary>
        void Remove(ISlide removingSlide);

#if DEBUG
        /// <summary>
        ///     Adds slide.
        /// </summary>
        void Add(ISlide addingSlide);
#endif
    }
}