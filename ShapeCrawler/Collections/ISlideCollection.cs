using System.Collections.Generic;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represent a collection of slides.
    /// </summary>
    public interface ISlideCollection : IReadOnlyList<ISlide>
    {
        /// <summary>
        ///     Removes slide from presentation.
        /// </summary>
        void Remove(ISlide removingSlide);

#if DEBUG
        /// <summary>
        ///     Adds existing slide from other presentation.
        /// </summary>
        /// <param name="copiedSlide">Slide of other presentation.</param>
        /// <param name="keepSourceFormat">Value indicating whether keep the formatting of the copying slide.</param>
        /// <returns>Added slide.</returns>
        ISlide AddExternal(ISlide copiedSlide, bool keepSourceFormat);
#endif
    }
}