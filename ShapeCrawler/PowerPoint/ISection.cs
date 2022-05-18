using System.Collections.Generic;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represent a presentation section.
    /// </summary>
    public interface ISection
    {
        /// <summary>
        ///     Gets section slides.
        /// </summary>
        List<ISlide> Slides { get; }

        string Name { get; }
    }
}