using System.Collections.Generic;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
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

        /// <summary>
        ///     Adds a slide into the collection at the specified position.
        /// </summary>
        void Add(ISlide addingSlide);

        /// <summary>
        ///     Inserts slide.
        /// </summary>
        /// <param name="position">Position (number) at which slide should be inserted.</param>
        /// <param name="outerSlide">The slide to insert.</param>
        void Insert(int position, ISlide outerSlide);
    }
}