using System.Collections.Generic;
using ShapeCrawler.SlideMasters;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collection of slides.
/// </summary>
public interface ISlideCollection : IReadOnlyList<ISlide>
{
    /// <summary>
    ///     Removes specified slide.
    /// </summary>
    void Remove(ISlide slide);

#if DEBUG
    /// <summary>
    ///     Creates a new slide.
    /// </summary>
    ISlide AddEmptySlide(ISlideLayout layout);
#endif
    
    /// <summary>
    ///     Adds specified slide.
    /// </summary>
    void Add(ISlide slide);

    /// <summary>
    ///     Inserts slide at specified position.
    /// </summary>
    /// <param name="position">Position at which specified slide will be inserted.</param>
    /// <param name="slide">The slide to insert.</param>
    void Insert(int position, ISlide slide);
}