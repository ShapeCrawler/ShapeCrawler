using System.Collections.Generic;

// ReSharper disable once CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of slides.
/// </summary>
public interface ISlides : IReadOnlyList<ISlide>
{
    /// <summary>
    ///     Removes specified slide.
    /// </summary>
    void Remove(ISlide slide);

    /// <summary>
    ///     Adds a new slide based on the existing layout.
    /// </summary>
    void AddEmptySlide(ISlideLayout layout);

    /// <summary>
    ///     Adds a new slide based on the predefined layout type.
    /// </summary>
    void AddEmptySlide(SlideLayoutType layoutType);

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