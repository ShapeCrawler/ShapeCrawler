using System.Collections.Generic;

// ReSharper disable InconsistentNaming

// ReSharper disable once CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of slides.
/// </summary>
public interface ISlideCollection : IReadOnlyList<ISlide>
{
    /// <summary>
    ///     Adds a new slide using specified layout.
    /// </summary>
    void Add(int layoutNumber);
    
    /// <summary>
    ///     Adds a new slide using specified layout at the specified position.
    /// </summary>
    /// <param name="layoutNumber">Slide Layout number.</param>
    /// <param name="slideNumber">Slide number.</param>
    void Add(int layoutNumber, int slideNumber);

    /// <summary>
    ///     Adds slide.
    /// </summary>
    void Add(ISlide slide);

    /// <summary>
    ///     Adds slide at the specified position.
    /// </summary>
    void Add(ISlide slide, int slideNumber);
}