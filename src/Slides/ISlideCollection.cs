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
    void Add(ISlideLayout layout);

    /// <summary>
    ///     Adds slide.
    /// </summary>
    void Add(ISlide slide);

    /// <summary>
    ///     Adds slide at the specified number position.
    /// </summary>
    void Add(ISlide slide, int number);
#if DEBUG
    /// <summary>
    ///     Adds a new slide from JSON.
    /// </summary>
    /// <param name="jsonSlide">Slide in JSON format.</param>
    [System.Diagnostics.CodeAnalysis.Experimental("SCE0001")]
    void AddJSON(string jsonSlide);
#endif
}