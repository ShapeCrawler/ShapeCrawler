using System.Collections.Generic;
using ShapeCrawler.SlideMasters;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Master.
/// </summary>
public interface ISlideMaster
{
    /// <summary>
    ///     Gets background image if slide master has background, otherwise <see langword="null"/>.
    /// </summary>
    IImage? Background { get; }

    /// <summary>
    ///     Gets collection of Slide Layouts.
    /// </summary>
    IReadOnlyList<ISlideLayout> SlideLayouts { get; }

    /// <summary>
    ///     Gets collection of shape.
    /// </summary>
    IShapeCollection Shapes { get; }
}