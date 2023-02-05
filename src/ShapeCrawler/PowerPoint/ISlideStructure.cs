using ShapeCrawler.Services;

namespace ShapeCrawler;

/// <summary>
///     Represents slide object.
/// </summary>
public interface ISlideStructure : IPresentationComponent
{
    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }
}