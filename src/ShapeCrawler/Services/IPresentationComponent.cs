namespace ShapeCrawler.Services;

/// <summary>
///     Represents a presentation component.
/// </summary>
public interface IPresentationComponent
{
    /// <summary>
    ///     Gets presentation.
    /// </summary>
    IPresentation Presentation { get; }
}