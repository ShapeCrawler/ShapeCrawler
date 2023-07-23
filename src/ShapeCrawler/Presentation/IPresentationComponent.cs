namespace ShapeCrawler;

/// <summary>
///     Represents a presentation component.
/// </summary>
public interface IPresentationComponent // TODO: make internal
{
    /// <summary>
    ///     Gets presentation.
    /// </summary>
    IPresentation Presentation { get; }
}