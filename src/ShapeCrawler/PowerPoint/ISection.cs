namespace ShapeCrawler;

/// <summary>
///     Represents a presentation section.
/// </summary>
public interface ISection
{
    /// <summary>
    ///     Gets section slides.
    /// </summary>
    ISectionSlideCollection Slides { get; }

    /// <summary>
    ///     Gets section name.
    /// </summary>
    string Name { get; }
}