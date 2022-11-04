using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a picture shape on a slide.
/// </summary>
public interface IPicture : IShape
{
    /// <summary>
    ///     Gets image, if picture is binary image, otherwise <see langword="null"/>.
    /// </summary>
    IImage? Image { get; }

    /// <summary>
    ///     Gets SVG content if picture is SVG graphic, otherwise <see langword="null"/>.
    /// </summary>
    string? SvgContent { get; }
}