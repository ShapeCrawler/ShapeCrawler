#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///    Represents a font.
/// </summary>
public interface IFont
{
    /// <summary>
    ///     Gets or sets font size in points.
    /// </summary>
    float Size { get; set; }
}