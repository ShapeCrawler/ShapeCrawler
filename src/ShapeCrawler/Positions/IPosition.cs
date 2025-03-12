#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents location of a shape.
/// </summary>
public interface IPosition
{
    /// <summary>
    ///     Gets or sets x-coordinate of the upper-left corner of the shape in points.
    /// </summary>
    float X { get; set; }

    /// <summary>
    ///     Gets or sets y-coordinate of the upper-left corner of the shape in points.
    /// </summary>
    float Y { get; set; }
}