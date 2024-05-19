// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents location of a shape.
/// </summary>
public interface IPosition
{
    /// <summary>
    ///     Gets or sets x-coordinate of the upper-left corner of the shape.
    /// </summary>
    decimal X { get; set; }

    /// <summary>
    ///     Gets or sets y-coordinate of the upper-left corner of the shape in pixels.
    /// </summary>
    decimal Y { get; set; }
}