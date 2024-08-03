

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a top border of a table cell.
/// </summary>
public interface IBorder
{
    /// <summary>
    ///     Gets or sets border width in points.
    /// </summary>
    float Width { get; set; }
}