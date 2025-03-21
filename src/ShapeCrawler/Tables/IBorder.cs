#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a top border of a table cell.
/// </summary>
public interface IBorder
{
    /// <summary>
    ///     Gets or sets border width in points.
    /// </summary>
    decimal Width { get; set; }

    /// <summary>
    ///     Gets or sets the border color.
    /// </summary>
    public string? Color { get; set; }
}