

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape outline.
/// </summary>
public interface IShapeOutline
{
    /// <summary>
    ///     Gets or sets outline weight in points.
    /// </summary>
    decimal Weight { get; set; }

    /// <summary>
    ///     Gets or sets color in hexadecimal format. Returns <see langword="null"/> if outline is not filled.
    /// </summary>
    string? HexColor { get; set; }
}