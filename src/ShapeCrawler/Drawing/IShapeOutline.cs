

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
    ///     Gets color in 6-digit hexadecimal format. Returns <see langword="null"/> if outline is not filled.
    /// </summary>
    string? HexColor { get; }

    /// <summary>
    ///     Sets color in 6-digit hexadecimal format.
    /// </summary>
    void SetHexColor(string value);

    /// <summary>
    ///     Sets shape outline to "No outline".
    /// </summary>
    void SetNoOutline();
}