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
    ///     Gets color in hexadecimal format. Returns <see langword="null"/> if outline is not filled.
    /// </summary>
    /// <returns>
    /// Hexadecimal color, for example, "FF0000".
    /// </returns>
    string? HexColor { get; }

    /// <summary>
    ///     Sets color in hexadecimal format.
    /// </summary>
    /// <example>
    ///    <code>
    ///     shape.Outline.SetHexColor("FF0000");
    ///    </code>
    /// </example>
    void SetHexColor(string value);

    /// <summary>
    ///     Sets shape outline to "No outline".
    /// </summary>
    void SetNoOutline();
}