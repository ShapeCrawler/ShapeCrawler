#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a font color.
/// </summary>
public interface IFontColor
{
    /// <summary>
    ///     Gets the color type.
    /// </summary>
    ColorType Type { get; }

    /// <summary>
    ///     Gets the color hexadecimal representation.
    /// </summary>
    string Hex { get; }

    /// <summary>
    ///     Set the color by its hexadecimal representation.
    /// </summary>
    void Set(string hex);
}