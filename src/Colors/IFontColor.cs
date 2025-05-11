#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a font color format.
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
    ///     Updates the color with the specified color in the hexadecimal representation.
    /// </summary>
    void Update(string hex);
}