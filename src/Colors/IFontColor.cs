#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a color format.
/// </summary>
public interface IFontColor
{
    /// <summary>
    ///     Gets color type.
    /// </summary>
    ColorType Type { get; }

    /// <summary>
    ///     Gets color hexadecimal representation.
    /// </summary>
    string Hex { get; }

    /// <summary>
    ///     Updates color with specified hexadecimal representation.
    /// </summary>
    void Update(string hex);
}