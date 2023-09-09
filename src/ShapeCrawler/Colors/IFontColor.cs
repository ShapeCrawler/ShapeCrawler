

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a color format.
/// </summary>
public interface IFontColor
{
    /// <summary>
    ///     Gets color type.
    /// </summary>
    SCColorType Type { get; }

    /// <summary>
    ///     Gets color hexadecimal representation.
    /// </summary>
    string Hex { get; }

    /// <summary>
    ///     Updates color with specified hexadecimal representation.
    /// </summary>
    void Update(string hex);
}