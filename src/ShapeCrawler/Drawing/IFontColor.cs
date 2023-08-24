

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
    SCColorType ColorType { get; }

    /// <summary>
    ///     Gets color hexadecimal representation.
    /// </summary>
    string ColorHex { get; }

    /// <summary>
    ///     Sets solid color by its hexadecimal representation.
    /// </summary>
    void SetColorByHex(string hex);
}