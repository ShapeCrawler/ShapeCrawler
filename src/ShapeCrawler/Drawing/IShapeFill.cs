using System.IO;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape fill.
/// </summary>
public interface IShapeFill
{
    /// <summary>
    ///     Gets fill type.
    /// </summary>
    public SCFillType Type { get; }

    /// <summary>
    ///     Gets picture image if it is picture fill, otherwise <see langword="null"/>.
    /// </summary>
    public IImage? Picture { get; }

    /// <summary>
    ///     Gets color in hexadecimal representation if it is filled with solid color, otherwise <see langword="null"/>.
    /// </summary>
    public string? Color { get; }

    /// <summary>
    /// Gets the Alpha percentage of the fill color.
    /// </summary>
    public double AlphaPercentage { get; }

    /// <summary>
    /// Gets the Luminance Modulation Percentage of the fill color.
    /// </summary>
    public double LuminanceModulation { get; }
    
    /// <summary>
    /// Gets the Luminance Offset Percentage of the fill color.
    /// </summary>
    public double LuminanceOffset { get; }

    /// <summary>
    ///     Fills the shape with picture.
    /// </summary>
    void SetPicture(Stream image);

    /// <summary>
    ///     Fills the shape with solid color in hexadecimal representation.
    /// </summary>
    void SetColor(string hex);
}