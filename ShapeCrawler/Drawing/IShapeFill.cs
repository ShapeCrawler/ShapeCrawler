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
    ///     Gets picture image if it is picture fill, otherwise <see langword="null"/>
    /// </summary>
    public IImage? Picture { get; }

    /// <summary>
    ///     Gets color hexadecimal representation if it is solid color, otherwise <see langword="null"/>.
    /// </summary>
    public string? HexSolidColor { get; }

    /// <summary>
    ///     Sets picture fill.
    /// </summary>
    void SetPicture(Stream image);

    /// <summary>
    ///     Sets solid color.
    /// </summary>
    void SetHexSolidColor(string hex);
}