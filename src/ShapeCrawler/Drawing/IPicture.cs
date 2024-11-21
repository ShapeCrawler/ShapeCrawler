// ReSharper disable CheckNamespace

namespace ShapeCrawler;

/// <summary>
///     Represents a picture shape.
/// </summary>
public interface IPicture : IShape
{
    /// <summary>
    ///     Gets image. Returns <see langword="null"/> if the picture is not binary picture. 
    /// </summary>
    IImage? Image { get; }

    /// <summary>
    ///     Gets SVG content. Returns <see langword="null"/> if the picture is not SVG graphic.
    /// </summary>
    string? SvgContent { get; }

    /// <summary>
    ///     Sends the shape backward in the z-order.
    /// </summary>
    void SendToBack();
}