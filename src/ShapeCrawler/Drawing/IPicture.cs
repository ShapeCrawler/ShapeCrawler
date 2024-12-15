#pragma warning disable IDE0130
using ShapeCrawler.Drawing;

namespace ShapeCrawler;
#pragma warning restore IDE0130

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
    ///     Gets or sets the cropping frame for this image.
    /// </summary>
    CroppingFrame Crop { get; set; }

    /// <summary>
    ///     Gets or sets the transparency for this image. Range is 0 (fully opaque, default) to 100 (fully transparent).
    /// </summary>
    decimal Transparency { get; set; }

    /// <summary>
    ///     Sends the shape backward in the z-order.
    /// </summary>
    void SendToBack();
}