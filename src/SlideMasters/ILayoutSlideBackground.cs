using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a slide layout background.
/// </summary>
public interface ILayoutSlideBackground
{
    /// <summary>
    ///     Gets the solid fill properties of the background.
    /// </summary>
    ISolidFill SolidFill { get; }

    /// <summary>
    ///     Sets the background to a solid color.
    /// </summary>
    /// <param name="hex">The color in hexadecimal format.</param>
    void SolidFillColor(string hex);

    /// <summary>
    ///     Sets the background to a picture.
    /// </summary>
    /// <param name="image">The image stream for the background picture.</param>
    void Picture(Stream image);

    /// <summary>
    ///     Gets the background picture image stream.
    /// </summary>
    /// <returns>The image stream of the background picture.</returns>
    MemoryStream Picture();
}

internal sealed class LayoutSlideBackground(SlideLayoutPart slideLayoutPart) : ILayoutSlideBackground
{
    private BackgroundSolidFill? solidFill;

    public ISolidFill SolidFill => this.solidFill ??= new BackgroundSolidFill(slideLayoutPart);

    public void SolidFillColor(string hex)
    {
        var pCommonSlideData = slideLayoutPart.SlideLayout.CommonSlideData
                               ?? slideLayoutPart.SlideLayout.AppendChild(new P.CommonSlideData());

        var pBackground = pCommonSlideData.GetFirstChild<P.Background>()
                          ?? pCommonSlideData.InsertAt(new P.Background(), 0);

        var pBackgroundProperties = pBackground.GetFirstChild<P.BackgroundProperties>()
                                    ?? pBackground.AppendChild(new P.BackgroundProperties());

        pBackgroundProperties.AddSolidFill(hex);
    }

    public void Picture(Stream image)
    {
        var pCommonSlideData = slideLayoutPart.SlideLayout.CommonSlideData
                               ?? slideLayoutPart.SlideLayout.AppendChild(new P.CommonSlideData());

        var pBackground = pCommonSlideData.GetFirstChild<P.Background>()
                          ?? pCommonSlideData.InsertAt(new P.Background(), 0);

        var pBackgroundProperties = pBackground.GetFirstChild<P.BackgroundProperties>()
                                    ?? pBackground.AppendChild(new P.BackgroundProperties());

        var rId = slideLayoutPart.AddImagePart(image, "image/png");
        pBackgroundProperties.GetFirstChild<A.GradientFill>()?.Remove();
        pBackgroundProperties.GetFirstChild<A.SolidFill>()?.Remove();
        pBackgroundProperties.GetFirstChild<A.PatternFill>()?.Remove();
        pBackgroundProperties.GetFirstChild<A.NoFill>()?.Remove();
        pBackgroundProperties.GetFirstChild<A.BlipFill>()?.Remove();

        var aBlipFill = new A.BlipFill(
            new A.Blip { Embed = rId },
            new A.Stretch(new A.FillRectangle()));

        var aOutline = pBackgroundProperties.GetFirstChild<A.Outline>();
        if (aOutline != null)
        {
            pBackgroundProperties.InsertBefore(aBlipFill, aOutline);
        }
        else
        {
            pBackgroundProperties.Append(aBlipFill);
        }
    }

    public MemoryStream Picture()
    {
        var pBackground = slideLayoutPart.SlideLayout.CommonSlideData?.GetFirstChild<P.Background>();
        var aBlipFill = pBackground?.GetFirstChild<P.BackgroundProperties>()?.GetFirstChild<A.BlipFill>();
        var aBlip = aBlipFill?.Blip;
        if (aBlip?.Embed?.Value is null)
        {
            throw new InvalidOperationException("Background picture not found.");
        }

        var imagePart = (ImagePart)slideLayoutPart.GetPartById(aBlip.Embed.Value);
        using var stream = imagePart.GetStream();
        var mStream = new MemoryStream();
        stream.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
}