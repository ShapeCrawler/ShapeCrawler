using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a slide layout background.
/// </summary>
public interface ISlideLayoutBackground
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
}

internal sealed class SlideLayoutBackground(SlideLayoutPart slideLayoutPart) : ISlideLayoutBackground
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
}