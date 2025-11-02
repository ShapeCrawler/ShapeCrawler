using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
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

/// <summary>
///     Represents solid fill properties.
/// </summary>
public interface ISolidFill
{
    /// <summary>
    ///     Gets the color in hexadecimal format.
    /// </summary>
    string Color { get; }
}

internal sealed class SlideLayoutBackground : ISlideLayoutBackground
{
    private readonly SlideLayoutPart slideLayoutPart;
    private BackgroundSolidFill? solidFill;

    internal SlideLayoutBackground(SlideLayoutPart slideLayoutPart)
    {
        this.slideLayoutPart = slideLayoutPart;
    }

    public ISolidFill SolidFill => this.solidFill ??= new BackgroundSolidFill(this.slideLayoutPart);

    public void SolidFillColor(string hex)
    {
        var pCommonSlideData = this.slideLayoutPart.SlideLayout.CommonSlideData
                               ?? this.slideLayoutPart.SlideLayout.AppendChild<P.CommonSlideData>(new());

        var pBackground = pCommonSlideData.GetFirstChild<P.Background>()
                          ?? pCommonSlideData.InsertAt<P.Background>(new(), 0);

        var pBackgroundProperties = pBackground.GetFirstChild<P.BackgroundProperties>()
                                    ?? pBackground.AppendChild<P.BackgroundProperties>(new());

        pBackgroundProperties.AddSolidFill(hex);
    }
}

internal sealed class BackgroundSolidFill : ISolidFill
{
    private readonly SlideLayoutPart slideLayoutPart;

    internal BackgroundSolidFill(SlideLayoutPart slideLayoutPart)
    {
        this.slideLayoutPart = slideLayoutPart;
    }

    public string Color
    {
        get
        {
            var pCommonSlideData = this.slideLayoutPart.SlideLayout.CommonSlideData;
            var pBackground = pCommonSlideData?.GetFirstChild<P.Background>();
            var pBackgroundProperties = pBackground?.GetFirstChild<P.BackgroundProperties>();

            if (pBackgroundProperties != null)
            {
                var aSolidFill = pBackgroundProperties.GetFirstChild<A.SolidFill>();
                if (aSolidFill != null)
                {
                    var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                    if (aRgbColorModelHex != null)
                    {
                        return aRgbColorModelHex.Val!.ToString()!;
                    }
                }
            }

            return string.Empty;
        }
    }
}