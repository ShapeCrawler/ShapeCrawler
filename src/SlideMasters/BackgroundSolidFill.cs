using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

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
            var pBackground = pCommonSlideData?.GetFirstChild<DocumentFormat.OpenXml.Presentation.Background>();
            var pBackgroundProperties = pBackground?.GetFirstChild<DocumentFormat.OpenXml.Presentation.BackgroundProperties>();

            if (pBackgroundProperties != null)
            {
                var aSolidFill = pBackgroundProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
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