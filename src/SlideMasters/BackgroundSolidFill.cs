using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.SlideMasters;

internal sealed class BackgroundSolidFill(SlideLayoutPart slideLayoutPart) : ISolidFill
{
    public string Color
    {
        get
        {
            var pCommonSlideData = slideLayoutPart.SlideLayout.CommonSlideData;
            var pBackground = pCommonSlideData?.GetFirstChild<DocumentFormat.OpenXml.Presentation.Background>();
            var pBackgroundProperties = pBackground?.GetFirstChild<DocumentFormat.OpenXml.Presentation.BackgroundProperties>();

            var aSolidFill = pBackgroundProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();

            var aRgbColorModelHex = aSolidFill?.RgbColorModelHex;

            return aRgbColorModelHex != null ? aRgbColorModelHex.Val!.ToString()! : string.Empty;
        }
    }
}