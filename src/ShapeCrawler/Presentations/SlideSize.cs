using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

internal sealed class SlideSize(P.SlideSize pSlideSize)
{
    internal decimal Width() => new Emus(pSlideSize.Cx!.Value).AsHorizontalPixels();

    internal decimal Height() => new Emus(pSlideSize.Cy!.Value).AsVerticalPixels();

    internal void UpdateWidth(decimal pixels)
    {
        var emus = new Pixels(pixels).AsHorizontalEmus();
        pSlideSize.Cx = new Int32Value((int)emus);
    }

    internal void UpdateHeight(decimal pixels)
    {
        var emus = new Pixels(pixels).AsVerticalEmus();
        pSlideSize.Cy = new Int32Value((int)emus);
    }
}