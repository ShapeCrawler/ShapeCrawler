using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideSize(P.SlideSize pSlideSize)
{
    internal decimal Width
    {
        get => new Emus(pSlideSize.Cx!.Value).AsPoints();
        set
        {
            var emus = new Points(value).AsEmus();
            pSlideSize.Cx = new Int32Value((int)emus);
        }
    }

    internal decimal Height
    {
        get => new Emus(pSlideSize.Cy!.Value).AsPoints();
        set
        {
            var emus = new Points(value).AsEmus();
            pSlideSize.Cy = new Int32Value((int)emus);
        }
    }
}