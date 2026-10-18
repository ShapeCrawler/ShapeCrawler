using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using SkiaSharp;

namespace ShapeCrawler.Slides;

internal sealed class SlideTemplate(SlidePart slidePart)
{
    internal void Render(SKCanvas canvas)
    {
        if (slidePart.Slide?.ShowMasterShapes?.Value == false)
        {
            return;
        }

        var layout = slidePart.SlideLayoutPart;
        if (layout?.SlideLayout?.ShowMasterShapes?.Value != false && layout?.SlideMasterPart is { } master)
        {
            RenderLayer(canvas, master);
        }

        if (layout is not null)
        {
            RenderLayer(canvas, layout);
        }
    }

    private static void RenderLayer(SKCanvas canvas, OpenXmlPart part)
    {
        foreach (var shape in new ShapeCollection(part).GetInternalShapes())
        {
            if (!shape.Hidden)
            {
                shape.RenderTemplate(canvas);
            }
        }
    }
}
