using DocumentFormat.OpenXml;

namespace ShapeCrawler.Drawing.ShapeFill;

internal sealed class CellFill : ShapeFill
{
    internal CellFill(SlideObject slideObject, TypedOpenXmlCompositeElement cellProperties)
        : base(slideObject, cellProperties)
    {
    }
}