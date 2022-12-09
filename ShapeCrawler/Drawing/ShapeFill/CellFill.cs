using DocumentFormat.OpenXml;

namespace ShapeCrawler.Drawing.ShapeFill;

internal class CellFill : ShapeFill
{
    internal CellFill(SlideObject slideObject, TypedOpenXmlCompositeElement cellProperties)
        : base(slideObject, cellProperties)
    {
    }
}