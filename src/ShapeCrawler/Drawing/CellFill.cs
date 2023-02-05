using DocumentFormat.OpenXml;

namespace ShapeCrawler.Drawing.ShapeFill;

internal sealed class CellFill : SCShapeFill
{
    internal CellFill(SlideObject slideObject, TypedOpenXmlCompositeElement cellProperties)
        : base(slideObject, cellProperties)
    {
    }
}