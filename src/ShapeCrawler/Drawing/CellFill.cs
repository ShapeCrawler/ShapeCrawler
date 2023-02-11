using DocumentFormat.OpenXml;

namespace ShapeCrawler.Drawing.ShapeFill;

internal sealed class CellFill : SCShapeFill
{
    internal CellFill(SlideStructure slideObject, TypedOpenXmlCompositeElement cellProperties)
        : base(slideObject, cellProperties)
    {
    }
}