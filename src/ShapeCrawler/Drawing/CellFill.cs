using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing.ShapeFill;

namespace ShapeCrawler.Drawing;

internal sealed class CellFill : SCShapeFill
{
    internal CellFill(SlideStructure slideObject, TypedOpenXmlCompositeElement cellProperties)
        : base(slideObject, cellProperties)
    {
    }
}