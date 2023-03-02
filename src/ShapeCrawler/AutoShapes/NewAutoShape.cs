using DocumentFormat.OpenXml;

namespace ShapeCrawler.AutoShapes;

internal readonly struct NewAutoShape
{
    internal TypedOpenXmlCompositeElement pShapeTreeChild { get; }
    internal SCAutoShape autoShape { get; }

    internal NewAutoShape(SCAutoShape newAutoShape, TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.autoShape = newAutoShape;
        this.pShapeTreeChild = pShapeTreeChild;
    }
}