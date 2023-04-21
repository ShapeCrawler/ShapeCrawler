using DocumentFormat.OpenXml;

namespace ShapeCrawler.AutoShapes;

internal readonly struct NewAutoShape
{
    internal NewAutoShape(SCAutoShape newAutoShape, TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.AutoShape = newAutoShape;
        this.PShapeTreeChild = pShapeTreeChild;
    }
    
    internal TypedOpenXmlCompositeElement PShapeTreeChild { get; }

    internal SCAutoShape AutoShape { get; }
}