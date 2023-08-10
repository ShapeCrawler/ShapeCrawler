using DocumentFormat.OpenXml;

namespace ShapeCrawler.AutoShapes;

internal readonly struct NewAutoShape
{
    internal NewAutoShape(SCSlideAutoShape newAutoShape, TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.AutoShape = newAutoShape;
        this.PShapeTreeChild = pShapeTreeChild;
    }
    
    internal TypedOpenXmlCompositeElement PShapeTreeChild { get; }

    internal SCSlideAutoShape AutoShape { get; }
}