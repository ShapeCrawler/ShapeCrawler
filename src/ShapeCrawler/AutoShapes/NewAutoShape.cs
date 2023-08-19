using DocumentFormat.OpenXml;

namespace ShapeCrawler.AutoShapes;

internal readonly struct NewAutoShape
{
    internal NewAutoShape(SlideAutoShape newAutoShape, TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.AutoShape = newAutoShape;
        this.PShapeTreeChild = pShapeTreeChild;
    }
    
    internal TypedOpenXmlCompositeElement PShapeTreeChild { get; }

    internal SlideAutoShape AutoShape { get; }
}