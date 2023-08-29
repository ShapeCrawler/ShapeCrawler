using DocumentFormat.OpenXml;

namespace ShapeCrawler.AutoShapes;

internal readonly struct NewAutoShape
{
    internal NewAutoShape(SlideShape.SlideShape newShape, TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.Shape = newShape;
        this.PShapeTreeChild = pShapeTreeChild;
    }
    
    internal TypedOpenXmlCompositeElement PShapeTreeChild { get; }

    internal SlideShape.SlideShape Shape { get; }
}