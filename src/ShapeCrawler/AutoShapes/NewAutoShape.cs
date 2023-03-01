using DocumentFormat.OpenXml;

namespace ShapeCrawler.AutoShapes;

internal class NewAutoShape
{
    internal readonly TypedOpenXmlCompositeElement pShapeTreeChild;
    internal readonly SCAutoShape newAutoShape;

    internal NewAutoShape(SCAutoShape newAutoShape, TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        this.newAutoShape = newAutoShape;
        this.pShapeTreeChild = pShapeTreeChild;
    }
}