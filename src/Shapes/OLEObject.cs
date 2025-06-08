using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
// ReSharper disable InconsistentNaming

namespace ShapeCrawler.Shapes;

internal sealed class OLEObject : Shape
{
    internal OLEObject(
        Position position,
        ShapeSize shapeSize,
        ShapeId shapeId,
        OpenXmlElement pShapeTreeElement) : base(position, shapeSize, shapeId, pShapeTreeElement)
    {
    }
}