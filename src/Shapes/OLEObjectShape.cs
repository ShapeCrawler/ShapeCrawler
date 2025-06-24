using DocumentFormat.OpenXml;
using Position = ShapeCrawler.Positions.Position;
// ReSharper disable InconsistentNaming

namespace ShapeCrawler.Shapes;

internal sealed class OLEObjectShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement) : Shape(position, shapeSize, shapeId, pShapeTreeElement)
{
    
}