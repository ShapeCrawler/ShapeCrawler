using DocumentFormat.OpenXml;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal sealed class MediaShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement) : Shape(position, shapeSize, shapeId, pShapeTreeElement)
{
    
}