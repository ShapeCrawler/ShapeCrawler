using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler;

internal sealed class LineShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement) : Shape(position, shapeSize, shapeId, pShapeTreeElement)
{
}