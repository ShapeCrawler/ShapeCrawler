using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Slides;

internal sealed class TableShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement) : Shape(position, shapeSize, shapeId, pShapeTreeElement)
{
}