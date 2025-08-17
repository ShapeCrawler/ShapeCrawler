using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;

// ReSharper disable InconsistentNaming
namespace ShapeCrawler.Shapes;

internal sealed class OleObject : Shape
{
    internal OleObject(
        Position position,
        ShapeSize shapeSize,
        ShapeId shapeId,
        OpenXmlElement pShapeTreeElement)
        : base(position, shapeSize, shapeId, pShapeTreeElement)
    {
    }
}