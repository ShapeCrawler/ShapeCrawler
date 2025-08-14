using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using Position = ShapeCrawler.Positions.Position;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal sealed class LineShape : Shape
{
    internal LineShape(Position position,
        ShapeSize shapeSize,
        ShapeId shapeId,
        P.ConnectionShape pConnectionShape) :
        base(position, shapeSize, shapeId, pConnectionShape)
    {
        this.Line = new Line(pConnectionShape, this);
    }

    public override ILine? Line { get; }
}