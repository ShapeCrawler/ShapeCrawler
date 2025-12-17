using Position = ShapeCrawler.Positions.Position;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class LineShape : DrawingShape
{
    internal LineShape(
        Position position,
        ShapeSize shapeSize,
        ShapeId shapeId,
        P.ConnectionShape pConnectionShape)
        : base(position, shapeSize, shapeId, pConnectionShape)
    {
        this.Line = new Line(pConnectionShape, this);
    }

    public override ILine? Line { get; }
}