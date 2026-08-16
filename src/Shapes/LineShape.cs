using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal sealed class LineShape : DrawingShape
{
    internal LineShape(
        Position position,
        ShapeSize shapeSize,
        ShapeId shapeId,
        DocumentFormat.OpenXml.OpenXmlElement shapeElement)
        : base(position, shapeSize, shapeId, shapeElement)
    {
        this.Line = new Line(shapeElement, this);
    }

    public override ILine? Line { get; }
}