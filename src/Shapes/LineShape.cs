using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal sealed class LineShape : DrawingShape
{
    internal LineShape(
        Position position,
        ShapeSize shapeSize,
        ShapeId shapeId,
        OpenXmlElement shapeElement)
        : base(position, shapeSize, shapeId, shapeElement)
    {
        this.Line = new Line(shapeElement, this);
    }

    public override ILine? Line { get; }
}
