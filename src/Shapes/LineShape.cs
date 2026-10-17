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

    /// <summary>
    ///     Gets or sets the geometry, identifying both straight connector presets as a line.
    /// </summary>
    public override Geometry GeometryType
    {
        get
        {
            var geometry = base.GeometryType;
            return geometry == Geometry.StraightConnector1 ? Geometry.Line : geometry;
        }

        set => base.GeometryType = value;
    }

    public override ILine? Line { get; }
}