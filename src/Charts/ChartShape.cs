using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal sealed class ChartShape : Shape
{
    internal ChartShape(Chart chart, P.GraphicFrame pGraphicFrame)
        : base(new Position(pGraphicFrame), new ShapeSize(pGraphicFrame), new ShapeId(pGraphicFrame), pGraphicFrame)
    {
        this.Chart = chart;
    }

    public override IChart? Chart { get; }

    public override Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("Geometry type cannot be set for Chart shape.");
    }
}