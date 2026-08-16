using FluentAssertions;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class AllLineTypesTests : SCTest
{
    [Test]
    public void AddLine_creates_all_connector_variants_and_they_can_be_read_and_updated()
    {
        var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        var connectorTypes = new[]
        {
            LineType.Line,
            LineType.Arrow,
            LineType.DoubleArrow,
            LineType.ElbowConnector,
            LineType.ElbowArrowConnector,
            LineType.ElbowDoubleArrowConnector,
            LineType.CurvedConnector,
            LineType.CurvedArrowConnector,
            LineType.CurvedDoubleArrowConnector,
        };

        foreach (var type in connectorTypes)
        {
            shapes.AddLine(10, 20, 100, 80, type);
        }

        shapes.Count.Should().Be(connectorTypes.Length);
        for (var i = 0; i < connectorTypes.Length; i++)
        {
            var line = shapes[i].Line;
            line.Should().NotBeNull();
            line!.Type.Should().Be(connectorTypes[i]);
            line.StartPoint = new Point(30, 40);
            line.EndPoint = new Point(130, 140);
            line.StartPoint.Should().Be(new Point(30, 40));
            line.EndPoint.Should().Be(new Point(130, 140));
        }

        ValidatePresentation(presentation);
    }

    [Test]
    public void AddCurve_AddFreeformShape_and_AddScribble_create_readable_updatable_lines()
    {
        var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        var points = new[] { new Point(10, 20), new Point(40, 60), new Point(80, 30) };

        shapes.AddCurve(points);
        shapes.AddFreeformShape(points);
        shapes.AddScribble(points);

        shapes[0].Line!.Type.Should().Be(LineType.Curve);
        shapes[1].Line!.Type.Should().Be(LineType.FreeformShape);
        shapes[2].Line!.Type.Should().Be(LineType.Scribble);
        foreach (var shape in shapes)
        {
            shape.Line!.StartPoint = new Point(20, 30);
            shape.Line.EndPoint = new Point(90, 100);
            shape.Line.StartPoint.Should().Be(new Point(20, 30));
            shape.Line.EndPoint.Should().Be(new Point(90, 100));
        }

        ValidatePresentation(presentation);
    }
}
