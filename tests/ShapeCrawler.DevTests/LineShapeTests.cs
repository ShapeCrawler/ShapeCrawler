using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class LineShapeTests : SCTest
{
    [TestCase(1)]
    [TestCase(2)]
    public void Straight_connectors_from_pptx_have_line_geometry_and_endpoints_in_points(int slideNumber)
    {
        // Arrange
        var presentation = new Presentation(TestAsset("straight-lines.pptx"));
        var expectedLines = new[]
        {
            (Name: "Horizontal arrow", Start: new Point(50.25m, 60.5m), End: new Point(150.75m, 60.5m)),
            (Name: "Vertical arrow", Start: new Point(200.25m, 80.5m), End: new Point(200.25m, 180.75m)),
            (Name: "Diagonal arrow", Start: new Point(300.25m, 50.5m), End: new Point(400.75m, 150.25m)),
            (Name: "Upward diagonal arrow", Start: new Point(300.25m, 250.75m), End: new Point(400.75m, 150.25m)),
            (Name: "Reversed horizontal arrow", Start: new Point(150.75m, 300.5m), End: new Point(50.25m, 300.5m)),
            (Name: "Reversed vertical arrow", Start: new Point(200.25m, 300.75m), End: new Point(200.25m, 200.5m)),
            (Name: "Reversed diagonal arrow", Start: new Point(500.75m, 300.25m), End: new Point(400.25m, 200.5m)),
            (Name: "Horizontal plain line", Start: new Point(50.25m, 60.5m), End: new Point(150.75m, 60.5m)),
        };

        // Act
        var shapes = presentation.Slide(slideNumber).Shapes.ToArray();

        // Assert
        shapes.Should().HaveCount(expectedLines.Length + 1);
        shapes.Single(shape => shape.Name == "Rectangle").GeometryType.Should().Be(Geometry.Rectangle);
        foreach (var expected in expectedLines)
        {
            var shape = shapes.Single(shape => shape.Name == expected.Name);
            shape.GeometryType.Should().Be(Geometry.Line);
            shape.Line.Should().NotBeNull();
            shape.Line!.StartPoint.Should().Be(expected.Start);
            shape.Line.EndPoint.Should().Be(expected.End);
        }

        ValidatePresentation(presentation);
    }

    [Test]
    public void StartPoint_returns_start_point_coordinates()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);
        var line = shapes.Last().Line;

        // Act
        var startPoint = line.StartPoint;

        // Assert
        startPoint.X.Should().Be(50);
        startPoint.Y.Should().Be(60);
    }

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
