using FluentAssertions;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class LineShapeTests
{
    [Fact]
    public void StartPoint_returns_start_point_coordinates()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var shapes = pres.Slides[0].Shapes;
        var line = shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Act
        var startPoint = line.StartPoint;

        // Assert
        startPoint.X.Should().Be(50);
        startPoint.Y.Should().Be(60);
    }
    
    [Fact]
    public void EndPoint_returns_end_point_coordinates()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var shapes = pres.Slides[0].Shapes;
        var line = shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Act
        var endPoint = line.EndPoint;

        // Assert
        endPoint.X.Should().Be(100);
        endPoint.Y.Should().Be(60);
    }
}