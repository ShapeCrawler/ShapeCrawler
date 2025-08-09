using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class LineShapeTests : SCTest
{
    [Test]
    public void StartPoint_returns_start_point_coordinates()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.Line("Line 1", startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);
            });
        });
        var shapes = pres.Slides[0].Shapes;
        var line = shapes.Last<ILine>();

        // Act
        var startPoint = line.StartPoint;

        // Assert
        startPoint.X.Should().Be(50);
        startPoint.Y.Should().Be(60);
    }
}