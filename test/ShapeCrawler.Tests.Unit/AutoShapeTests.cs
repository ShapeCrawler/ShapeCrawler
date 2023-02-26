using System.Linq;
using FluentAssertions;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class AutoShapeTests
{
    [Fact(Skip = "In Progress")]
    public void Duplicate_duplicates_AutoShape()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var shapes = pres.Slides[0].Shapes;
        var autoShapes = shapes.AutoShapes;
        autoShapes.AddRectangle(10, 20, 30, 40);
        var autoShape = autoShapes.Single();

        // Act
        var autoShapeCopy = autoShape.Duplicate();

        // Assert
        shapes.Should().HaveCount(2);
        autoShapes.Should().HaveCount(2);
        shapes.AutoShapes.Should().HaveCount(2);
        autoShapeCopy.Id.Should().Be(2, "because it is the second shape in the collection");
    }
}