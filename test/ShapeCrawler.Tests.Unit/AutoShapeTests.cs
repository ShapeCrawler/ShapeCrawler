using System.Linq;
using FluentAssertions;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class AutoShapeTests
{
    #if DEBUG
    [Fact(Skip = "Not implemented yet")]
    public void Duplicate_duplicates_AutoShape()
    {
        // Arrange
        var pres = SCPresentation.Create();
        pres.Slides[0].Shapes.AutoShapes.AddRectangle(10, 20, 30, 40);
        var shapes = pres.Slides[0].Shapes;
        var autoShape = shapes.AutoShapes.Single();

        // Act
        var autoShapeCopy = autoShape.Duplicate();

        // Assert
        shapes.AutoShapes.Should().HaveCount(2);
        autoShapeCopy.Id.Should().Be(2, "because it is the second shape in the collection");
    }
#endif
}