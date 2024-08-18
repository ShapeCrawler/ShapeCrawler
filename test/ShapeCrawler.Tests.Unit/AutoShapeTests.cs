using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class AutoShapeTests : SCTest
{
    [Test]
    public void Duplicate_duplicates_AutoShape()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddRectangle(10, 20, 30, 40);
        var rtSlideShape = (IRootShape)shapes.Single();

        // Act
        rtSlideShape.Duplicate();

        // Assert
        var autoShapeCopy = (IRootShape)shapes.Last(); 
        shapes.Should().HaveCount(2);
        autoShapeCopy.Id.Should().Be(2, "because it is the second shape in the collection");
    }
}