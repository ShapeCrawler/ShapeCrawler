using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class AutoShapeTests : SCTest
{
    [Fact]
    public void Duplicate_duplicates_AutoShape()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddRectangle(10, 20, 30, 40);
        var autoShape = (IAutoShape)shapes.Single();

        // Act
        var autoShapeCopy = autoShape.Duplicate();

        // Assert
        shapes.Should().HaveCount(2);
        autoShapeCopy.Id.Should().Be(2, "because it is the second shape in the collection");
    }
}