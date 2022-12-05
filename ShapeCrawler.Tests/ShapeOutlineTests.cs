using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests;

public class ShapeOutlineTests : ShapeCrawlerTest
{
    [Theory]
    [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 4", 0)]
    [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 6", 0.25)]
    public void Weight_Getter_returns_outline_weight_in_points(IShape shape, double expectedWeight)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        
        // Act
        var outlineWeight = autoShape.Outline.Weight;
        
        // Assert
        outlineWeight.Should().Be(expectedWeight);
    }
}