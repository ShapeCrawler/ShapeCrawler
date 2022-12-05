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
    [SlideShapeData("020.pptx", 1, "Shape 1", 0)]
    public void Weight_Getter_returns_outline_weight_in_points(IShape shape, double expectedWeight)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        
        // Act
        var outlineWeight = autoShape.Outline.Weight;
        
        // Assert
        outlineWeight.Should().Be(expectedWeight);
    }

    [Theory]
    [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 4")]
    [SlideShapeData("020.pptx", 1, "Shape 1")]
    [SlideShapeData("autoshape-case011_save-as-png.pptx", 1, "AutoShape 1")]
    public void Weight_Setter_sets_outline_weight_in_points(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var outline = autoShape.Outline;
        
        // Act
        outline.Weight = 0.25;

        // Assert
        outline.Weight.Should().Be(0.25);
        var errors = PptxValidator.Validate(autoShape.SlideObject.Presentation);
        errors.Should().BeEmpty();
    }
}