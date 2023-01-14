using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Helpers.Attributes;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

namespace ShapeCrawler.UnitTests;

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

    [Theory]
    [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 6", "000000")]
    public void Color_Getter_returns_outline_color_in_hex_format(IShape shape, string expectedColor)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var outline = autoShape.Outline;
        
        // Act
        var outlineColor = outline.Color;
        
        // Assert
        outlineColor.Should().Be(expectedColor);
    }
    
    [Theory]
    [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 6")]
    public void Color_Setter_sets_outline_color(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var outline = autoShape.Outline;
        
        // Act
        outline.Color = "be3455";

        // Assert
        outline.Color.Should().Be("be3455");
        var errors = PptxValidator.Validate(autoShape.SlideObject.Presentation);
        errors.Should().BeEmpty();
    }
}