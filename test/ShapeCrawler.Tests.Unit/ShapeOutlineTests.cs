using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;

namespace ShapeCrawler.Tests.Unit;

public class ShapeOutlineTests : SCTest
{
    [Xunit.Theory]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 4", 0)]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 6", 0.25)]
    [SlideShapeData("020.pptx", 1, "Shape 1", 0)]
    public void Weight_Getter_returns_outline_weight_in_points(IShape shape, double expectedWeight)
    {
        // Arrange
        var autoShape = (IShape)shape;
        
        // Act
        var outlineWeight = autoShape.Outline.Weight;
        
        // Assert
        outlineWeight.Should().Be(expectedWeight);
    }

    [Test]
    [TestCase("autoshape-grouping.pptx", 1, "TextBox 4")]
    [TestCase("020.pptx", 1, "Shape 1")]
    [TestCase("autoshape-case011_save-as-png.pptx", 1, "AutoShape 1")]
    public void Weight_Setter_sets_outline_weight_in_points(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new SCPresentation(StreamOf(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var outline = shape.Outline;
        
        // Act
        outline.Weight = 0.25;

        // Assert
        outline.Weight.Should().Be(0.25);
        pres.Validate();
    }

    [Xunit.Theory]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 6", "000000")]
    public void Color_Getter_returns_outline_color_in_hex_format(IShape shape, string expectedColor)
    {
        // Arrange
        var autoShape = (IShape)shape;
        var outline = autoShape.Outline;
        
        // Act
        var outlineColor = outline.HexColor;
        
        // Assert
        outlineColor.Should().Be(expectedColor);
    }
    
    [Test]
    [TestCase("autoshape-grouping.pptx", 1, "TextBox 6")]
    public void Color_Setter_sets_outline_color(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new SCPresentation(StreamOf(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var outline = shape.Outline;
        
        // Act
        outline.HexColor = "be3455";

        // Assert
        outline.HexColor.Should().Be("be3455");
        pres.Validate();
    }
}