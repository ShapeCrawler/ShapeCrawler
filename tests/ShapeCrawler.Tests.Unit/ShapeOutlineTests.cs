using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class ShapeOutlineTests : SCTest
{
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 4", 0)]
    [SlideShape("020.pptx", 1, "Shape 1", 0)]
    public void Weight_Getter_returns_outline_weight_in_points(IShape shape, int expectedWeight)
    {
        // Arrange
        var autoShape = shape;
        
        // Act
        var outlineWeight = autoShape.Outline.Weight;
        
        // Assert
        outlineWeight.Should().Be(expectedWeight);
    }
    
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 6", 0.25)]
    public void Weight_Getter_returns_outline_weight_in_decimal_points(IShape shape, double expectedWeight)
    {
        // Arrange
        var autoShape = shape;
        
        // Act
        var outlineWeight = autoShape.Outline.Weight;
        
        // Assert
        outlineWeight.Should().Be((decimal)expectedWeight);
    }
    
    [Test]
    [TestCase("autoshape-grouping.pptx", 1, "TextBox 4")]
    [TestCase("020.pptx", 1, "Shape 1")]
    [TestCase("autoshape-case011_save-as-png.pptx", 1, "AutoShape 1")]
    public void Weight_Setter_sets_outline_weight_in_points(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var outline = shape.Outline;
        
        // Act
        outline.Weight = 0.25m;

        // Assert
        outline.Weight.Should().Be(0.25m);
        pres.Validate();
    }

    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 6", "000000")]
    public void HexColor_Getter_returns_outline_color_in_hex_format(IShape shape, string expectedColor)
    {
        // Arrange
        var outline = shape.Outline;
        
        // Act-Assert
        outline.HexColor.Should().Be(expectedColor);
    }
    
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 4")]
    public void HexColor_Getter_returns_null_for_NoOutline(IShape shape)
    {
        // Arrange
        var outline = shape.Outline;
        
        // Act-Assert
        outline.HexColor.Should().BeNull();
    }
    
    [Test]
    [TestCase("autoshape-grouping.pptx", 1, "TextBox 6")]
    public void SetHexColor_sets_outline_color(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var outline = shape.Outline;
        
        // Act
        outline.SetHexColor("be3455");

        // Assert
        outline.HexColor.Should().Be("be3455");
        pres.Validate();
    }

    [Test]
    [TestCase("autoshape-grouping.pptx", 1, "TextBox 6")]
    public void SetNoOutline_removes_outline_color(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var outline = shape.Outline;
        
        // Act
        outline.SetNoOutline();

        // Assert
        outline.HexColor.Should().BeNull();
        pres.Validate();
    }
}