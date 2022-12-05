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

    [Fact]
    public void Weight_Setter_sets_outline_weight_in_points()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var outline = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 4").Outline;
        
        // Act
        outline.Weight = 0.25;

        // Assert
        outline.Weight.Should().Be(0.25);
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }
}