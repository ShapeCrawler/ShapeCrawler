using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit;

public class ShapeFillTests : SCTest
{
    [Test]
    public void Fill_is_not_null()
    {
        // Arrange
        var pptx = GetInputStream("021.pptx");
        var pres = SCPresentation.Open(pptx);
        var autoShape = (IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 108);

        // Act-Assert
        autoShape.Fill.Should().NotBeNull();
    }

    [Test]
    public void Picture_SetImage_updates_picture_fill()
    {
        // Arrange
        var pptx = GetInputStream("009_table.pptx");
        var image = GetInputStream("test-image-2.png");
        var shape = (IAutoShape)SCPresentation.Open(pptx).Slides[2].Shapes.First(sp => sp.Id == 4);
        var fill = shape.Fill;
        var imageSizeBefore = fill.Picture!.BinaryData.GetAwaiter().GetResult().Length;

        // Act
        fill.Picture.SetImage(image);

        // Assert
        var imageSizeAfter = shape.Fill.Picture.BinaryData.GetAwaiter().GetResult().Length;
        imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
    }

    [Test]
    public void AutoShape_Fill_Type_returns_NoFill_When_shape_is_Not_filled()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 6);

        // Act
        var fillType = autoShape.Fill.Type;

        // Assert
        fillType.Should().Be(SCFillType.NoFill);
    }

    [Test]
    public void HexSolidColor_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 2);

        // Act
        var shapeSolidColorName = autoShape.Fill.Color;

        // Assert
        shapeSolidColorName.Should().BeEquivalentTo("ff0000");
    }

    [Test]
    public void ThemeColor_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[3].Shapes.First(sp => sp.Name == "Rectangle 3");

        // Act
        var shapeSolidColorName = autoShape.Fill.Color;

        // Assert
        shapeSolidColorName.Should().BeEquivalentTo("FFAB40");
    }

    [Test]
    public void ThemeColorWithAlpha_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeAlpha");

        // Act
        var alpha = autoShape.Fill.AlphaPercentage;

        // Assert
        alpha.Should().Be(60);
    }

    [Test]
    public void ThemeColorWithLuminanceLight_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeLumLight");

        // Act
        var fill = autoShape.Fill;

        // Assert
        fill.LuminanceModulationPercentage.Should().Be(20);
        fill.LuminanceOffsetPercentage.Should().Be(80);
    }

    [Test]
    public void ThemeColorWithLuminanceDark_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeLumDark");

       // Act
        var fill = autoShape.Fill;

        // Assert
        fill.LuminanceModulationPercentage.Should().Be(75);
        fill.LuminanceOffsetPercentage.Should().Be(0);
    }

    [Test]
    public async Task Picture_BinaryData_returns_binary_content_of_picture_image()
    {
        // Arrange
        var pptxStream = GetInputStream("009_table.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var shapeFill = pres.Slides[2].Shapes.GetByName<IAutoShape>("AutoShape 1").Fill;

        // Act
        var imageBytes = await shapeFill.Picture!.BinaryData;

        // Assert
        imageBytes.Length.Should().BePositive();
    }
    
    [Test]
    public void SetColor_sets_green_color()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];
        var shape = slide.Shapes.AddRectangle(0, 0, 100, 100);
        
        // Act
        shape.Fill!.SetColor("00FF00");

        // Assert
        PptxValidator.Validate(pres).Should().BeEmpty();
    }
}