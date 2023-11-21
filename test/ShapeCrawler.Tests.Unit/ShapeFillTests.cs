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
        var pptx = StreamOf("021.pptx");
        var pres = new Presentation(pptx);
        var autoShape = pres.Slides[0].Shapes.First(sp => sp.Id == 108);

        // Act-Assert
        autoShape.Fill.Should().NotBeNull();
    }

    [Test]
    public void Picture_SetImage_updates_picture_fill()
    {
        // Arrange
        var pptx = StreamOf("009_table.pptx");
        var image = StreamOf("test-image-2.png");
        var shape = new Presentation(pptx).Slides[2].Shapes.First(sp => sp.Id == 4);
        var fill = shape.Fill;
        var imageSizeBefore = fill.Picture!.AsByteArray().Length;

        // Act
        fill.Picture.Update(image);

        // Assert
        var imageSizeAfter = shape.Fill.Picture.AsByteArray().Length;
        imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
    }

    [Test]
    public void AutoShape_Fill_Type_returns_NoFill_When_shape_is_Not_filled()
    {
        // Arrange
        var autoShape = new Presentation(StreamOf("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 6);

        // Act
        var fillType = autoShape.Fill.Type;

        // Assert
        fillType.Should().Be(FillType.NoFill);
    }

    [Test]
    public void HexSolidColor_getter_returns_color_name()
    {
        // Arrange
        var autoShape = new Presentation(StreamOf("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 2);

        // Act
        var shapeSolidColorName = autoShape.Fill.Color;

        // Assert
        shapeSolidColorName.Should().BeEquivalentTo("ff0000");
    }

    [Test]
    public void Color_Getter_returns_color_hex()
    {
        // Arrange
        var pres = new Presentation(StreamOf("009_table.pptx"));
        var shapeFill = pres.Slides[3].Shapes.First(sp => sp.Name == "Rectangle 3").Fill;

        // Act
        var colorHex = shapeFill.Color;

        // Assert
        colorHex.Should().BeEquivalentTo("FFAB40");
    }

    [Test]
    public void Alpha_returns_opacity_level_of_fill_color_in_percentages()
    {
        // Arrange
        var pres = new Presentation(StreamOf("009_table.pptx"));
        var shapeFill = pres.Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeAlpha").Fill;

        // Act
        var alpha = shapeFill.Alpha;

        // Assert
        alpha.Should().Be(60);
    }

    [Test]
    public void ThemeColorWithLuminanceLight_getter_returns_color_name()
    {
        // Arrange
        var autoShape = new Presentation(StreamOf("009_table.pptx")).Slides[3].Shapes
            .First(sp => sp.Name == "SolidSchemeLumLight");

        // Act
        var fill = autoShape.Fill;

        // Assert
        fill.LuminanceModulation.Should().Be(20);
        fill.LuminanceOffset.Should().Be(80);
    }

    [Test]
    public void Luminance_properties()
    {
        // Arrange
        var pres = new Presentation(StreamOf("009_table.pptx"));
        var shapeFill = pres.Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeLumDark").Fill;

        // Act-Assert
        shapeFill.LuminanceModulation.Should().Be(75);
        shapeFill.LuminanceOffset.Should().Be(0);
    }

    [Test]
    public void Picture_BinaryData_returns_binary_content_of_picture_image()
    {
        // Arrange
        var pptxStream = StreamOf("009_table.pptx");
        var pres = new Presentation(pptxStream);
        var shapeFill = pres.Slides[2].Shapes.GetByName("AutoShape 1").Fill;

        // Act
        var imageBytes = shapeFill.Picture!.AsByteArray();

        // Assert
        imageBytes.Length.Should().BePositive();
    }

    [Test]
    public void SetColor_sets_green_color()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddRectangle(0, 0, 100, 100);
        var shape = slide.Shapes.Last();

        // Act
        shape.Fill!.SetColor("00FF00");

        // Assert
        pres.Validate();
    }

    [Test]
    [TestCase("autoshape-case005_text-frame.pptx", 1, "AutoShape 1")]
    [TestCase("autoshape-case005_text-frame.pptx", 1, "AutoShape 2")]
    [TestCase("autoshape-grouping.pptx", 1, "AutoShape 1")]
    public void SetColor_sets_solid_color(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var shapeFill = shape.Fill;

        // Act
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        pres.Validate();
    }

    [Theory]
    // [SlideShapeData("table-case001.pptx", slideNumber: 1, shapeName: "Table 1")]
    [TestCase("table-case001.pptx", 1, "Table 1")]
    public void SetColor_sets_solid_color_as_fill_of_Table_Cell(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var table = (ITable)shape;
        var shapeFill = table[0, 0].Fill;

        // Act
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        pres.Validate();
    }

    [Test]
    [TestCase("009_table.pptx", 2, "AutoShape 2")]
    public void SetColor_sets_solid_color_After_picture(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var shapeFill = shape.Fill;
        var imageStream = StreamOf("test-image-1.png");

        // Act
        shapeFill.SetPicture(imageStream);
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        pres.Validate();
    }
}