using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Helpers.Attributes;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests;

public class ShapeFillTests : ShapeCrawlerTest
{
    [Fact]
    public void Fill_is_not_null()
    {
        // Arrange
        var pptx = GetTestStream("021.pptx");
        var pres = SCPresentation.Open(pptx);
        var autoShape = (IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 108);

        // Act-Assert
        autoShape.Fill.Should().NotBeNull();
    }

    [Theory]
    [SlideShapeData("008.pptx", slideNumber: 1, shapeName: "AutoShape 1")]
    [SlideShapeData("autoshape-case009.pptx", slideNumber: 1, shapeName: "AutoShape 1")]
    [LayoutShapeData("autoshape-case003.pptx", slideNumber: 1, shapeName: "AutoShape 1")]
    [MasterShapeData("autoshape-case003.pptx", shapeName: "AutoShape 1")]
    public void SetPicture_updates_fill_with_specified_picture_image_When_shape_is_Not_filled(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var fill = autoShape.Fill;
        var imageStream = GetTestStream("test-image-1.png");

        // Act
        fill.SetPicture(imageStream);

        // Assert
        var pictureBytes = fill.Picture!.BinaryData.Result;
        var imageBytes = imageStream.ToArray();
        pictureBytes.SequenceEqual(imageBytes).Should().BeTrue();
    }

    [Theory]
    [SlideShapeData("autoshape-case005_text-frame.pptx", slideNumber: 1, shapeName: "AutoShape 1")]
    public void SetColor_sets_solid_color(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var shapeFill = autoShape.Fill;

        // Act
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
        errors.Should().BeEmpty();
    }
    
    [Theory]
    [SlideShapeData("table-case001.pptx", slideNumber: 1, shapeName: "Table 1")]
    public void SetColor_sets_solid_color_as_fill_of_Table_Cell(IShape shape)
    {
        // Arrange
        var table = (ITable)shape;
        var shapeFill = table[0, 0].Fill;

        // Act
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
        errors.Should().BeEmpty();
    }
    
    [Theory]
    [SlideShapeData("009_table.pptx", slideNumber: 2, shapeName: "AutoShape 2")]
    public void SetHexSolidColor_sets_solid_color_After_picture(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var shapeFill = autoShape.Fill;
        var imageStream = GetTestStream("test-image-1.png");

        // Act
        shapeFill.SetPicture(imageStream);
        shapeFill.SetColor("32a852");
        
        // Assert
        shapeFill.Color.Should().Be("32a852");
        var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void Picture_SetImage_updates_picture_fill()
    {
        // Arrange
        var shape = (IAutoShape)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4);
        var fill = shape.Fill;
        var newImage = TestFiles.Images.img02_stream;
        var imageSizeBefore = fill.Picture!.BinaryData.GetAwaiter().GetResult().Length;

        // Act
        fill.Picture.SetImage(newImage);

        // Assert
        var imageSizeAfter = shape.Fill.Picture.BinaryData.GetAwaiter().GetResult().Length;
        imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
    }

    [Theory]
    [MemberData(nameof(TestCasesFillType))]
    public void Type_returns_fill_type(IAutoShape shape, SCFillType expectedFill)
    {
        // Act
        var fillType = shape.Fill.Type;

        // Assert
        fillType.Should().Be(expectedFill);
    }

    public static IEnumerable<object[]> TestCasesFillType()
    {
        var pptxStream = GetTestStream("009_table.pptx");
        var pres = SCPresentation.Open(pptxStream);

        var withNoFill = pres.Slides[1].Shapes.GetById<IAutoShape>(6);
        yield return new object[] { withNoFill, SCFillType.NoFill };

        var withSolid = pres.Slides[1].Shapes.GetById<IAutoShape>(2);
        yield return new object[] { withSolid, SCFillType.Solid };

        var withGradient = pres.Slides[1].Shapes.GetByName<IAutoShape>("AutoShape 1");
        yield return new object[] { withGradient, SCFillType.Gradient };

        var withPicture = pres.Slides[2].Shapes.GetById<IAutoShape>(4);
        yield return new object[] { withPicture, SCFillType.Picture };

        var withPattern = pres.Slides[1].Shapes.GetByName<IAutoShape>("AutoShape 2");
        yield return new object[] { withPattern, SCFillType.Pattern };

        pptxStream = GetTestStream("autoshape-case003.pptx");
        pres = SCPresentation.Open(pptxStream);
        var withSlideBg = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
        yield return new object[] { withSlideBg, SCFillType.SlideBackground };
    }

    [Fact]
    public void AutoShape_Fill_Type_returns_NoFill_When_shape_is_Not_filled()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 6);

        // Act
        var fillType = autoShape.Fill.Type;

        // Assert
        fillType.Should().Be(SCFillType.NoFill);
    }

    [Fact]
    public void HexSolidColor_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 2);

        // Act
        var shapeSolidColorName = autoShape.Fill.Color;

        // Assert
        shapeSolidColorName.Should().BeEquivalentTo("ff0000");
    }

    [Fact]
    public async void Picture_BinaryData_returns_binary_content_of_picture_image()
    {
        // Arrange
        var pptxStream = GetTestStream("009_table.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var shapeFill = pres.Slides[2].Shapes.GetByName<IAutoShape>("AutoShape 1").Fill;

        // Act
        var imageBytes = await shapeFill.Picture!.BinaryData;

        // Assert
        imageBytes.Length.Should().BePositive();
    }
}