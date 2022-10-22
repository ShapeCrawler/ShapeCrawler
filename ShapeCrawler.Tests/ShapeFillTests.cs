using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Helpers.Attributes;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests;

public class ShapeFillTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
{
    private readonly PresentationFixture _fixture;

    public ShapeFillTests(PresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Fill_is_not_null()
    {
        // Arrange
        var autoShape = (IAutoShape)_fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 108);

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
    public void SetHexSolidColor_sets_solid_color(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var shapeFill = autoShape.Fill;

        // Act
        shapeFill.SetHexSolidColor("32a852");

        // Assert
        shapeFill.HexSolidColor.Should().Be("32a852");
    }

    [Fact]
    public void Picture_SetImage_updates_picture_fill()
    {
        // Arrange
        var pres = SCPresentation.Open(TestFiles.Presentations.pre009);
        var shape = (IAutoShape)pres.Slides[2].Shapes.First(sp => sp.Id == 4);
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
        var autoShape = (IAutoShape)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 6);

        // Act
        var fillType = autoShape.Fill.Type;

        // Assert
        fillType.Should().Be(SCFillType.NoFill);
    }

    [Fact]
    public void HexSolidColor_getter_returns_color_name()
    {
        // Arrange
        var autoShape = (IAutoShape)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

        // Act
        var shapeSolidColorName = autoShape.Fill.HexSolidColor;

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