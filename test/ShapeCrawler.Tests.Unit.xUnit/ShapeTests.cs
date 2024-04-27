using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class ShapeTests : SCTest
{
    [Theory]
    [InlineData("050_title-placeholder.pptx", 1, 2, 777)]
    [InlineData("051_title-placeholder.pptx", 1, 3074, 864)]
    public void Width_returns_width_of_Title_placeholder(
        string filename, 
        int slideNumber, 
        int shapeId,
        int expectedWidth)
    {
        // Arrange
        var pres = new Presentation(StreamOf(filename));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId);

        // Act
        var shapeWidth = shape.Width;

        // Assert
        shapeWidth.Should().Be(expectedWidth);
    }

    [Theory]
    [SlideShapeData("006_1 slides.pptx", 1, "Shape 2", 149)]
    [SlideShapeData( "009_table.pptx", 2, "Object 3", 39)]
    [SlideShapeData( "autoshape-grouping.pptx", 1, "Group 2", 108)]
    public void Height_returns_shape_height_in_pixels(IShape shape, int expectedHeight)
    {
        // Act
        var height = shape.Height;

        // Assert
        height.Should().Be(expectedHeight);
    }

    [Theory]
    [MemberData(nameof(GeometryTypeTestCases))]
    public void GeometryType_returns_shape_geometry_type(IShape shape, Geometry expectedGeometryType)
    {
        // Assert
        shape.GeometryType.Should().Be(expectedGeometryType);
    }

    [Theory]
    [SlideShapeData("054_get_shape_xpath.pptx", 1, "Title 1", "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]")]
    [SlideShapeData("054_get_shape_xpath.pptx", 1, "SubTitle 2", "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[2]")]
    public void SDKXPath_returns_shape_xpath(IShape shape, string expectedXPath)
    {
        // Act
        var shapeXPath = shape.SDKXPath;

        // Assert
        shapeXPath.Should().Be(expectedXPath);
    }

    public static IEnumerable<object[]> GeometryTypeTestCases()
    {
        var pptxStream = StreamOf("021.pptx");
        var presentation = new Presentation(pptxStream);
        var shapeCase1 = presentation.Slides[3].Shapes.First(sp => sp.Id == 2);
        var shapeCase2 = presentation.Slides[3].Shapes.First(sp => sp.Id == 3);

        yield return new object[] { shapeCase1, Geometry.Rectangle };
        yield return new object[] { shapeCase2, Geometry.Ellipse };
    }
    
    public static IEnumerable<object[]> TestCasesGetShapeById()
    {
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, 1, null
        };
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, 2, "Title 1"
        };
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, 3, "SubTitle 2"
        };
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, 4, null
        };
    }

    [Theory]
    [MemberData(nameof(TestCasesGetShapeById))]
    public void TryGetSlideShapeById(string presentationName, int slideNumber, int shapeId, string? expectedShapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presentationName));
        var slide = pres.Slides[slideNumber];
        var shape = slide.Shapes.TryGetById<IShape>(shapeId);

        // Act
        var shapeName = shape?.Name;

        // Assert
        shapeName.Should().Be(expectedShapeName);
    }

    public static IEnumerable<object[]> TestCasesGetShapeByName()
    {
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, "Foo", null
        };
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, "Title 1", 2
        };
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, "SubTitle 2", 3
        };
        yield return new object[]
        {
            "054_get_shape_xpath.pptx", 0, "Bar", null
        };
    }

    [Theory]
    [MemberData(nameof(TestCasesGetShapeByName))]
    public void TryGetSlideShapeByName(string presentationName, int slideNumber, string shapeName,int? expectedShapeId)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presentationName));
        var slide = pres.Slides[slideNumber];
        var shape = slide.Shapes.TryGetByName<IShape>(shapeName);

        // Act
        var shapeId = shape?.Id;

        // Assert
        shapeId.Should().Be(expectedShapeId);
    }
}