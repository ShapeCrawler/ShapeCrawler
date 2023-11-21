using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class ShapeTests : SCTest
{
    [Theory]
    [MemberData(nameof(TestCasesXGetter))]
    public void X_Getter_returns_x_coordinate_in_pixels(TestCase<IShape, int> testCase)
    {
        // Arrange
        var shape = testCase.Param1;
        var expectedX = testCase.Param2;
            
        // Act
        var xCoordinate = shape.X;
            
        // Assert
        xCoordinate.Should().Be(expectedX);
    }

    public static IEnumerable<object[]> TestCasesXGetter
    {
        get
        {
            var pptxStream1 = StreamOf("021.pptx");
            var pres1 = new Presentation(pptxStream1);
            var shape1 = pres1.Slides[3].Shapes.GetById<IShape>(2);
            var testCase1 = new TestCase<IShape, int>(1, shape1, 383);
            yield return new object[] { testCase1 };
                
            var pptxStream2 = StreamOf("008.pptx");
            var pres2 = new Presentation(pptxStream2);
            var shape2 = pres2.Slides[0].Shapes.GetById<IShape>(3);
            var testCase2 = new TestCase<IShape, int>(2, shape2, 66);
            yield return new object[] { testCase2 };
                
            var pptxStream3 = StreamOf("006_1 slides.pptx");
            var pres3 = new Presentation(pptxStream3);
            var shape3 = pres3.Slides[0].Shapes.GetById<IShape>(2);
            var testCase3 = new TestCase<IShape, int>(3, shape3, 160);
            yield return new object[] { testCase3 };
                
            var pptxStream4 = StreamOf("009_table.pptx");
            var pres4 = new Presentation(pptxStream4);
            var shape4 = pres4.Slides[1].Shapes.GetById<IShape>(9);
            var testCase4 = new TestCase<IShape, int>(4, shape4, 73);
            yield return new object[] { testCase4 };
                
            var pptxStream5 = StreamOf("025_chart.pptx");
            var pres5 = new Presentation(pptxStream5);
            var shape5 = pres5.Slides[2].Shapes.GetById<IShape>(7);
            var testCase5 = new TestCase<IShape, int>(5, shape5, 79);
            yield return new object[] { testCase5 };
                
            var pptxStream6 = StreamOf("018.pptx");
            var pres6 = new Presentation(pptxStream6);
            var shape6 = pres6.Slides[0].Shapes.GetByName<IShape>("Picture Placeholder 1");
            var testCase6 = new TestCase<IShape, int>(6, shape6, 9);
            yield return new object[] { testCase6 };
                
            var pptxStream7 = StreamOf("009_table.pptx");
            var pres7 = new Presentation(pptxStream7);
            var shape7 = pres7.Slides[1].Shapes.GetByName<IGroupShape>("Group 1").Shapes.GetByName<IShape>("Shape 1");
            var testCase7 = new TestCase<IShape, int>(7, shape7, 53);
            yield return new object[] { testCase7 };
        }
    }

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
}