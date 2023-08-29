using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class ShapeTests : SCTest
{
    [Theory]
    [SlideShapeData("021.pptx", 4, 2, SCPlaceholderType.Footer)]
    [SlideShapeData("008.pptx", 1, 3, SCPlaceholderType.DateAndTime)]
    [SlideShapeData("019.pptx", 1, 2, SCPlaceholderType.SlideNumber)]
    [SlideShapeData("013.pptx", 1, 281, SCPlaceholderType.Content)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Content Placeholder 1", SCPlaceholderType.Content)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Text Placeholder 1", SCPlaceholderType.Text)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Picture Placeholder 1", SCPlaceholderType.Picture)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Table Placeholder 1", SCPlaceholderType.Table)]
    [SlideShapeData("autoshape-case016.pptx", 1, "SmartArt Placeholder 1", SCPlaceholderType.SmartArt)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Media Placeholder 1", SCPlaceholderType.Media)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Online Image Placeholder 1", SCPlaceholderType.OnlineImage)]
    public void PlaceholderType_returns_placeholder_type(IShape shape, SCPlaceholderType expectedType)
    {
        // Act
        var placeholderType = shape.Placeholder!.Type;

        // Assert
        placeholderType.Should().Be(expectedType);
    }

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
            var pres1 = new SCPresentation(pptxStream1);
            var shape1 = pres1.Slides[3].Shapes.GetById<IShape>(2);
            var testCase1 = new TestCase<IShape, int>(1, shape1, 383);
            yield return new object[] { testCase1 };
                
            var pptxStream2 = StreamOf("008.pptx");
            var pres2 = new SCPresentation(pptxStream2);
            var shape2 = pres2.Slides[0].Shapes.GetById<IShape>(3);
            var testCase2 = new TestCase<IShape, int>(2, shape2, 66);
            yield return new object[] { testCase2 };
                
            var pptxStream3 = StreamOf("006_1 slides.pptx");
            var pres3 = new SCPresentation(pptxStream3);
            var shape3 = pres3.Slides[0].Shapes.GetById<IShape>(2);
            var testCase3 = new TestCase<IShape, int>(3, shape3, 160);
            yield return new object[] { testCase3 };
                
            var pptxStream4 = StreamOf("009_table.pptx");
            var pres4 = new SCPresentation(pptxStream4);
            var shape4 = pres4.Slides[1].Shapes.GetById<IShape>(9);
            var testCase4 = new TestCase<IShape, int>(4, shape4, 73);
            yield return new object[] { testCase4 };
                
            var pptxStream5 = StreamOf("025_chart.pptx");
            var pres5 = new SCPresentation(pptxStream5);
            var shape5 = pres5.Slides[2].Shapes.GetById<IShape>(7);
            var testCase5 = new TestCase<IShape, int>(5, shape5, 79);
            yield return new object[] { testCase5 };
                
            var pptxStream6 = StreamOf("018.pptx");
            var pres6 = new SCPresentation(pptxStream6);
            var shape6 = pres6.Slides[0].Shapes.GetByName<IShape>("Picture Placeholder 1");
            var testCase6 = new TestCase<IShape, int>(6, shape6, 9);
            yield return new object[] { testCase6 };
                
            var pptxStream7 = StreamOf("009_table.pptx");
            var pres7 = new SCPresentation(pptxStream7);
            var shape7 = pres7.Slides[1].Shapes.GetByName<IGroupShape>("Group 1").Shapes.GetByName<IShape>("Shape 1");
            var testCase7 = new TestCase<IShape, int>(7, shape7, 53);
            yield return new object[] { testCase7 };
        }
    }

    [Theory]
    [InlineData("050_title-placeholder.pptx", 1, 2, 777)]
    [InlineData("051_title-placeholder.pptx", 1, 3074, 864)]
    public void Width_returns_width_of_Title_placeholder(string filename, int slideNumber, int shapeId,
        int expectedWidth)
    {
        // Arrange
        var pptx = StreamOf(filename);
        var pres = new SCPresentation(pptx);
        var autoShape = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId);

        // Act
        var shapeWidth = autoShape.Width;

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
    public void GeometryType_returns_shape_geometry_type(IShape shape, SCGeometry expectedGeometryType)
    {
        // Assert
        shape.GeometryType.Should().Be(expectedGeometryType);
    }

    public static IEnumerable<object[]> GeometryTypeTestCases()
    {
        var pptxStream = StreamOf("021.pptx");
        var presentation = new SCPresentation(pptxStream);
        var shapeCase1 = presentation.Slides[3].Shapes.First(sp => sp.Id == 2);
        var shapeCase2 = presentation.Slides[3].Shapes.First(sp => sp.Id == 3);

        yield return new object[] { shapeCase1, SCGeometry.Rectangle };
        yield return new object[] { shapeCase2, SCGeometry.Ellipse };
    }
}