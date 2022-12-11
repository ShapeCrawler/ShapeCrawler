using System.Linq;
using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using Xunit;

namespace ShapeCrawler.Tests;

public class SlideMasterTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
{
    private readonly PresentationFixture _fixture;

    public SlideMasterTests(PresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void ShapeXAndY_ReturnXAndYAxesCoordinatesOfTheMasterShape()
    {
        // Arrange
        ISlideMaster slideMaster = _fixture.Pre001.SlideMasters[0];
        IShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);

        // Act
        int shapeXCoordinate = shape.X;
        int shapeYCoordinate = shape.Y;

        // Assert
        shapeXCoordinate.Should().Be((int)(838200 * TestHelper.HorizontalResolution / 914400));
        shapeYCoordinate.Should().Be((int)(365125 * TestHelper.VerticalResolution / 914400));
    }

    [Fact]
    public void ShapeWidthAndHeight_ReturnWidthAndHeightSizesOfTheMaster()
    {
        // Arrange
        ISlideMaster slideMaster = _fixture.Pre001.SlideMasters[0];
        IShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);
        float horizontalResolution = TestHelper.HorizontalResolution;
        float verticalResolution = TestHelper.VerticalResolution;

        // Act
        int shapeWidth = shape.Width;
        int shapeHeight = shape.Height;

        // Assert
        shapeWidth.Should().Be((int)(10515600 * horizontalResolution / 914400));
        shapeHeight.Should().Be((int)(1325563 * verticalResolution / 914400));
    }

    [Fact]
    public void SlideLayout_Name_returns_name_of_slide_layout()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case011_save-as-png.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var layoutName = slideMaster.SlideLayouts[0].Name;

        // Assert
        layoutName.Should().Be("Title Slide");
    }
    
    [Fact]
    public void AutoShapePlaceholderType_ReturnsPlaceholderType()
    {
        // Arrange
        ISlideMaster slideMaster = _fixture.Pre001.SlideMasters[0];
        IShape masterAutoShapeCase1 = slideMaster.Shapes.First(sp => sp.Id == 2);
        IShape masterAutoShapeCase2 = slideMaster.Shapes.First(sp => sp.Id == 8);
        IShape masterAutoShapeCase3 = slideMaster.Shapes.First(sp => sp.Id == 7);

        // Act
        SCPlaceholderType? shapePlaceholderTypeCase1 = masterAutoShapeCase1.Placeholder?.Type;
        SCPlaceholderType? shapePlaceholderTypeCase2 = masterAutoShapeCase2.Placeholder?.Type;
        SCPlaceholderType? shapePlaceholderTypeCase3 = masterAutoShapeCase3.Placeholder?.Type;

        // Assert
        shapePlaceholderTypeCase1.Should().Be(SCPlaceholderType.Title);
        shapePlaceholderTypeCase2.Should().BeNull();
        shapePlaceholderTypeCase3.Should().BeNull();
    }

    [Fact]
    public void ShapeGeometryType_ReturnsShapesGeometryFormType()
    {
        // Arrange
        ISlideMaster slideMaster = _fixture.Pre001.SlideMasters[0];
        IShape shapeCase1 = slideMaster.Shapes.First(sp => sp.Id == 2);
        IShape shapeCase2 = slideMaster.Shapes.First(sp => sp.Id == 8);

        // Act
        SCGeometry geometryTypeCase1 = shapeCase1.GeometryType;
        SCGeometry geometryTypeCase2 = shapeCase2.GeometryType;

        // Assert
        geometryTypeCase1.Should().Be(SCGeometry.Rectangle);
        geometryTypeCase2.Should().Be(SCGeometry.Custom);
    }

    [Fact]
    public void AutoShapeTextBoxText_ReturnsText_WhenTheSlideMasterAutoShapesTextBoxIsNotEmpty()
    {
        // Arrange
        ISlideMaster slideMaster = _fixture.Pre001.SlideMasters[0];
        IAutoShape autoShape = (IAutoShape)slideMaster.Shapes.First(sp => sp.Id == 8);

        // Act-Assert
        autoShape.TextFrame.Text.Should().BeEquivalentTo("id8");
    }
}