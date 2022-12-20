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

    [Fact]
    public void Theme_FontSettings_Head_Getter_returns_name_of_theme_heading_font()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var headingFontName = slideMaster.Theme.FontScheme.Head;

        // Assert
        headingFontName.Should().Be("Arial");
    }
    
    [Fact]
    public void Theme_FontSettings_Head_Setter_sets_name_of_theme_heading_font()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.Head = "Times New Roman";

        // Assert
        slideMaster.Theme.FontScheme.Head.Should().Be("Times New Roman");
    }
    
    [Fact]
    public void Theme_FontSettings_Body_Getter_returns_name_of_theme_body_font()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var bodyFontName = slideMaster.Theme.FontScheme.Body;

        // Assert
        bodyFontName.Should().Be("Times New Roman");
    }
    
    [Fact]
    public void Theme_FontSettings_Body_Setter_sets_name_of_theme_body_font()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.Body = "Arial";

        // Assert
        slideMaster.Theme.FontScheme.Body.Should().Be("Arial");
    }

    [Fact]
    public void Theme_ColorScheme_color_getter_returns_scheme_with_colors()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var dark1 = slideMaster.Theme.ColorScheme.Dark1;
        var light1 = slideMaster.Theme.ColorScheme.Light1;
        var dark2 = slideMaster.Theme.ColorScheme.Dark2;
        var light2 = slideMaster.Theme.ColorScheme.Light2;
        var accent1 = slideMaster.Theme.ColorScheme.Accent1;
        var accent2 = slideMaster.Theme.ColorScheme.Accent2;
        var accent3 = slideMaster.Theme.ColorScheme.Accent3;
        var accent4 = slideMaster.Theme.ColorScheme.Accent4;
        var accent5 = slideMaster.Theme.ColorScheme.Accent5;
        var accent6 = slideMaster.Theme.ColorScheme.Accent6;
        var hyperlink = slideMaster.Theme.ColorScheme.Hyperlink;
        var followedHyperlink = slideMaster.Theme.ColorScheme.FollowedHyperlink;

        // Assert
        dark1.Should().Be("000000");
        light1.Should().Be("FFFFFF");
        dark2.Should().Be("1F497D");
        light2.Should().Be("EEECE1");
        accent1.Should().Be("4F81BD");
        accent2.Should().Be("C0504D");
        accent3.Should().Be("9BBB59");
        accent4.Should().Be("8064A2");
        accent5.Should().Be("4BACC6");
        accent6.Should().Be("F79646");
        hyperlink.Should().Be("002857");
        followedHyperlink.Should().Be("800080");
    }
    
    [Fact]
    public void Theme_ColorScheme_color_setter_sets_scheme_color()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.ColorScheme.Dark1 = "FFC0CB";
        slideMaster.Theme.ColorScheme.Light2 = "FFC0CB";
        slideMaster.Theme.ColorScheme.Accent1 = "FFC0CB";
        slideMaster.Theme.ColorScheme.Hyperlink = "FFC0CB";
        slideMaster.Theme.ColorScheme.FollowedHyperlink = "FFC0CB";

        // Assert
        slideMaster.Theme.ColorScheme.Dark1.Should().Be("FFC0CB");
        slideMaster.Theme.ColorScheme.Light2.Should().Be("FFC0CB");
        slideMaster.Theme.ColorScheme.Accent1.Should().Be("FFC0CB");
        slideMaster.Theme.ColorScheme.Hyperlink.Should().Be("FFC0CB");
        slideMaster.Theme.ColorScheme.FollowedHyperlink.Should().Be("FFC0CB");
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }
}