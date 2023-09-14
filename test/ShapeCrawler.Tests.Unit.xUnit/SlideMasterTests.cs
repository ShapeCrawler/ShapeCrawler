using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class SlideMasterTests : SCTest
{
    [Fact]
    public void ShapeXAndY_ReturnXAndYAxesCoordinatesOfTheMasterShape()
    {
        // Arrange
        var pptx = StreamOf("001.pptx");
        var pres = new SCPresentation(pptx);
        ISlideMaster slideMaster = pres.SlideMasters[0];
       IShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);

        // Act
        int shapeXCoordinate = shape.X;
        int shapeYCoordinate = shape.Y;

        // Assert
        shapeXCoordinate.Should().Be((int)(838200 * Helpers.TestHelper.HorizontalResolution / 914400));
        shapeYCoordinate.Should().Be((int)(365125 * Helpers.TestHelper.VerticalResolution / 914400));
    }

    [Fact]
    public void ShapeWidthAndHeight_ReturnWidthAndHeightSizesOfTheMaster()
    {
        // Arrange
        var pptx = StreamOf("001.pptx");
        var pres = new SCPresentation(pptx);
        ISlideMaster slideMaster = pres.SlideMasters[0];
        IShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);
        float horizontalResolution = Helpers.TestHelper.HorizontalResolution;
        float verticalResolution = Helpers.TestHelper.VerticalResolution;

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
        var pptx = StreamOf("autoshape-case011_save-as-png.pptx");
        var pres = new SCPresentation(pptx);
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
        ISlideMaster slideMaster = new SCPresentation(StreamOf("001.pptx")).SlideMasters[0];
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
        var pptx = StreamOf("001.pptx");
        var pres = new SCPresentation(pptx);
        ISlideMaster slideMaster = pres.SlideMasters[0];
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
        ISlideMaster slideMaster = new SCPresentation(StreamOf("001.pptx")).SlideMasters[0];
        IShape autoShape = (IShape)slideMaster.Shapes.First(sp => sp.Id == 8);

        // Act-Assert
        autoShape.TextFrame.Text.Should().BeEquivalentTo("id8");
    }

    [Fact]
    public void Theme_FontScheme_HeadLatinFont_Getter_returns_font_name_for_the_Latin_characters_of_Heading()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var headingFontName = slideMaster.Theme.FontScheme.HeadLatinFont;

        // Assert
        headingFontName.Should().Be("Arial");
    }
    
    [Fact]
    public void Theme_FontScheme_HeadEastAsianFont_Getter_returns_font_name_for_the_EastAsian_characters_of_Heading()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var fontName = slideMaster.Theme.FontScheme.HeadEastAsianFont;

        // Assert
        fontName.Should().Be("SimSun");
    }
    
    [Fact]
    public void Theme_FontScheme_HeadLatinFont_Setter_sets_font_name_for_the_Latin_characters_of_heading()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.HeadLatinFont = "Times New Roman";

        // Assert
        slideMaster.Theme.FontScheme.HeadLatinFont.Should().Be("Times New Roman");
    }
    
    [Fact]
    public void Theme_FontScheme_HeadEastAsianFont_Setter_sets_font_for_the_East_Asian_characters_of_heading()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.HeadEastAsianFont = "MingLiU-ExtB";

        // Assert
        slideMaster.Theme.FontScheme.HeadEastAsianFont.Should().Be("MingLiU-ExtB");
        pres.Validate();
    }
    
    [Fact]
    public void Theme_FontScheme_BodyLatinFont_Getter_returns_font_name_for_the_Latin_characters_of_body()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var bodyFontName = slideMaster.Theme.FontScheme.BodyLatinFont;

        // Assert
        bodyFontName.Should().Be("Times New Roman");
    }
    
    [Fact]
    public void Theme_FontScheme_BodyEastAsianFont_Getter_returns_font_for_the_EastAsian_characters_of_body()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var bodyFontName = slideMaster.Theme.FontScheme.BodyEastAsianFont;

        // Assert
        bodyFontName.Should().Be("MingLiU-ExtB");
    }
    
    [Fact]
    public void Theme_FontScheme_BodyLatinFont_Setter_font_name_for_the_Latin_characters_of_body()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.BodyLatinFont = "Arial";

        // Assert
        slideMaster.Theme.FontScheme.BodyLatinFont.Should().Be("Arial");
    }
    
    [Fact]
    public void Theme_FontScheme_BodyEastAsianFont_Setter_font_for_the_East_Asian_characters_of_body()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.BodyEastAsianFont = "SimSun";

        // Assert
        slideMaster.Theme.FontScheme.BodyEastAsianFont.Should().Be("SimSun");
    }

    [Fact]
    public void Theme_ColorScheme_color_getter_returns_scheme_with_colors()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
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
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
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
        pres.Validate();
    }
}