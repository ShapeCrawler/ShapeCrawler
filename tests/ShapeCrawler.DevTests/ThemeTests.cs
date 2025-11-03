using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class ThemeTests : SCTest
{
    [Test]
    public void Theme_FontScheme_HeadLatinFont_Getter_returns_font_name_for_the_Latin_characters_of_Heading()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var headingFontName = slideMaster.Theme.FontScheme.HeadLatinFont;

        // Assert
        headingFontName.Should().Be("Arial");
    }

    [Test]
    public void Theme_FontScheme_HeadEastAsianFont_Getter_returns_font_name_for_the_EastAsian_characters_of_Heading()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var fontName = slideMaster.Theme.FontScheme.HeadEastAsianFont;

        // Assert
        fontName.Should().Be("SimSun");
    }

    [Test]
    public void Theme_FontScheme_HeadLatinFont_Setter_sets_font_name_for_the_Latin_characters_of_heading()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.HeadLatinFont = "Times New Roman";

        // Assert
        slideMaster.Theme.FontScheme.HeadLatinFont.Should().Be("Times New Roman");
    }

    [Test]
    public void Theme_FontScheme_HeadEastAsianFont_Setter_sets_font_for_the_East_Asian_characters_of_heading()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.HeadEastAsianFont = "MingLiU-ExtB";

        // Assert
        slideMaster.Theme.FontScheme.HeadEastAsianFont.Should().Be("MingLiU-ExtB");
        pres.Validate();
    }

    [Test]
    public void Theme_FontScheme_BodyLatinFont_Getter_returns_font_name_for_the_Latin_characters_of_body()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var bodyFontName = slideMaster.Theme.FontScheme.BodyLatinFont;

        // Assert
        bodyFontName.Should().Be("Times New Roman");
    }

    [Test]
    public void Theme_FontScheme_BodyEastAsianFont_Getter_returns_font_for_the_EastAsian_characters_of_body()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var bodyFontName = slideMaster.Theme.FontScheme.BodyEastAsianFont;

        // Assert
        bodyFontName.Should().Be("MingLiU-ExtB");
    }

    [Test]
    public void Theme_FontScheme_BodyLatinFont_Setter_font_name_for_the_Latin_characters_of_body()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.BodyLatinFont = "Arial";

        // Assert
        slideMaster.Theme.FontScheme.BodyLatinFont.Should().Be("Arial");
    }

    [Test]
    public void Theme_FontScheme_BodyEastAsianFont_Setter_font_for_the_East_Asian_characters_of_body()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        slideMaster.Theme.FontScheme.BodyEastAsianFont = "SimSun";

        // Assert
        slideMaster.Theme.FontScheme.BodyEastAsianFont.Should().Be("SimSun");
    }

    [Test]
    public void Theme_ColorScheme_color_getter_returns_scheme_with_colors()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
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

    [Test]
    public void Theme_ColorScheme_color_setter_sets_scheme_color()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
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