using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class ColorFormatTests : SCTest
{
    [Test]
    public void ColorHex_Getter_returns_White_color()
    {
        // Arrange
        var shape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);
        var colorFormat = shape.TextFrame!.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FFFFFF");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_SlideLayout_Placeholder()
    {
        // Arrange
        var titlePh = (IAutoShape)SCPresentation.Open(GetInputStream("001.pptx")).Slides[0].SlideLayout.Shapes.First(sp => sp.Id == 2);
        var colorFormat = titlePh.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("000000");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_SlideMaster_Non_Placeholder()
    {
        // Arrange
        IAutoShape nonPlaceholder = (IAutoShape)SCPresentation.Open(GetInputStream("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 8);
        IColorFormat colorFormat = nonPlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FFFFFF");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_Title_SlideMaster_Placeholder()
    {
        // Arrange
        IAutoShape titlePlaceholder = (IAutoShape)SCPresentation.Open(GetInputStream("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 2);
        IColorFormat colorFormat = titlePlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("000000");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_Table_Cell_on_Slide()
    {
        // Arrange
        var table = (ITable)SCPresentation.Open(GetInputStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);
        var colorFormat = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FF0000");
    }

    [Test]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsTakenFromThemeScheme()
    {
        // Arrange
        var nonPhAutoShape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        var colorFormat = nonPhAutoShape.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act
        SCColorType colorType = colorFormat.ColorType;

        // Assert
        colorType.Should().Be(SCColorType.Scheme);
    }

    [Test]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsSetAsRGB()
    {
        // Arrange
        IAutoShape placeholder = (IAutoShape)SCPresentation.Open(GetInputStream("014.pptx")).Slides[5].Shapes.First(sp => sp.Id == 52);
        IColorFormat colorFormat = placeholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act
        SCColorType colorType = colorFormat.ColorType;

        // Assert
        colorType.Should().Be(SCColorType.RGB);
    }
}