using DocumentFormat.OpenXml.Vml.Office;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class FontColorTests : SCTest
{
    [Test]
    public void ColorHex_Getter_returns_White_color()
    {
        // Arrange
        var shape = (IShape)new SCPresentation(StreamOf("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);
        var colorFormat = shape.TextFrame!.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FFFFFF");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_SlideLayout_Placeholder()
    {
        // Arrange
        var titlePh = (IShape)new SCPresentation(StreamOf("001.pptx")).Slides[0].SlideLayout.Shapes.First(sp => sp.Id == 2);
        var colorFormat = titlePh.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("000000");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_SlideMaster_Non_Placeholder()
    {
        // Arrange
        IShape nonPlaceholder = (IShape)new SCPresentation(StreamOf("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 8);
        IFontColor colorFormat = nonPlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FFFFFF");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_Title_SlideMaster_Placeholder()
    {
        // Arrange
        IShape titlePlaceholder = (IShape)new SCPresentation(StreamOf("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 2);
        IFontColor colorFormat = titlePlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("000000");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_Table_Cell_on_Slide()
    {
        // Arrange
        var table = (ITable)new SCPresentation(StreamOf("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);
        var colorFormat = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FF0000");
    }

    [Test]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsTakenFromThemeScheme()
    {
        // Arrange
        var nonPhAutoShape = (IShape)new SCPresentation(StreamOf("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        var colorFormat = nonPhAutoShape.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act
        SCColorType colorType = colorFormat.ColorType;

        // Assert
        colorType.Should().Be(SCColorType.Theme);
    }

    [Test]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsSetAsRGB()
    {
        // Arrange
        IShape placeholder = (IShape)new SCPresentation(StreamOf("014.pptx")).Slides[5].Shapes.First(sp => sp.Id == 52);
        IFontColor colorFormat = placeholder.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act
        SCColorType colorType = colorFormat.ColorType;

        // Assert
        colorType.Should().Be(SCColorType.RGB);
    }

    [Test]
    [MasterPortion("autoshape-case001.pptx", "AutoShape 1", 1,  1)]
    [SlidePortion("020.pptx", 1, "TextBox 1", 1,  1)]
    [SlidePortion("001.pptx", 1, 3, 1,  1)]
    [SlidePortion("001.pptx", 3, 4, 1,  1)]
    [SlidePortion("001.pptx", 5, 5, 1,  1)]
    public void SetColorHex_updates_font_color(IPresentation pres, TestPortionQuery portionQuery)
    {
        // Arrange
        var mStream = new MemoryStream();
        var color = portionQuery.Get(pres).Font!.Color;

        // Act
        color.SetColorByHex("#008000");

        // Assert
        color.ColorHex.Should().Be("008000");

        pres.SaveAs(mStream);
        pres = new SCPresentation(mStream);
        color = portionQuery.Get(pres).Font!.Color;
        color.ColorHex.Should().Be("008000");
    }
}