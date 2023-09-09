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
        colorFormat.Hex.Should().Be("FFFFFF");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_SlideLayout_Placeholder()
    {
        // Arrange
        var pres = new SCPresentation(StreamOf("001.pptx"));
        var titlePlaceholder = pres.Slides[0].SlideLayout.Shapes.GetById<IShape>(2);
        var fontColor = titlePlaceholder.TextFrame.Paragraphs[0].Portions[0].Font!.Color;

        // Act-Assert
        fontColor.Hex.Should().Be("000000");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_SlideMaster_Non_Placeholder()
    {
        // Arrange
        IShape nonPlaceholder = (IShape)new SCPresentation(StreamOf("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 8);
        IFontColor colorFormat = nonPlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.Hex.Should().Be("FFFFFF");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_Title_SlideMaster_Placeholder()
    {
        // Arrange
        IShape titlePlaceholder = (IShape)new SCPresentation(StreamOf("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 2);
        IFontColor colorFormat = titlePlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        colorFormat.Hex.Should().Be("000000");
    }

    [Test]
    public void ColorHex_Getter_returns_color_of_Table_Cell_on_Slide()
    {
        // Arrange
        var pres = new SCPresentation(StreamOf("001.pptx"));
        var table = pres.Slides[1].Shapes.GetById<ITable>(4);
        var fontColor = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act-Assert
        fontColor.Hex.Should().Be("FF0000");
    }

    [Test]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsTakenFromThemeScheme()
    {
        // Arrange
        var pres = new SCPresentation(StreamOf("020.pptx"));
        var nonPhAutoShape = pres.Slides[0].Shapes.GetById<IShape>(2);
        var fontColor = nonPhAutoShape.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act
        var colorType = fontColor.Type;

        // Assert
        colorType.Should().Be(SCColorType.Theme);
    }

    [Test]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsSetAsRGB()
    {
        // Arrange
        var pres = new SCPresentation(StreamOf("014.pptx"));
        var placeholder = pres.Slides[5].Shapes.GetById<IShape>(52);
        var fontColor = placeholder.TextFrame.Paragraphs[0].Portions[0].Font.Color;

        // Act
        var colorType = fontColor.Type;

        // Assert
        colorType.Should().Be(SCColorType.RGB);
    }

    [Test]
    [SlideQueryPortion("020.pptx", 1, "TextBox 1", 1,  1)]
    [SlideQueryPortion("001.pptx", 1, 3, 1,  1)]
    [SlideQueryPortion("001.pptx", 3, 4, 1,  1)]
    [SlideQueryPortion("001.pptx", 5, 5, 1,  1)]
    public void SetColorHex_updates_font_color(IPresentation pres, TestPortionQuery portionQuery)
    {
        // Arrange
        var mStream = new MemoryStream();
        var color = portionQuery.Get(pres).Font!.Color;

        // Act
        color.Update("#008000");

        // Assert
        color.Hex.Should().Be("008000");

        pres.SaveAs(mStream);
        pres = new SCPresentation(mStream);
        color = portionQuery.Get(pres).Font!.Color;
        color.Hex.Should().Be("008000");
    }
    
    [Test]
    [MasterPortion("autoshape-case001.pptx", "AutoShape 1", 1,  1)]
    public void SetColorHex_updates_font_color_of_master(IPresentation pres, TestPortionQuery portionQuery)
    {
        // Arrange
        var mStream = new MemoryStream();
        var color = portionQuery.Get(pres).Font!.Color;

        // Act
        color.Update("#008000");

        // Assert
        color.Hex.Should().Be("008000");

        pres.SaveAs(mStream);
        pres = new SCPresentation(mStream);
        color = portionQuery.Get(pres).Font!.Color;
        color.Hex.Should().Be("008000");
    }
    
    [Test]
    [SlidePortion("Test Case #1", "020.pptx", slide: 1, shapeId: 2, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #2", "020.pptx", slide: 1, shapeId: 3, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #3", "020.pptx", slide: 3, shapeId: 8, paragraph: 2, portion: 1, expectedResult: "FFFF00")]
    [SlidePortion("Test Case #4", "001.pptx", slide: 1, shapeId: 4, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #5", "002.pptx", slide: 2, shapeId: 3, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #6", "026.pptx", slide: 1, shapeId: 128, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #7", "autoshape-case017_slide-number.pptx", slide: 1, shapeId: 5, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #8", "031.pptx", slide: 1, shapeId: 44, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #9", "033.pptx", slide: 1, shapeId: 3, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #10", "038.pptx", slide: 1, shapeId: 102, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #11", "001.pptx", slide: 3, shapeId: 4, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #12", "001.pptx", slide: 5, shapeId: 5, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #13", "034.pptx", slide: 1, shapeId: 2, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #14", "035.pptx", slide: 1, shapeId: 9, paragraph: 1, portion: 1, expectedResult: "000000")]
    [SlidePortion("Test Case #15", "036.pptx", slide: 1, shapeId: 6146, paragraph: 1, portion: 1, expectedResult: "404040")]
    [SlidePortion("Test Case #16", "037.pptx", slide: 1, shapeId: 7, paragraph: 1, portion: 1, expectedResult: "1A1A1A")]
    [SlidePortion("Test Case #17", "014.pptx", slide: 1, shapeId: 61, paragraph: 1, portion: 1, expectedResult: "595959")]
    [SlidePortion("Test Case #18", "014.pptx", slide: 6, shapeId: 52, paragraph: 1, portion: 1, expectedResult: "FFFFFF")]
    [SlidePortion("Test Case #19", "032.pptx", slide: 1, shapeId: 10242, paragraph: 1, portion: 1, expectedResult: "0070C0")]
    public void ColorHex_Getter_returns_color_hex(IParagraphPortion portion, string expectedColorHex)
    {
        // Arrange
        var fontColor = portion.Font!.Color;

        // Act
        var colorHex = fontColor.Hex;

        // Assert
        colorHex.Should().Be(expectedColorHex);
    }
}