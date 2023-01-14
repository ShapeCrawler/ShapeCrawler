using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

namespace ShapeCrawler.UnitTests;

public class ColorFormatTests : ShapeCrawlerTest
{
    [Theory]
    [MemberData(nameof(TestCasesSetColorHex))]
    public void SetColorHex_updates_font_color(TestElementQuery colorFormatQuery)
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = colorFormatQuery.Presentation;
        var colorFormat = colorFormatQuery.GetTestColorFormat();

        // Act
        colorFormat.SetColorByHex("#008000");

        // Assert
        colorFormat.ColorHex.Should().Be("008000");

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        colorFormatQuery.Presentation = pres;
        colorFormat = colorFormatQuery.GetTestColorFormat();
        colorFormat.ColorHex.Should().Be("008000");
    }

    public static TheoryData<TestElementQuery> TestCasesSetColorHex
    {
        get
        {
            var testCases = new TheoryData<TestElementQuery>();

            testCases.Add(new TestElementQuery
            {
                Presentation = SCPresentation.Open(GetTestStream("autoshape-case001.pptx")),
                Location = Location.SlideMaster,
                SlideMasterNumber = 1,
                ShapeName = "AutoShape 1",
                ParagraphNumber = 1,
                PortionNumber = 1
            });

            var pptxStream = GetTestStream("020.pptx");
            var portionQuery = new TestElementQuery
            {
                Presentation = SCPresentation.Open(pptxStream),
                Location = Location.Slide,
                SlideIndex = 0,
                ShapeName = "TextBox 1",
                ParagraphIndex = 0,
                PortionIndex = 0
            };
            testCases.Add(portionQuery);

            var pptx = GetTestStream("001.pptx");
            portionQuery = new TestElementQuery
            {
                Presentation = SCPresentation.Open(pptx),
                Location = Location.Slide,
                SlideIndex = 0,
                ShapeId = 3,
                ParagraphIndex = 0,
                PortionIndex = 0
            };
            testCases.Add(portionQuery);

            pptx = GetTestStream("001.pptx");
            portionQuery = new TestElementQuery
            {
                Presentation = SCPresentation.Open(pptx),
                Location = Location.Slide,
                SlideIndex = 2,
                ShapeId = 4,
                ParagraphIndex = 0,
                PortionIndex = 0
            };
            testCases.Add(portionQuery);

            pptx = GetTestStream("001.pptx");
            portionQuery = new TestElementQuery
            {
                Presentation = SCPresentation.Open(pptx),
                Location = Location.Slide,
                SlideIndex = 4,
                ShapeId = 5,
                ParagraphIndex = 0,
                PortionIndex = 0
            };
            testCases.Add(portionQuery);

            return testCases;
        }
    }

    [Theory]
    [MemberData(nameof(TestCasesColorGetter))]
    public void ColorHex_Getter_returns_color(TestCase<IParagraph, string> testCase)
    {
        // Arrange
        var paragraph = testCase.Param1;
        var colorHex = testCase.Param2;
        var colorFormat = paragraph.Portions[0].Font.ColorFormat;
        var expectedColor = colorHex;

        // Act
        var actualColor = colorFormat.ColorHex;

        // Assert
        actualColor.Should().Be(expectedColor);
    }

    public static IEnumerable<object[]> TestCasesColorGetter
    {
        get
        {
            var stream1 = GetTestStream("020.pptx");
            var pres1 = SCPresentation.Open(stream1);
            var paragraph1 = pres1.Slides[0].Shapes.GetById<IAutoShape>(2).TextFrame!.Paragraphs[0];
            var testCase1 = new TestCase<IParagraph, string>(1, paragraph1, "000000");
            yield return new object[] { testCase1 };

            var stream2 = GetTestStream("020.pptx");
            var pres2 = SCPresentation.Open(stream2);
            var paragraph2 = pres2.Slides[0].Shapes.GetById<IAutoShape>(3).TextFrame!.Paragraphs[0];
            var testCase2 = new TestCase<IParagraph, string>(2, paragraph2, "000000");
            yield return new object[] { testCase2 };

            var stream3 = GetTestStream("020.pptx");
            var pres3 = SCPresentation.Open(stream3);
            var paragraph3 = pres3.Slides[2].Shapes.GetById<IAutoShape>(8).TextFrame!.Paragraphs[1];
            var testCase3 = new TestCase<IParagraph, string>(3, paragraph3, "FFFF00");
            yield return new object[] { testCase3 };

            var stream4 = GetTestStream("001.pptx");
            var pres4 = SCPresentation.Open(stream4);
            var paragraph4 = pres4.Slides[0].Shapes.GetById<IAutoShape>(4).TextFrame!.Paragraphs[0];
            var testCase4 = new TestCase<IParagraph, string>(4, paragraph4, "000000");
            yield return new object[] { testCase4 };

            var stream5 = GetTestStream("002.pptx");
            var pres5 = SCPresentation.Open(stream5);
            var paragraph5 = pres5.Slides[1].Shapes.GetById<IAutoShape>(3).TextFrame!.Paragraphs[0];
            var testCase5 = new TestCase<IParagraph, string>(5, paragraph5, "000000");
            yield return new object[] { testCase5 };

            var stream6 = GetTestStream("026.pptx");
            var pres6 = SCPresentation.Open(stream6);
            var paragraph6 = pres6.Slides[0].Shapes.GetById<IAutoShape>(128).TextFrame!.Paragraphs[0];
            var testCase6 = new TestCase<IParagraph, string>(6, paragraph6, "000000");
            yield return new object[] { testCase6 };

            var stream7 = GetTestStream("030.pptx");
            var pres7 = SCPresentation.Open(stream7);
            var paragraph7 = pres7.Slides[0].Shapes.GetById<IAutoShape>(5).TextFrame!.Paragraphs[0];
            var testCase7 = new TestCase<IParagraph, string>(7, paragraph7, "000000");
            yield return new object[] { testCase7 };

            var stream8 = GetTestStream("031.pptx");
            var pres8 = SCPresentation.Open(stream8);
            var paragraph8 = pres8.Slides[0].Shapes.GetById<IAutoShape>(44).TextFrame!.Paragraphs[0];
            var testCase8 = new TestCase<IParagraph, string>(8, paragraph8, "000000");
            yield return new object[] { testCase8 };

            var stream9 = GetTestStream("033.pptx");
            var pres9 = SCPresentation.Open(stream9);
            var paragraph9 = pres9.Slides[0].Shapes.GetById<IAutoShape>(3).TextFrame!.Paragraphs[0];
            var testCase9 = new TestCase<IParagraph, string>(9, paragraph9, "000000");
            yield return new object[] { testCase9 };

            var stream10 = GetTestStream("038.pptx");
            var pres10 = SCPresentation.Open(stream10);
            var paragraph10 = pres10.Slides[0].Shapes.GetById<IAutoShape>(102).TextFrame!.Paragraphs[0];
            var testCase10 = new TestCase<IParagraph, string>(10, paragraph10, "000000");
            yield return new object[] { testCase10 };

            var stream11 = GetTestStream("001.pptx");
            var pres11 = SCPresentation.Open(stream11);
            var paragraph11 = pres11.Slides[2].Shapes.GetById<IAutoShape>(4).TextFrame!.Paragraphs[0];
            var testCase11 = new TestCase<IParagraph, string>(11, paragraph11, "000000");
            yield return new object[] { testCase11 };

            var stream12 = GetTestStream("001.pptx");
            var pres12 = SCPresentation.Open(stream12);
            var paragraph12 = pres12.Slides[4].Shapes.GetById<IAutoShape>(5).TextFrame!.Paragraphs[0];
            var testCase12 = new TestCase<IParagraph, string>(12, paragraph12, "000000");
            yield return new object[] { testCase12 };

            var stream13 = GetTestStream("014.pptx");
            var pres13 = SCPresentation.Open(stream13);
            var paragraph13 = pres13.Slides[0].Shapes.GetById<IAutoShape>(61).TextFrame!.Paragraphs[0];
            var testCase13 = new TestCase<IParagraph, string>(13, paragraph13, "595959");
            yield return new object[] { testCase13 };

            var stream14 = GetTestStream("014.pptx");
            var pres14 = SCPresentation.Open(stream14);
            var paragraph14 = pres14.Slides[5].Shapes.GetById<IAutoShape>(52).TextFrame!.Paragraphs[0];
            var testCase14 = new TestCase<IParagraph, string>(14, paragraph14, "FFFFFF");
            yield return new object[] { testCase14 };
            
            var stream15 = GetTestStream("032.pptx");
            var pres15 = SCPresentation.Open(stream15);
            var paragraph15 = pres15.Slides[0].Shapes.GetById<IAutoShape>(10242).TextFrame!.Paragraphs[0];
            var testCase15 = new TestCase<IParagraph, string>(15, paragraph15, "0070C0");
            yield return new object[] { testCase15 };
            
            var stream16 = GetTestStream("034.pptx");
            var pres16 = SCPresentation.Open(stream16);
            var paragraph16 = pres16.Slides[0].Shapes.GetById<IAutoShape>(2).TextFrame!.Paragraphs[0];
            var testCase16 = new TestCase<IParagraph, string>(16, paragraph16, "000000");
            yield return new object[] { testCase16 };
            
            var stream17 = GetTestStream("035.pptx");
            var pres17 = SCPresentation.Open(stream17);
            var paragraph17 = pres17.Slides[0].Shapes.GetById<IAutoShape>(9).TextFrame!.Paragraphs[0];
            var testCase17 = new TestCase<IParagraph, string>(17, paragraph17, "000000");
            yield return new object[] { testCase17 };
            
            var stream18 = GetTestStream("036.pptx");
            var pres18 = SCPresentation.Open(stream18);
            var paragraph18 = pres18.Slides[0].Shapes.GetById<IAutoShape>(6146).TextFrame!.Paragraphs[0];
            var testCase18 = new TestCase<IParagraph, string>(18, paragraph18, "404040");
            yield return new object[] { testCase18 };
            
            var stream19 = GetTestStream("037.pptx");
            var pres19 = SCPresentation.Open(stream19);
            var paragraph19 = pres19.Slides[0].Shapes.GetById<IAutoShape>(7).TextFrame!.Paragraphs[0];
            var testCase19 = new TestCase<IParagraph, string>(19, paragraph19, "1A1A1A");
            yield return new object[] { testCase19 };
        }
    }

    [Fact]
    public void ColorHex_Getter_returns_White_color()
    {
        // Arrange
        var pres20 = SCPresentation.Open(GetTestStream("020.pptx"));
        var shape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);
        var colorFormat = shape.TextFrame!.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FFFFFF");
    }

    [Fact]
    public void ColorHex_Getter_returns_color_of_SlideLayout_Placeholder()
    {
        // Arrange
        var pres1 = SCPresentation.Open(GetTestStream("001.pptx"));
        IAutoShape titlePh = (IAutoShape)SCPresentation.Open(GetTestStream("001.pptx")).Slides[0].SlideLayout.Shapes.First(sp => sp.Id == 2);
        IColorFormat colorFormat = titlePh.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("000000");
    }

    [Fact]
    public void ColorHex_Getter_returns_color_of_SlideMaster_Non_Placeholder()
    {
        // Arrange
        IAutoShape nonPlaceholder = (IAutoShape)SCPresentation.Open(GetTestStream("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 8);
        IColorFormat colorFormat = nonPlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FFFFFF");
    }

    [Fact]
    public void ColorHex_Getter_returns_color_of_Title_SlideMaster_Placeholder()
    {
        // Arrange
        IAutoShape titlePlaceholder = (IAutoShape)SCPresentation.Open(GetTestStream("001.pptx")).SlideMasters[0].Shapes.First(sp => sp.Id == 2);
        IColorFormat colorFormat = titlePlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("000000");
    }

    [Fact]
    public void ColorHex_Getter_returns_color_of_Table_Cell_on_Slide()
    {
        // Arrange
        var pres1 = SCPresentation.Open(GetTestStream("001.pptx"));
        ITable table = (ITable)SCPresentation.Open(GetTestStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);
        IColorFormat colorFormat = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act-Assert
        colorFormat.ColorHex.Should().Be("FF0000");
    }

    [Fact]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsTakenFromThemeScheme()
    {
        // Arrange
        var pres20 = SCPresentation.Open(GetTestStream("020.pptx"));
        IAutoShape nonPhAutoShape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        IColorFormat colorFormat = nonPhAutoShape.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act
        SCColorType colorType = colorFormat.ColorType;

        // Assert
        colorType.Should().Be(SCColorType.Scheme);
    }

    [Fact]
    public void ColorType_ReturnsSchemeColorType_WhenFontColorIsSetAsRGB()
    {
        // Arrange
        IAutoShape placeholder = (IAutoShape)SCPresentation.Open(GetTestStream("014.pptx")).Slides[5].Shapes.First(sp => sp.Id == 52);
        IColorFormat colorFormat = placeholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

        // Act
        SCColorType colorType = colorFormat.ColorType;

        // Assert
        colorType.Should().Be(SCColorType.RGB);
    }
}