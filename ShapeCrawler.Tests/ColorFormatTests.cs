using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
using Xunit;

namespace ShapeCrawler.Tests
{
    public class ColorFormatTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ColorFormatTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

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
                
                portionQuery = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(Resources._020),
                    Location = Location.Slide,
                    SlideIndex = 0,
                    ShapeId = 3,
                    ParagraphIndex = 0,
                    PortionIndex = 0
                };
                testCases.Add(portionQuery);
                
                portionQuery = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(Resources._001),
                    Location = Location.Slide,
                    SlideIndex = 2,
                    ShapeId = 4,
                    ParagraphIndex = 0,
                    PortionIndex = 0
                };
                testCases.Add(portionQuery);
                
                portionQuery = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(Resources._001),
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
        public void Color_Getter_returns_color(TestCase<IParagraph, string> testCase)
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
                var autoShape1 = pres1.Slides[0].Shapes.GetById<IAutoShape>(2);
                var paragraph1 = autoShape1.TextFrame.Paragraphs[0];
                var testCase1 = new TestCase<IParagraph, string>(1, paragraph1, "000000");
                yield return new object[] { testCase1 };
                
                var stream2 = GetTestStream("020.pptx");
                var pres2 = SCPresentation.Open(stream2);
                var autoShape2 = pres2.Slides[0].Shapes.GetById<IAutoShape>(3);
                var paragraph2 = autoShape2.TextFrame.Paragraphs[0];
                var testCase2 = new TestCase<IParagraph, string>(2, paragraph2, "000000");
                yield return new object[] { testCase2 };
                
                var stream3 = GetTestStream("020.pptx");
                var pres3 = SCPresentation.Open(stream3);
                var autoShape3 = pres3.Slides[2].Shapes.GetById<IAutoShape>(8);
                var paragraph3 = autoShape3.TextFrame.Paragraphs[1];
                var testCase3 = new TestCase<IParagraph, string>(3, paragraph3, "FFFF00");
                yield return new object[] { testCase3 };
                
                var stream4 = GetTestStream("001.pptx");
                var pres4 = SCPresentation.Open(stream4);
                var autoShape4 = pres4.Slides[0].Shapes.GetById<IAutoShape>(4);
                var paragraph4 = autoShape4.TextFrame.Paragraphs[0];
                var testCase4 = new TestCase<IParagraph, string>(4, paragraph4, "000000");
                yield return new object[] { testCase4 };
                
                var stream5 = GetTestStream("002.pptx");
                var pres5 = SCPresentation.Open(stream5);
                var autoShape5 = pres5.Slides[1].Shapes.GetById<IAutoShape>(3);
                var paragraph5 = autoShape5.TextFrame.Paragraphs[0];
                var testCase5 = new TestCase<IParagraph, string>(5, paragraph5, "000000");
                yield return new object[] { testCase5 };
                
                var stream6 = GetTestStream("026.pptx");
                var pres6 = SCPresentation.Open(stream6);
                var autoShape6 = pres6.Slides[0].Shapes.GetById<IAutoShape>(128);
                var paragraph6 = autoShape6.TextFrame.Paragraphs[0];
                var testCase6 = new TestCase<IParagraph, string>(6, paragraph6, "000000");
                yield return new object[] { testCase6 };
                
                var stream7 = GetTestStream("030.pptx");
                var pres7 = SCPresentation.Open(stream7);
                var autoShape7 = pres7.Slides[0].Shapes.GetById<IAutoShape>(5);
                var paragraph7 = autoShape7.TextFrame.Paragraphs[0];
                var testCase7 = new TestCase<IParagraph, string>(7, paragraph7, "000000");
                yield return new object[] { testCase7 };
                
                var stream8 = GetTestStream("031.pptx");
                var pres8 = SCPresentation.Open(stream8);
                var autoShape8 = pres8.Slides[0].Shapes.GetById<IAutoShape>(44);
                var paragraph8 = autoShape8.TextFrame.Paragraphs[0];
                var testCase8 = new TestCase<IParagraph, string>(8, paragraph8, "000000");
                yield return new object[] { testCase8 };
                
                var stream9 = GetTestStream("033.pptx");
                var pres9 = SCPresentation.Open(stream9);
                var autoShape9 = pres9.Slides[0].Shapes.GetById<IAutoShape>(3);
                var paragraph9 = autoShape9.TextFrame.Paragraphs[0];
                var testCase9 = new TestCase<IParagraph, string>(9, paragraph9, "000000");
                yield return new object[] { testCase9 };
                
                var stream10 = GetTestStream("038.pptx");
                var pres10 = SCPresentation.Open(stream10);
                var autoShape10 = pres10.Slides[0].Shapes.GetById<IAutoShape>(102);
                var paragraph10 = autoShape10.TextFrame.Paragraphs[0];
                var testCase10 = new TestCase<IParagraph, string>(10, paragraph10, "000000");
                yield return new object[] { testCase10 };
            }
        }

        [Fact]
        public void Color_Getter_returns_White_color()
        {
            // Arrange
            var shape = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 4);
            var colorFormat = shape.TextFrame!.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.ColorHex.Should().Be("FFFFFF");
        }

        [Fact]
        public void Color_Getter_returns_color_of_Slide_Placeholder()
        {
            // Arrange
            var placeholderCase1 = (IAutoShape)_fixture.Pre001.Slides[2].Shapes.First(sp => sp.Id == 4);
            var placeholderCase2 = (IAutoShape)_fixture.Pre001.Slides[4].Shapes.First(sp => sp.Id == 5);
            var placeholderCase3 = (IAutoShape)_fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61);
            var placeholderCase4 = (IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52);
            var placeholderCase5 = (IAutoShape)_fixture.Pre032.Slides[0].Shapes.First(sp => sp.Id == 10242);
            var titlePhCase6 = (IAutoShape)_fixture.Pre034.Slides[0].Shapes.First(sp => sp.Id == 2);
            var titlePhCase7 = (IAutoShape)_fixture.Pre035.Slides[0].Shapes.First(sp => sp.Id == 9);
            var bodyPhCase8 = (IAutoShape)_fixture.Pre036.Slides[0].Shapes.First(sp => sp.Id == 6146);
            var bodyPhCase9 = (IAutoShape)_fixture.Pre037.Slides[0].Shapes.First(sp => sp.Id == 7);
            var colorFormatC1 = placeholderCase1.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC2 = placeholderCase2.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC3 = placeholderCase3.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC4 = placeholderCase4.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC5 = placeholderCase5.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC6 = titlePhCase6.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC7 = titlePhCase7.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC8 = bodyPhCase8.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;
            var colorFormatC9 = bodyPhCase9.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormatC1.ColorHex.Should().Be("000000");
            colorFormatC2.ColorHex.Should().Be("000000");
            colorFormatC3.ColorHex.Should().Be("595959");
            colorFormatC4.ColorHex.Should().Be("FFFFFF");
            colorFormatC5.ColorHex.Should().Be("0070C0");
            colorFormatC6.ColorHex.Should().Be("000000");
            colorFormatC7.ColorHex.Should().Be("000000");
            colorFormatC8.ColorHex.Should().Be("404040");
            colorFormatC9.ColorHex.Should().Be("1A1A1A");
        }

        [Fact]
        public void Color_Getter_returns_color_of_SlideLayout_Placeholder()
        {
            // Arrange
            IAutoShape titlePh = (IAutoShape)_fixture.Pre001.Slides[0].SlideLayout.Shapes.First(sp => sp.Id == 2);
            IColorFormat colorFormat = titlePh.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.ColorHex.Should().Be("000000");
        }

        [Fact]
        public void Color_Getter_returns_color_of_SlideMaster_Non_Placeholder()
        {
            // Arrange
            IAutoShape nonPlaceholder = (IAutoShape)_fixture.Pre001.SlideMasters[0].Shapes.First(sp => sp.Id == 8);
            IColorFormat colorFormat = nonPlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.ColorHex.Should().Be("FFFFFF");
        }

        [Fact]
        public void Color_Getter_returns_color_of_Title_SlideMaster_Placeholder()
        {
            // Arrange
            IAutoShape titlePlaceholder = (IAutoShape)_fixture.Pre001.SlideMasters[0].Shapes.First(sp => sp.Id == 2);
            IColorFormat colorFormat = titlePlaceholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.ColorHex.Should().Be("000000");
        }

        [Fact]
        public void Color_Getter_returns_color_of_Table_Cell_on_Slide()
        {
            // Arrange
            ITable table = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4);
            IColorFormat colorFormat = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.ColorHex.Should().Be("FF0000");
        }

        [Fact]
        public void ColorType_ReturnsSchemeColorType_WhenFontColorIsTakenFromThemeScheme()
        {
            // Arrange
            IAutoShape nonPhAutoShape = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 2);
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
            IAutoShape placeholder = (IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52);
            IColorFormat colorFormat = placeholder.TextFrame.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act
            SCColorType colorType = colorFormat.ColorType;

            // Assert
            colorType.Should().Be(SCColorType.RGB);
        }
    }
}