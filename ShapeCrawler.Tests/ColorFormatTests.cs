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
            Color expectedColor = ColorTranslator.FromHtml("#008000");
            var mStream = new MemoryStream();
            var pres = colorFormatQuery.Presentation;
            var colorFormat = colorFormatQuery.GetTestColorFormat();

            // Act
            colorFormat.SetColorHex("#008000");

            // Assert
            colorFormat.Color.Should().Be(expectedColor);

            pres.SaveAs(@"c:\temp\result.pptx");
            pres.SaveAs(mStream);
            pres = SCPresentation.Open(mStream);
            colorFormatQuery.Presentation = pres;
            colorFormat = colorFormatQuery.GetTestColorFormat();
            colorFormat.Color.Should().Be(expectedColor);
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

#if DEBUG
        [Fact]
        public void Color_GetterReturnsColor_OfNonPlaceholder()
        {
            // Arrange
            IAutoShape nonPhAutoShapeCase1 = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 2);
            IAutoShape nonPhAutoShapeCase2 = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3);
            IAutoShape nonPhAutoShapeCase3 = (IAutoShape)_fixture.Pre020.Slides[2].Shapes.First(sp => sp.Id == 8);
            IAutoShape nonPhAutoShapeCase4 = (IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 4);
            IAutoShape nonPhAutoShapeCase5 = (IAutoShape)_fixture.Pre002.Slides[1].Shapes.First(sp => sp.Id == 3);
            IAutoShape nonPhAutoShapeCase6 = (IAutoShape)_fixture.Pre026.Slides[0].Shapes.First(sp => sp.Id == 128);
            IAutoShape nonPhAutoShapeCase7 = (IAutoShape)_fixture.Pre030.Slides[0].Shapes.First(sp => sp.Id == 5);
            IAutoShape nonPhAutoShapeCase8 = (IAutoShape)_fixture.Pre031.Slides[0].Shapes.First(sp => sp.Id == 44);
            IAutoShape nonPhAutoShapeCase9 = (IAutoShape)_fixture.Pre033.Slides[0].Shapes.First(sp => sp.Id == 3);
            IAutoShape nonPhAutoShapeCase10 = (IAutoShape)_fixture.Pre038.Slides[0].Shapes.First(sp => sp.Id == 102);
            IColorFormat colorFormatC1 = nonPhAutoShapeCase1.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC2 = nonPhAutoShapeCase2.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC3 = nonPhAutoShapeCase3.TextBox.Paragraphs[1].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC4 = nonPhAutoShapeCase4.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC5 = nonPhAutoShapeCase5.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC6 = nonPhAutoShapeCase6.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC7 = nonPhAutoShapeCase7.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC8 = nonPhAutoShapeCase8.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC9 = nonPhAutoShapeCase9.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC10 = nonPhAutoShapeCase10.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormatC1.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC2.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC3.Color.Should().Be(ColorTranslator.FromHtml("#FFFF00"));
            colorFormatC4.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC5.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC6.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC7.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC8.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC9.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC10.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
        }

        [Fact]
        public void Color_GetterReturnsWhiteColor_WhenFontHasPredefinedWhiteColor()
        {
            // Arrange
            IAutoShape nonPhAutoShapeCase = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 4);
            IColorFormat colorFormat = nonPhAutoShapeCase.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.Color.Should().Be(Color.White);
        }

        [Fact]
        public void Color_GetterReturnsColor_OfSlidePlaceholder()
        {
            // Arrange
            IAutoShape placeholderCase1 = (IAutoShape)_fixture.Pre001.Slides[2].Shapes.First(sp => sp.Id == 4);
            IAutoShape placeholderCase2 = (IAutoShape)_fixture.Pre001.Slides[4].Shapes.First(sp => sp.Id == 5);
            IAutoShape placeholderCase3 = (IAutoShape)_fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61);
            IAutoShape placeholderCase4 = (IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52);
            IAutoShape placeholderCase5 = (IAutoShape)_fixture.Pre032.Slides[0].Shapes.First(sp => sp.Id == 10242);
            IAutoShape titlePhCase6 = (IAutoShape)_fixture.Pre034.Slides[0].Shapes.First(sp => sp.Id == 2);
            IAutoShape titlePhCase7 = (IAutoShape)_fixture.Pre035.Slides[0].Shapes.First(sp => sp.Id == 9);
            IAutoShape bodyPhCase8 = (IAutoShape)_fixture.Pre036.Slides[0].Shapes.First(sp => sp.Id == 6146);
            IAutoShape bodyPhCase9 = (IAutoShape)_fixture.Pre037.Slides[0].Shapes.First(sp => sp.Id == 7);
            IColorFormat colorFormatC1 = placeholderCase1.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC2 = placeholderCase2.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC3 = placeholderCase3.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC4 = placeholderCase4.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC5 = placeholderCase5.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC6 = titlePhCase6.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC7 = titlePhCase7.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC8 = bodyPhCase8.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC9 = bodyPhCase9.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormatC1.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC2.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC3.Color.Should().Be(ColorTranslator.FromHtml("#595959"));
            colorFormatC4.Color.Should().Be(ColorTranslator.FromHtml("#FFFFFF"));
            colorFormatC5.Color.Should().Be(ColorTranslator.FromHtml("#0070C0"));
            colorFormatC6.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC7.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC8.Color.Should().Be(ColorTranslator.FromHtml("#404040"));
            colorFormatC9.Color.Should().Be(ColorTranslator.FromHtml("#1A1A1A"));
        }

        [Fact]
        public void Color_GetterReturnsColor_OfSlideLayoutPlaceholder()
        {
            // Arrange
            IAutoShape titlePh = (IAutoShape)_fixture.Pre001.Slides[0].SlideLayout.Shapes.First(sp => sp.Id == 2);
            IColorFormat colorFormat = titlePh.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
        }

        [Fact]
        public void Color_GetterReturnsColor_OfSlideMasterNonPlaceholder()
        {
            // Arrange
            Color whiteColor = ColorTranslator.FromHtml("#FFFFFF");
            IAutoShape nonPlaceholder = (IAutoShape)_fixture.Pre001.SlideMasters[0].Shapes.First(sp => sp.Id == 8);
            IColorFormat colorFormat = nonPlaceholder.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.Color.Should().Be(whiteColor);
        }

        [Fact]
        public void Color_GetterReturnsColor_OfTitlePlaceholderOnSlideMaster()
        {
            // Arrange
            Color blackColor = ColorTranslator.FromHtml("#000000");
            IAutoShape titlePlaceholder = (IAutoShape)_fixture.Pre001.SlideMasters[0].Shapes.First(sp => sp.Id == 2);
            IColorFormat colorFormat = titlePlaceholder.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.Color.Should().Be(blackColor);
        }

        [Fact]
        public void Color_GetterReturnsColor_OfTableCellOnSlide()
        {
            // Arrange
            Color redColor = ColorTranslator.FromHtml("#FF0000");
            ITable table = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4);
            IColorFormat colorFormat = table.Rows[0].Cells[0].TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.Color.Should().Be(redColor);
        }

        [Fact]
        public void ColorType_ReturnsSchemeColorType_WhenFontColorIsTakenFromThemeScheme()
        {
            // Arrange
            IAutoShape nonPhAutoShape = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 2);
            IColorFormat colorFormat = nonPhAutoShape.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

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
            IColorFormat colorFormat = placeholder.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act
            SCColorType colorType = colorFormat.ColorType;

            // Assert
            colorType.Should().Be(SCColorType.RGB);
        }
#endif
    }
}