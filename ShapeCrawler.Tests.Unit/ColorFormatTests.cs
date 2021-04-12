using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    public class ColorFormatTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ColorFormatTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

#if DEBUG
        [Theory(Skip = "In Progress")]
        [MemberData(nameof(TestCasesColorSetter))]
        public void Color_SetterSetsGreenColorForFont_WhenGreenIsSpecified(SCPresentation presentation, SlideElementQuery portionRequest)
        {
            // Arrange
            Color greenColor = ColorTranslator.FromHtml("#008000");
            MemoryStream mStream = new ();
            IColorFormat colorFormat = TestHelper.GetPortion(presentation, portionRequest).Font.ColorFormat;

            // Act
            colorFormat.Color = greenColor;

            // Assert
            colorFormat.Color.Should().Be(greenColor);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            colorFormat = TestHelper.GetPortion(presentation, portionRequest).Font.ColorFormat;
            colorFormat.Color.Should().Be(greenColor);
        }
#endif

        public static IEnumerable<object[]> TestCasesColorSetter()
        {
            SCPresentation presentationCase1 = SCPresentation.Open(Resources._020, true);
            SlideElementQuery portionRequestCase1 = new();
            portionRequestCase1.SlideIndex = 0;
            portionRequestCase1.ShapeId = 2;
            portionRequestCase1.ParagraphIndex = 0;
            portionRequestCase1.PortionIndex = 0;

            SCPresentation presentationCase2 = SCPresentation.Open(Resources._020, true);
            SlideElementQuery portionRequestCase2 = new();
            portionRequestCase2.SlideIndex = 0;
            portionRequestCase2.ShapeId = 3;
            portionRequestCase2.ParagraphIndex = 0;
            portionRequestCase2.PortionIndex = 0;

            SCPresentation presentationCase3 = SCPresentation.Open(Resources._001, true);
            SlideElementQuery portionRequestCase3 = new();
            portionRequestCase3.SlideIndex = 2;
            portionRequestCase3.ShapeId = 4;
            portionRequestCase3.ParagraphIndex = 0;
            portionRequestCase3.PortionIndex = 0;

            SCPresentation presentationCase4 = SCPresentation.Open(Resources._001, true);
            SlideElementQuery portionRequestCase4 = new();
            portionRequestCase4.SlideIndex = 4;
            portionRequestCase4.ShapeId = 5;
            portionRequestCase4.ParagraphIndex = 0;
            portionRequestCase4.PortionIndex = 0;

            var testCases = new List<object[]>
            {
                new object[] {presentationCase1, portionRequestCase1},
                new object[] {presentationCase2, portionRequestCase2},
                new object[] {presentationCase3, portionRequestCase3},
                new object[] {presentationCase4, portionRequestCase4}
            };

            return testCases;
        }

        [Fact]
        public void Color_GetterReturnsRGBColorInHexFormat_OfNonPlaceholder()
        {
            // Arrange
            IAutoShape nonPhAutoShapeCase1 = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 2);
            IAutoShape nonPhAutoShapeCase2 = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3);
            IAutoShape nonPhAutoShapeCase3 = (IAutoShape)_fixture.Pre020.Slides[2].Shapes.First(sp => sp.Id == 8);
            IAutoShape nonPhAutoShapeCase4 = (IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 4);
            IAutoShape nonPhAutoShapeCase5 = (IAutoShape)_fixture.Pre002.Slides[1].Shapes.First(sp => sp.Id == 3);
            IAutoShape nonPhAutoShapeCase6 = (IAutoShape)_fixture.Pre026.Slides[0].Shapes.First(sp => sp.Id == 128);
            IColorFormat colorFormatC1 = nonPhAutoShapeCase1.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC2 = nonPhAutoShapeCase2.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC3 = nonPhAutoShapeCase3.TextBox.Paragraphs[1].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC4 = nonPhAutoShapeCase4.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC5 = nonPhAutoShapeCase5.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC6 = nonPhAutoShapeCase6.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormatC1.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC2.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC3.Color.Should().Be(ColorTranslator.FromHtml("#FFFF00"));
            colorFormatC4.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC5.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC6.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
        }

        [Fact(Skip = "In Progress")]
        public void Color_GetterReturnsWhiteColor_WhenColorIsWhiteFromPredefinedCollectionOfColors()
        {
            // Arrange
            IAutoShape nonPhAutoShapeCase = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 4);
            IColorFormat colorFormat = nonPhAutoShapeCase.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormat.Color.Should().Be(Color.White);
        }

        [Fact]
        public void Color_GetterReturnsColor_OfPlaceholder()
        {
            // Arrange
            IAutoShape placeholderCase1 = (IAutoShape)_fixture.Pre001.Slides[2].Shapes.First(sp => sp.Id == 4);
            IAutoShape placeholderCase2 = (IAutoShape)_fixture.Pre001.Slides[4].Shapes.First(sp => sp.Id == 5);
            IAutoShape placeholderCase3 = (IAutoShape)_fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61);
            IAutoShape placeholderCase4 = (IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52);
            IColorFormat colorFormatC1 = placeholderCase1.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC2 = placeholderCase2.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC3 = placeholderCase3.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;
            IColorFormat colorFormatC4 = placeholderCase4.TextBox.Paragraphs[0].Portions[0].Font.ColorFormat;

            // Act-Assert
            colorFormatC1.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC2.Color.Should().Be(ColorTranslator.FromHtml("#000000"));
            colorFormatC3.Color.Should().Be(ColorTranslator.FromHtml("#595959"));
            colorFormatC4.Color.Should().Be(ColorTranslator.FromHtml("#FFFFFF"));
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
    }
}
