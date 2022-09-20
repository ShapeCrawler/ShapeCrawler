using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
using Xunit;
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests
{
    public class FontTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public FontTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Name_GetterReturnsFontNameOfTheParagraphPortion()
        {
            // Arrange
            ITextFrame textBox1 = ((IAutoShape)_fixture.Pre002.Slides[1].Shapes.First(sp => sp.Id == 3)).TextFrame;
            ITextFrame textBox2 = ((IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 4)).TextFrame;
            ITextFrame textBox3 = ((IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 7)).TextFrame;

            // Act
            string portionFontNameCase1 = textBox1.Paragraphs[0].Portions[0].Font.Name;
            string portionFontNameCase2 = textBox2.Paragraphs[0].Portions[0].Font.Name;
            string portionFontNameCase3 = textBox3.Paragraphs[0].Portions[0].Font.Name;

            // Assert
            portionFontNameCase1.Should().BeEquivalentTo("Palatino Linotype");
            portionFontNameCase2.Should().BeEquivalentTo("Broadway");
            portionFontNameCase3.Should().BeEquivalentTo("Calibri Light");
        }

        [Fact]
        public void Name_GetterReturnsCalibriLightAsFontName_WhenFontNameIsCalibriLight()
        {
            // Arrange
            ITextFrame textBox = ((IAutoShape)_fixture.Pre001.Slides[4].Shapes.First(sp => sp.Id == 5)).TextFrame;

            // Act
            string portionFontName = textBox.Paragraphs[0].Portions[0].Font.Name;

            // Assert
            portionFontName.Should().BeEquivalentTo("Calibri Light");
        }

        [Fact]
        public void Name_SetterChangesFontName()
        {
            // Arrange
            const string newFont = "Time New Roman";
            IAutoShape autoShape =
                SCPresentation.Open(TestFiles.Presentations.pre001).Slides[0].Shapes.First(sp => sp.Id == 4) as IAutoShape;
            IPortion paragraphPortion = autoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            paragraphPortion.Font.Name = newFont;

            // Assert
            paragraphPortion.Font.Name.Should().BeEquivalentTo(newFont);
        }

        [Fact]
        public void Name_SetterThrowsException_WhenAnUserTryChangeFontNameForPortionOfAPlaceholderShape()
        {
            // Arrange
            IAutoShape autoShape = (IAutoShape)SCPresentation.Open(Resources._001).Slides[2].Shapes
                .First(sp => sp.Id == 4);
            IPortion paragraphPortion = autoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            Action action = () => paragraphPortion.Font.Name = "Time New Roman";

            // Assert
            action.Should().Throw<Exception>();
        }

        [Fact]
        public void Size_Getter_returns_font_size()
        {
            // Arrange
            IPortion portionCase1 = ((IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase2 = ((IAutoShape)_fixture.Pre015.Slides[0].Shapes.First(sp => sp.Id == 5)).TextFrame.Paragraphs[0].Portions[2];
            IPortion portionCase3 = ((IAutoShape)_fixture.Pre015.Slides[1].Shapes.First(sp => sp.Id == 61)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase4 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase5 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextFrame.Paragraphs[0].Portions[1];
            IPortion portionCase6 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase7 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 3)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase8 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 4103)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase9 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase10 = ((IAutoShape)_fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase11 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 2)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase12 = ((IAutoShape)_fixture.Pre010.Slides[0].Shapes.First(sp => sp.Id == 2)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase13 = ((IAutoShape)_fixture.Pre014.Slides[3].Shapes.First(sp => sp.Id == 5)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase14 = ((IAutoShape)_fixture.Pre014.Slides[4].Shapes.First(sp => sp.Id == 4)).TextFrame.Paragraphs[0].Portions[0];
            IPortion portionCase15 = ((IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52)).TextFrame.Paragraphs[0].Portions[0];

            // Act-Assert
            portionCase1.Font.Size.Should().Be(18);
            portionCase2.Font.Size.Should().Be(18);
            portionCase3.Font.Size.Should().Be(18);
            portionCase4.Font.Size.Should().Be(18);
            portionCase5.Font.Size.Should().Be(20);
            portionCase6.Font.Size.Should().Be(44);
            portionCase7.Font.Size.Should().Be(32);
            portionCase8.Font.Size.Should().Be(18);
            portionCase9.Font.Size.Should().Be(12);
            portionCase10.Font.Size.Should().Be(21);
            portionCase11.Font.Size.Should().Be(20);
            portionCase12.Font.Size.Should().Be(15);
            portionCase13.Font.Size.Should().Be(12);
            portionCase14.Font.Size.Should().Be(12);
            portionCase15.Font.Size.Should().Be(27);
        }

        [Fact]
        public void Size_Getter_returns_font_size_of_Placeholder()
        {
            // Arrange
            IAutoShape autoShapeCase1 = (IAutoShape)_fixture.Pre028.Slides[0].Shapes.First(sp => sp.Id == 4098);
            IAutoShape autoShapeCase2 = (IAutoShape)_fixture.Pre029.Slides[0].Shapes.First(sp => sp.Id == 3);
            IPortion portionC1 = autoShapeCase1.TextFrame.Paragraphs[0].Portions[0];
            IPortion portionC2 = autoShapeCase2.TextFrame.Paragraphs[0].Portions[0];

            // Act-Assert
            portionC1.Font.Size.Should().Be(32);
            portionC2.Font.Size.Should().Be(25);
        }

        [Fact]
        public void Size_Getter_returns_Font_Size_of_Non_Placeholder_Table()
        {
            // Arrange
            var table = (ITable)this._fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3);
            var cellPortion = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0];

            // Act-Assert
            cellPortion.Font.Size.Should().Be(18);
        }

        [Fact]
        public void Size_Setter_changes_Font_Size_of_paragraph_portion()
        {
            // Arrange
            int newFontSize = 28;
            var savedPreStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._001);
            IPortion portion = GetPortion(presentation);
            int oldFontSize = portion.Font.Size;

            // Act
            portion.Font.Size = newFontSize;

            // Assert
            presentation.SaveAs(savedPreStream);
            presentation = SCPresentation.Open(savedPreStream);
            portion = GetPortion(presentation);
            portion.Font.Size.Should().NotBe(oldFontSize);
            portion.Font.Size.Should().Be(newFontSize);
            portion.Font.CanChangeSize().Should().BeTrue();
        }

        [Fact]
        public void IsBold_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsBold()
        {
            // Arrange
            IAutoShape nonPlaceholderAutoShapeCase1 = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3);
            IFont fontC1 = nonPlaceholderAutoShapeCase1.TextFrame.Paragraphs[0].Portions[0].Font;

            // Act-Assert
            fontC1.IsBold.Should().BeTrue();
        }

        [Fact]
        public void IsBold_GetterReturnsTrue_WhenFontOfPlaceholderTextIsBold()
        {
            // Arrange
            IAutoShape placeholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[1].Shapes.First(sp => sp.Id == 6);
            IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            bool isBold = portion.Font.IsBold;

            // Assert
            isBold.Should().BeTrue();
        }

        [Fact]
        public void IsBold_GetterReturnsFalse_WhenFontOfNonPlaceholderTextIsNotBold()
        {
            // Arrange
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 2);
            IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            bool isBold = portion.Font.IsBold;

            // Assert
            isBold.Should().BeFalse();
        }

        [Fact]
        public void IsBold_GetterReturnsFalse_WhenFontOfPlaceholderTextIsNotBold()
        {
            // Arrange
            IAutoShape placeholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[2].Shapes.First(sp => sp.Id == 7);
            IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            bool isBold = portion.Font.IsBold;

            // Assert
            isBold.Should().BeFalse();
        }

        [Fact]
        public void IsBold_Setter_AddsBoldForNonPlaceholderTextFont()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020);
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion.Font.IsBold = true;

            // Assert
            portion.Font.IsBold.Should().BeTrue();
            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
            portion.Font.IsBold.Should().BeTrue();
        }

        [Theory]
        [MemberData(nameof(TestCasesIsBold))]
        public void IsBold_Setter_AddsBoldForPlaceholderTextFont(TestElementQuery portionQuery)
        {
            // Arrange
            MemoryStream mStream = new();
            var portion = portionQuery.GetParagraphPortion();
            var pres = portionQuery.Presentation;

            // Act
            portion.Font.IsBold = true;

            // Assert
            portion.Font.IsBold.Should().BeTrue();

            pres.SaveAs(mStream);
            pres = SCPresentation.Open(mStream);
            portionQuery.Presentation = pres;
            portion = portionQuery.GetParagraphPortion();
            portion.Font.IsBold.Should().BeTrue();
        }

        public static IEnumerable<object[]> TestCasesIsBold()
        {
            TestElementQuery portionRequestCase1 = new();
            portionRequestCase1.Presentation = SCPresentation.Open(Resources._020);
            portionRequestCase1.SlideIndex = 2;
            portionRequestCase1.ShapeId = 7;
            portionRequestCase1.ParagraphIndex = 0;
            portionRequestCase1.PortionIndex = 0;

            TestElementQuery portionRequestCase2 = new();
            portionRequestCase2.Presentation = SCPresentation.Open(Resources._026);
            portionRequestCase2.SlideIndex = 0;
            portionRequestCase2.ShapeId = 128;
            portionRequestCase2.ParagraphIndex = 0;
            portionRequestCase2.PortionIndex = 0;

            var testCases = new List<object[]>
            {
                new object[] {portionRequestCase1},
                new object[] {portionRequestCase2}
            };

            return testCases;
        }

        [Fact]
        public void IsItalic_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsItalic()
        {
            // Arrange
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3);
            IFont font = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0].Font;

            // Act
            bool isItalicFont = font.IsItalic;

            // Assert
            isItalicFont.Should().BeTrue();
        }

        [Fact]
        public void IsItalic_GetterReturnsTrue_WhenFontOfPlaceholderTextIsItalic()
        {
            // Arrange
            IAutoShape placeholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[2].Shapes.First(sp => sp.Id == 7);
            IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act-Assert
            portion.Font.IsItalic.Should().BeTrue();
        }

        [Fact]
        public void IsItalic_Setter_SetsItalicFontForForNonPlaceholderText()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020);
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion.Font.IsItalic = true;

            // Assert
            portion.Font.IsItalic.Should().BeTrue();
            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
            portion.Font.IsItalic.Should().BeTrue();
        }

        [Fact]
        public void IsItalic_SetterSetsNonItalicFontForPlaceholderText_WhenFalseValueIsPassed()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020);
            IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion.Font.IsItalic = false;

            // Assert
            portion.Font.IsItalic.Should().BeFalse();
            presentation.SaveAs(mStream);

            presentation = SCPresentation.Open(mStream);
            placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
            portion.Font.IsItalic.Should().BeFalse();
        }

        [Fact]
        public void Underline_SetUnderlineFont_WhenValueEqualsSetPassed()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020);
            IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion.Font.Underline = DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single;

            // Assert
            portion.Font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
            presentation.SaveAs(mStream);

            presentation = SCPresentation.Open(mStream);
            placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
            portion.Font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
        }

        private static IPortion GetPortion(IPresentation presentation)
        {
            IAutoShape autoShape = presentation.Slides[0].Shapes.First(sp => sp.Id == 4) as IAutoShape;
            IPortion portion = autoShape.TextFrame.Paragraphs[0].Portions[0];
            return portion;
        }
    }
}
