using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

namespace ShapeCrawler.Tests.Unit
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
            ITextBox textBox1 = ((IAutoShape)_fixture.Pre002.Slides[1].Shapes.First(sp => sp.Id == 3)).TextBox;
            ITextBox textBox2 = ((IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 4)).TextBox;
            ITextBox textBox3 = ((IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 7)).TextBox;

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
            ITextBox textBox = ((IAutoShape)_fixture.Pre001.Slides[4].Shapes.First(sp => sp.Id == 5)).TextBox;

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
                SCPresentation.Open(Resources._001, true).Slides[0].Shapes.First(sp => sp.Id == 4) as IAutoShape;
            IPortion paragraphPortion = autoShape.TextBox.Paragraphs[0].Portions[0];

            // Act
            paragraphPortion.Font.Name = newFont;

            // Assert
            paragraphPortion.Font.Name.Should().BeEquivalentTo(newFont);
        }

        [Fact]
        public void Name_SetterThrowsException_WhenAnUserTryChangeFontNameForPortionOfAPlaceholderShape()
        {
            // Arrange
            IAutoShape autoShape = (IAutoShape)SCPresentation.Open(Resources._001, true).Slides[2].Shapes
                .First(sp => sp.Id == 4);
            IPortion paragraphPortion = autoShape.TextBox.Paragraphs[0].Portions[0];

            // Act
            Action action = () => paragraphPortion.Font.Name = "Time New Roman";

            // Assert
            action.Should().Throw<PlaceholderCannotBeChangedException>();
        }

        [Fact]
        public void Size_GetterReturnsFontSizeOfTheParagraphPortion()
        {
            // Arrange
            IPortion portionCase1 = ((IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase2 = ((IAutoShape)_fixture.Pre015.Slides[0].Shapes.First(sp => sp.Id == 5)).TextBox.Paragraphs[0].Portions[2];
            IPortion portionCase3 = ((IAutoShape)_fixture.Pre015.Slides[1].Shapes.First(sp => sp.Id == 61)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase4 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase5 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[1];
            IPortion portionCase6 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase7 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 3)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase8 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 4103)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase9 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase10 = ((IAutoShape)_fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase11 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase12 = ((IAutoShape)_fixture.Pre010.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase13 = ((IAutoShape)_fixture.Pre014.Slides[3].Shapes.First(sp => sp.Id == 5)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase14 = ((IAutoShape)_fixture.Pre014.Slides[4].Shapes.First(sp => sp.Id == 4)).TextBox.Paragraphs[0].Portions[0];
            IPortion portionCase15 = ((IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52)).TextBox.Paragraphs[0].Portions[0];

            // Act-Assert
            portionCase1.Font.Size.Should().Be(1800);
            portionCase2.Font.Size.Should().Be(1800);
            portionCase3.Font.Size.Should().Be(1867);
            portionCase4.Font.Size.Should().Be(1800);
            portionCase5.Font.Size.Should().Be(2000);
            portionCase6.Font.Size.Should().Be(4400);
            portionCase7.Font.Size.Should().Be(3200);
            portionCase8.Font.Size.Should().Be(1800);
            portionCase9.Font.Size.Should().Be(1200);
            portionCase10.Font.Size.Should().Be(2177);
            portionCase11.Font.Size.Should().Be(2000);
            portionCase12.Font.Size.Should().Be(1539);
            portionCase13.Font.Size.Should().Be(1200);
            portionCase14.Font.Size.Should().Be(1200);
            portionCase15.Font.Size.Should().Be(2700);
        }

        [Fact]
        public void Size_GetterReturnsFontSize_OfPlaceholder()
        {
            // Arrange
            IAutoShape autoShapeCase1 = (IAutoShape) _fixture.Pre028.Slides[0].Shapes.First(sp => sp.Id == 4098);
            IAutoShape autoShapeCase2 = (IAutoShape) _fixture.Pre029.Slides[0].Shapes.First(sp => sp.Id == 3);
            IPortion portionC1 = autoShapeCase1.TextBox.Paragraphs[0].Portions[0];
            IPortion portionC2 = autoShapeCase2.TextBox.Paragraphs[0].Portions[0];

            // Act-Assert
            portionC1.Font.Size.Should().Be(3200);
            portionC2.Font.Size.Should().Be(2500);
        }

        [Fact]
        public void Size_SetterChangesFontSizeOfParagraphPortion()
        {
            // Arrange
            int newFontSize = 28;
            var savedPreStream = new MemoryStream();
            SCPresentation presentation = SCPresentation.Open(Resources._001, true);
            IPortion portion = GetPortion(presentation);
            int oldFontSize = portion.Font.Size;

            // Act
            portion.Font.Size = newFontSize;
            presentation.SaveAs(savedPreStream);

            // Assert
            presentation = SCPresentation.Open(savedPreStream, false);
            portion = GetPortion(presentation);
            portion.Font.Size.Should().NotBe(oldFontSize);
            portion.Font.Size.Should().Be(newFontSize);
            portion.Font.SizeCanBeChanged().Should().BeTrue();
        }

        [Fact]
        public void IsBold_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsBold()
        {
            // Arrange
            IAutoShape nonPlaceholderAutoShapeCase1 = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3);
            IFont fontC1 = nonPlaceholderAutoShapeCase1.TextBox.Paragraphs[0].Portions[0].Font;

            // Act-Assert
            fontC1.IsBold.Should().BeTrue();
        }

        [Fact]
        public void IsBold_GetterReturnsTrue_WhenFontOfPlaceholderTextIsBold()
        {
            // Arrange
            IAutoShape placeholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[1].Shapes.First(sp => sp.Id == 6);
            IPortion portion = placeholderAutoShape.TextBox.Paragraphs[0].Portions[0];

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
            IPortion portion = nonPlaceholderAutoShape.TextBox.Paragraphs[0].Portions[0];

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
            IPortion portion = placeholderAutoShape.TextBox.Paragraphs[0].Portions[0];

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
            IPresentation presentation = SCPresentation.Open(Resources._020, true);
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            IPortion portion = nonPlaceholderAutoShape.TextBox.Paragraphs[0].Portions[0];

            // Act
            portion.Font.IsBold = true;

            // Assert
            portion.Font.IsBold.Should().BeTrue();
            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            portion = nonPlaceholderAutoShape.TextBox.Paragraphs[0].Portions[0];
            portion.Font.IsBold.Should().BeTrue();
        }

        [Theory]
        [MemberData(nameof(TestCasesIsBold))]
        public void IsBold_Setter_AddsBoldForPlaceholderTextFont(SCPresentation presentation, SlideElementQuery portionRequest)
        {
            // Arrange
            MemoryStream mStream = new ();
            IPortion portion = TestHelper.GetParagraphPortion(presentation, portionRequest);

            // Act
            portion.Font.IsBold = true;

            // Assert
            portion.Font.IsBold.Should().BeTrue();

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            portion = TestHelper.GetParagraphPortion(presentation, portionRequest);
            portion.Font.IsBold.Should().BeTrue();
        }

        public static IEnumerable<object[]> TestCasesIsBold()
        {
            SCPresentation presentationCase1 = SCPresentation.Open(Resources._020, true);
            SlideElementQuery portionRequestCase1 = new();
            portionRequestCase1.SlideIndex = 2;
            portionRequestCase1.ShapeId = 7;
            portionRequestCase1.ParagraphIndex = 0;
            portionRequestCase1.PortionIndex = 0;

            SCPresentation presentationCase2 = SCPresentation.Open(Resources._026, true);
            SlideElementQuery portionRequestCase2 = new();
            portionRequestCase2.SlideIndex = 0;
            portionRequestCase2.ShapeId = 128;
            portionRequestCase2.ParagraphIndex = 0;
            portionRequestCase2.PortionIndex = 0;

            var testCases = new List<object[]>
            {
                new object[] {presentationCase1, portionRequestCase1},
                new object[] {presentationCase2, portionRequestCase2}
            };

            return testCases;
        }

        [Fact]
        public void IsItalic_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsItalic()
        {
            // Arrange
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3);
            IFont font = nonPlaceholderAutoShape.TextBox.Paragraphs[0].Portions[0].Font;

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
            IPortion portion = placeholderAutoShape.TextBox.Paragraphs[0].Portions[0];

            // Act-Assert
            portion.Font.IsItalic.Should().BeTrue();
        }

        [Fact]
        public void IsItalic_Setter_SetsItalicFontForForNonPlaceholderText()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020, true);
            IAutoShape nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            IPortion portion = nonPlaceholderAutoShape.TextBox.Paragraphs[0].Portions[0];

            // Act
            portion.Font.IsItalic = true;

            // Assert
            portion.Font.IsItalic.Should().BeTrue();
            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
            portion = nonPlaceholderAutoShape.TextBox.Paragraphs[0].Portions[0];
            portion.Font.IsItalic.Should().BeTrue();
        }

        [Fact]
        public void IsItalic_SetterSetsNonItalicFontForPlaceholderText_WhenFalseValueIsPassed()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020, true);
            IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            IPortion portion = placeholderAutoShape.TextBox.Paragraphs[0].Portions[0];

            // Act
            portion.Font.IsItalic = false;

            // Assert
            portion.Font.IsItalic.Should().BeFalse();
            presentation.SaveAs(mStream);

            presentation = SCPresentation.Open(mStream, false);
            placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            portion = placeholderAutoShape.TextBox.Paragraphs[0].Portions[0];
            portion.Font.IsItalic.Should().BeFalse();
        }

        private static IPortion GetPortion(SCPresentation presentation)
        {
            IAutoShape autoShape = presentation.Slides[0].Shapes.First(sp => sp.Id == 4) as IAutoShape;
            IPortion portion = autoShape.TextBox.Paragraphs[0].Portions[0];
            return portion;
        }
    }
}
