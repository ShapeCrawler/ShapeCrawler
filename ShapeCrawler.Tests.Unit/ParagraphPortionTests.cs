using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class ParagraphPortionTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ParagraphPortionTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Remove_RemovesPortionFromCollection()
        {
            // Arrange
            var presentation = PresentationSc.Open(Resources._002, true);
            var portions = GetPortions(presentation);
            var portion = portions.First();
            var countBefore = portions.Count;

            // Act
            portion.Remove();

            // Assert
            portions.Should().HaveCount(1);
            portions.Should().HaveCountLessThan(countBefore);

            var memoryStream = new MemoryStream();
            presentation.SaveAs(memoryStream);
            var savedPresentation = PresentationSc.Open(memoryStream, false);
            portions = GetPortions(savedPresentation);
            portions.Should().HaveCount(1);
        }

        [Fact]
        public void Text_GetterReturnsParagraphPortionText()
        {
            // Arrange
            Portion portion = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox
                .Paragraphs[0].Portions[0];

            // Act
            string paragraphPortionText = portion.Text;

            // Assert
            paragraphPortionText.Should().BeEquivalentTo("0:0_p1_lvl1");
        }

        [Fact]
        public void FontName_GetterReturnsFontNameOfTheParagraphPortion()
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

        [Fact(Skip = "In Progress")]
        public void FontName_GetterReturnsCalibriLightAsFontName()
        {
            // Arrange
            ITextBox textBox4 = ((IAutoShape)_fixture.Pre001.Slides[4].Shapes.First(sp => sp.Id == 5)).TextBox;

            // Act
            string portionFontNameCase4 = textBox4.Paragraphs[0].Portions[0].Font.Name;

            // Assert
            portionFontNameCase4.Should().BeEquivalentTo("Calibri Light");
        }

        [Fact]
        public void FontName_SetterChangeFontName()
        {
            // Arrange
            const string newFont = "Time New Roman";
            IAutoShape autoShape =
                PresentationSc.Open(Resources._001, true).Slides[0].Shapes.First(sp => sp.Id == 4) as IAutoShape;
            Portion paragraphPortion = autoShape.TextBox.Paragraphs[0].Portions[0];
            
            // Act
            paragraphPortion.Font.Name = newFont;

            // Assert
            paragraphPortion.Font.Name.Should().BeEquivalentTo(newFont);
        }

        [Fact]
        public void FontName_SetterThrowsException_WhenAnUserTryChangeFontNameForPortionOfAPlaceholderShape()
        {
            // Arrange
            IAutoShape autoShape = (IAutoShape) PresentationSc.Open(Resources._001, true).Slides[2].Shapes
                .First(sp => sp.Id == 4);
            Portion paragraphPortion = autoShape.TextBox.Paragraphs[0].Portions[0];

            // Act
            Action action = () => paragraphPortion.Font.Name = "Time New Roman";

            // Assert
            action.Should().Throw<PlaceholderCannotBeChangedException>();
        }

        [Fact]
        public void FontSize_GetterReturnsFontSizeOfTheParagraphPortion()
        {
            // Arrange
            Portion portionCase1 = ((IAutoShape)_fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase2 = ((IAutoShape)_fixture.Pre015.Slides[0].Shapes.First(sp => sp.Id == 5)).TextBox.Paragraphs[0].Portions[2];
            Portion portionCase3 = ((IAutoShape)_fixture.Pre015.Slides[1].Shapes.First(sp => sp.Id == 61)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase4 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase5 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[1];
            Portion portionCase6 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase7 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 3)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase8 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 4103)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase9 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase10 = ((IAutoShape)_fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase11 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase12 = ((IAutoShape)_fixture.Pre010.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase13 = ((IAutoShape)_fixture.Pre014.Slides[3].Shapes.First(sp => sp.Id == 5)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase14 = ((IAutoShape)_fixture.Pre014.Slides[4].Shapes.First(sp => sp.Id == 4)).TextBox.Paragraphs[0].Portions[0];
            Portion portionCase15 = ((IAutoShape)_fixture.Pre014.Slides[5].Shapes.First(sp => sp.Id == 52)).TextBox.Paragraphs[0].Portions[0];

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
        public void FontSize_SetterChangesFontSizeOfParagraphPortion()
        {
            // Arrange
            static Portion GetPortion(PresentationSc presentation)
            {
                IAutoShape autoShape = presentation.Slides[0].Shapes.First(sp => sp.Id == 4) as IAutoShape;
                Portion portion = autoShape.TextBox.Paragraphs[0].Portions[0];
                return portion;
            }
            int newFontSize = 28;
            var savedPreStream = new MemoryStream();
            PresentationSc presentation = PresentationSc.Open(Resources._001, true);
            Portion portion = GetPortion(presentation);
            int oldFontSize = portion.Font.Size;

            // Act
            portion.Font.Size = newFontSize;
            presentation.SaveAs(savedPreStream);

            // Assert
            presentation = PresentationSc.Open(savedPreStream, false);
            portion = GetPortion(presentation);
            portion.Font.Size.Should().NotBe(oldFontSize);
            portion.Font.Size.Should().Be(newFontSize);
            portion.Font.SizeCanBeChanged().Should().BeTrue();
        }

        private static PortionCollection GetPortions(PresentationSc presentation)
        {
            IAutoShape shape5 = presentation.Slides[1].Shapes.First(x => x.Id == 5) as IAutoShape;
            var portions = shape5.TextBox.Paragraphs[0].Portions;
            return portions;
        }
    }
}
