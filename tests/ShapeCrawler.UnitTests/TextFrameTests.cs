using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.TextBody;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Properties;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TextFrameTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public TextFrameTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void HasTextFrameAndText_PropertiesReturnCorrectValue_WhenTheirGettersAreCalled()
        {
            // Arrange
            var shapeCase1 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3);
            var shapeCase2 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2);
            var textFrameCase1 = shapeCase1.TextFrame;
            var textFrameCase2 = shapeCase2.TextFrame;

            // Act
            var hasTextFrameCase1 = shapeCase1.HasTextFrame;
            var hasTextFrameCase2 = shapeCase2.HasTextFrame;
            var frameTextCase1 = textFrameCase1.Text;
            var frameTextCase2 = textFrameCase2.Text;

            // Assert
            hasTextFrameCase1.Should().BeTrue();
            hasTextFrameCase2.Should().BeTrue();
            frameTextCase1.Should().BeEquivalentTo("25.01.2020");
            frameTextCase2.Should().BeEquivalentTo("test footer");
        }

        [Fact]
        public void ParagraphBulletFontNameProperty_ReturnsFontName()
        {
            // Arrange
            var shapes = _fixture.Pre002.Slides[1].Shapes;
            var shape3Pr1Bullet = shapes.First(x => x.Id == 3).TextFrame.Paragraphs[0].Bullet;
            var shape4Pr2Bullet = shapes.First(x => x.Id == 4).TextFrame.Paragraphs[1].Bullet;

            // Act
            var shape3BulletFontName = shape3Pr1Bullet.FontName;
            var shape4BulletFontName = shape4Pr2Bullet.FontName;

            // Assert
            shape3BulletFontName.Should().BeNull();
            shape4BulletFontName.Should().Be("Calibri");
        }

        [Fact]
        public void ParagraphBulletTypeProperty_ReturnsCorrectValue()
        {
            // Arrange
            var shapeList = _fixture.Pre002.Slides[1].Shapes;
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape5 = shapeList.First(x => x.Id == 5);
            var shape4Pr2Bullet = shape4.TextFrame.Paragraphs[1].Bullet;
            var shape5Pr1Bullet = shape5.TextFrame.Paragraphs[0].Bullet;
            var shape5Pr2Bullet = shape5.TextFrame.Paragraphs[1].Bullet;

            // Act
            var shape5Pr1BulletType = shape5Pr1Bullet.Type;
            var shape5Pr2BulletType = shape5Pr2Bullet.Type;
            var shape4Pr2BulletType = shape4Pr2Bullet.Type;

            // Assert
            shape5Pr1BulletType.Should().BeEquivalentTo(BulletType.Numbered);
            shape5Pr2BulletType.Should().BeEquivalentTo(BulletType.Picture);
            shape4Pr2BulletType.Should().BeEquivalentTo(BulletType.Character);
        }

        [Fact]
        public void ParagraphBulletColorHexAndCharAndSizeProperties_ReturnCorrectValues()
        {
            // Arrange
            var shapeList = _fixture.Pre002.Slides[1].Shapes;
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape4Pr2Bullet = shape4.TextFrame.Paragraphs[1].Bullet;

            // Act
            var bulletColorHex = shape4Pr2Bullet.ColorHex;
            var bulletChar = shape4Pr2Bullet.Char;
            var bulletSize = shape4Pr2Bullet.Size;

            // Assert
            bulletColorHex.Should().Be("C00000");
            bulletChar.Should().Be("'");
            bulletSize.Should().Be(120);
        }

        [Fact]
        public void ParagraphPortionRemove_RemovesPortionFromCollection()
        {
            // Arrange
            var presentation = Presentation.Open(Properties.Resources._002, true);
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
            var savedPresentation = new Presentation(memoryStream, false);
            portions = GetPortions(savedPresentation);
            portions.Should().HaveCount(1);
        }

        [Fact]
        public void ParagraphPortionFontName_GetterReturnsFontNameOfParagraphPortion()
        {
            // Arrange
            var textFrameCase1 = _fixture.Pre002.Slides[1].Shapes.First(s => s.Id == 3).TextFrame;
            var textFrameCase2 = _fixture.Pre001.Slides[0].Shapes.First(s => s.Id == 4).TextFrame;

            // Act
            var portionFontNameCase1 = textFrameCase1.Paragraphs[0].Portions[0].FontName;
            var portionFontNameCase2 = textFrameCase2.Paragraphs[0].Portions[0].FontName;

            // Assert
            portionFontNameCase1.Should().BeEquivalentTo("Palatino Linotype");
            portionFontNameCase2.Should().BeEquivalentTo("Broadway");
        }

        [Theory]
        [MemberData(nameof(ParagraphTextTestCases))]
        public void ParagraphText_IsChanged_WhenTextIsChangedViaSetter(Paragraph paragraph)
        {
            // Arrange
            const string expectedText = "a new paragraph text";

            // Act
            paragraph.Text = expectedText;

            // Assert
            paragraph.Text.Should().BeEquivalentTo(expectedText);
            paragraph.Portions.Should().HaveCount(1);
        }

        [Fact]
        public void ParagraphTextProperty_ReturnsCorrectValue_WhenItsGetterIsCalled()
        {
            // Arrange
            var presentation = _fixture.Pre008;
            var textFrame = presentation.Slides.First().Shapes.Single(e => e.Id == 37).TextFrame;
            var paragraph1 = textFrame.Paragraphs[0];
            var paragraph2 = textFrame.Paragraphs[1];

            // Act
            var paragraphText1 = paragraph1.Text;
            var paragraphText2 = paragraph2.Text;

            // Assert
            paragraphText1.Should().BeEquivalentTo("P1t1 P1t2");
            paragraphText2.Should().BeEquivalentTo("p2");
        }

        #region Helpers

        public static IEnumerable<object[]> ParagraphTextTestCases()
        {
            var paragraphNumber = 2;
            var pre002 = Presentation.Open(Resources._002, true);
            var shape4 = pre002.Slides[1].Shapes.First(x => x.Id == 4);
            var paragraph = shape4.TextFrame.Paragraphs[--paragraphNumber];
            yield return new[] {paragraph};

            paragraphNumber = 3;
            pre002 = Presentation.Open(Resources._002, true);
            shape4 = pre002.Slides[1].Shapes.First(x => x.Id == 4);
            paragraph = shape4.TextFrame.Paragraphs[--paragraphNumber];
            yield return new[] { paragraph };
        }

        private static PortionCollection GetPortions(Presentation presentation)
        {
            var shape5 = presentation.Slides[1].Shapes.First(x => x.Id == 5);
            var portions = shape5.TextFrame.Paragraphs.First().Portions;
            return portions;
        }

        #endregion Helpers
    }
}
