using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.TextBody;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TextFrameTests : IClassFixture<TestFileFixture>
    {
        private readonly TestFileFixture _fixture;

        public TextFrameTests(TestFileFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Bullet_FontName()
        {
            // Arrange
            var shapeList = _fixture.pre002.Slides[1].Shapes;
            var shape3 = shapeList.First(x => x.Id == 3);
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape3Pr1Bullet = shape3.TextFrame.Paragraphs[0].Bullet;
            var shape4Pr2Bullet = shape4.TextFrame.Paragraphs[1].Bullet;

            // Act
            var shape3BulletFontName = shape3Pr1Bullet.FontName;
            var shape4BulletFontName = shape4Pr2Bullet.FontName;

            // Assert
            shape3BulletFontName.Should().BeNull();
            shape4BulletFontName.Should().Be("Calibri");
        }

        [Fact]
        public void Bullet_Type()
        {
            // Arrange
            var shapeList = _fixture.pre002.Slides[1].Shapes;
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
        public void ParagraphBullet_ColorHexAndCharAndSize()
        {
            // Arrange
            var shapeList = _fixture.pre002.Slides[1].Shapes;
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
            var presentation = PresentationEx.Open(Properties.Resources._002, true);
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
            var savedPresentation = new PresentationEx(memoryStream, false);
            portions = GetPortions(savedPresentation);
            portions.Should().HaveCount(1);
        }

        private static PortionCollection GetPortions(PresentationEx presentation)
        {
            var shape5 = presentation.Slides[1].Shapes.First(x => x.Id == 5);
            var portions = shape5.TextFrame.Paragraphs.First().Portions;
            return portions;
        }

        private static string GetParagraphText(PresentationEx presentation)
        {
            var shape4 = presentation.Slides[1].Shapes.First(x => x.Id == 4);
            var paragraphText = shape4.TextFrame.Paragraphs.First().Text;
            return paragraphText;
        }
    }
}
