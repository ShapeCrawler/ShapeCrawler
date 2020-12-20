using System.Linq;
using FluentAssertions;
using SlideDotNet.Enums;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TestFile_002 : IClassFixture<TestFile_002Fixture>
    {
        private readonly TestFile_002Fixture _fixture;

        public TestFile_002(TestFile_002Fixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void ShapesCount_ShouldReturn3_SlideContain3Shapes()
        {
            // Arrange
            var shapes = _fixture.pre002.Slides[0].Shapes;

            // Act
            var shapesCount = shapes.Count;

            // Assert
            shapesCount.Should().Be(3);
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
        public void Bullet_ColorHexAndCharAndSize()
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
    }
}
