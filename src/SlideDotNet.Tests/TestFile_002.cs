using System;
using System.IO;
using SlideDotNet.Models;
using Xunit;
using System.Linq;
using FluentAssertions;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace SlideDotNet.Tests
{
    public class TestFile_002Fixture : IDisposable
    {
        public PresentationEx pre002 { get; }

        public TestFile_002Fixture()
        {
            var ms = new MemoryStream(Properties.Resources._002);
            pre002 = new PresentationEx(ms);
        }

        public void Dispose()
        {
            pre002.Close();
        }
    }

    public class TestFile_002 : IClassFixture<TestFile_002Fixture>
    {
        private readonly TestFile_002Fixture _fixture;

        public TestFile_002(TestFile_002Fixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Slide_TestShapeCountAndBulletHex()
        {
            // Arrange
            var sld2Shape4 = _fixture.pre002.Slides[1].Shapes.First(x => x.Id == 4);

            // Act
            var shapeList = _fixture.pre002.Slides[0].Shapes;
            var bulletHex = sld2Shape4.TextFrame.Paragraphs[1].Bullet.ColorHex;

            // Assert
            shapeList.Should().HaveCount(3);
            bulletHex.Should().Be("C00000");
        }
    }
}
