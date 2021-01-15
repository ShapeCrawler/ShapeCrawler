using System;
using System.Linq;
using ShapeCrawler.Enums;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_019 : IClassFixture<TestFile_019Fixture>
    {
        private readonly TestFile_019Fixture _fixture;

        public TestFile_019(TestFile_019Fixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Picture_DoNotParseStrangePicture_Test()
        {
            // Arrange
            var pre = _fixture.pre019;

            // Act - Assert
            Assert.ThrowsAny<Exception>(() => pre.Slides[1].Shapes.Single(x => x.Id == 47));
        }
    }
}
