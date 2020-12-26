using FluentAssertions;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TestFile_002_Read : IClassFixture<TestFileFixture>
    {
        private readonly TestFileFixture _fixture;

        public TestFile_002_Read(TestFileFixture fixture)
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
    }
}
