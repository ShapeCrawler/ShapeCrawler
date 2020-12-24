using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Models;
using Xunit;


// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class SlideTests : IClassFixture<TestFileFixture>
    {
        private readonly TestFileFixture _fixture;

        public SlideTests(TestFileFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Hide_HidesSlide()
        {
            // Arrange
            var pre = PresentationEx.Open(Properties.Resources._001, true);
            var slide = pre.Slides.First();

            // Act
            slide.Hide();

            // Assert
            slide.Hidden.Should().Be(true);
        }


        [Fact]
        public void Hidden_ReturnsTrue_WhenSlideIsHidden()
        { 
            // Arrange
            Slide slide = _fixture.pre002.Slides[2];

            // Act
            bool hidden = slide.Hidden;

            // Assert
            hidden.Should().BeTrue();
        }
    }
}
