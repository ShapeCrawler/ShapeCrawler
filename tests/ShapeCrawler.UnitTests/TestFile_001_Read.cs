using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Models;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TestFile_001_Read
    {
        [Fact]
        public void SlidesCount_ShouldReturnTwo_PresentationContainsTwoSlides()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._001);

            // Act
            var sldNumber = pre.Slides.Count();
            
            // Assert
            sldNumber.Should().Be(2);
        }

        [Fact]
        public void Slide_CustomData_returns_null_when_CustomData_was_not_assigned()
        {
            var pre = new PresentationEx(Properties.Resources._001);
            var slide = pre.Slides.First();

            // Act
            var customData = slide.CustomData;

            // Assert
            customData.Should().BeNull();
        }
    }
}
