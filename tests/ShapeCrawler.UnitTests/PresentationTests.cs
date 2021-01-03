using FluentAssertions;
using ShapeCrawler.Models;
using System.IO;
using System.Linq;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

namespace ShapeCrawler.UnitTests
{
    public class PresentationTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public PresentationTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Slides_CollectionContainsTwoItems_WhenThePresentationHasTwoSlides()
        {
            // Act
            var slides = _fixture.Pre001.Slides;

            // Assert
            slides.Should().HaveCount(2);
        }

        [Fact]
        public void SlidesRemove_RemovesSlideFromPresentation_WhenSlideInstanceIsPassedInTheMethod()
        {
            // Arrange
            var stream = new MemoryStream(Properties.Resources._007_2_slides);
            var presentation = new Presentation(stream, true);
            var removingSlide = presentation.Slides.First();

            // Act
            presentation.Slides.Remove(removingSlide);
            presentation.Close();

            // Assert
            presentation = new Presentation(stream, false);
            presentation.Slides.Should().HaveCount(1);
        }
    }
}
