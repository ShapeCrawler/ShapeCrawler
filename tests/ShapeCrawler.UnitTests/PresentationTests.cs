using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Models;
using ShapeCrawler.Statics;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    public class PresentationTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public PresentationTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }


        [Fact]
        public void Open_ThrowsPresentationIsLargeException_WhenThePresentationContentSizeIsBeyondThePermitted()
        {
            // Arrange
            var bytes = new byte[Limitations.MaxPresentationSize + 1];

            // Act
            Action act = () => Presentation.Open(bytes, false);

            // Arrange
            act.Should().Throw<PresentationIsLargeException>();
        }

        [Fact]
        public void Slides_CollectionContainsNumberOfSlidesInThePresentation()
        {
            // Act
            var slidesCollectionCase2 = _fixture.Pre017.Slides;

            // Assert
            slidesCollectionCase2.Should().HaveCount(1);
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
