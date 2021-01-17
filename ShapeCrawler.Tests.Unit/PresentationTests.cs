using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Statics;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class PresentationTests : IClassFixture<PptxFixture>
    {
        private readonly PptxFixture _fixture;

        public PresentationTests(PptxFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Open_ThrowsPresentationIsLargeException_WhenThePresentationContentSizeIsBeyondThePermitted()
        {
            // Arrange
            var bytes = new byte[Limitations.MaxPresentationSize + 1];

            // Act
            Action act = () => PresentationSc.Open(bytes, false);

            // Arrange
            act.Should().Throw<PresentationIsLargeException>();
        }

        [Fact]
        public void SlideWidthAndSlideHeight_ReturnWidthAndHeightSizesOfThePresentationSlides()
        {
            // Arrange
            PresentationSc presentation = _fixture.Pre009;

            // Act
            int slideWidth = presentation.SlideWidth;
            int slideHeight = presentation.SlideHeight;

            // Assert
            slideWidth.Should().Be(9144000);
            slideHeight.Should().Be(5143500);
        }

        [Fact]
        public void SlidesCount_ReturnsOne_WhenPresentationContainsOneSlide()
        {
            // Act
            var numberSlidesCase1 = _fixture.Pre017.Slides.Count;
            var numberSlidesCase2 = _fixture.Pre016.Slides.Count;

            // Assert
            numberSlidesCase1.Should().Be(1);
            numberSlidesCase2.Should().Be(1);
        }

        [Fact]
        public void SlidesRemove_RemovesSlideFromPresentation_WhenSlideInstanceIsPassedInTheMethod()
        {
            // Arrange
            var stream = new MemoryStream(Properties.Resources._007_2_slides);
            var presentation = new PresentationSc(stream, true);
            var removingSlide = presentation.Slides.First();

            // Act
            presentation.Slides.Remove(removingSlide);
            presentation.Close();

            // Assert
            presentation = new PresentationSc(stream, false);
            presentation.Slides.Should().HaveCount(1);
        }

        [Fact]
        public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
        {
            // Arrange
            PresentationSc presentationCase1 = _fixture.Pre001;
            PresentationSc presentationCase2 = _fixture.Pre002;

            // Act
            int slideMastersCountCase1 = presentationCase1.SlideMasters.Count;
            int slideMastersCountCase2 = presentationCase2.SlideMasters.Count;

            // Assert
            slideMastersCountCase1.Should().Be(1);
            slideMastersCountCase2.Should().Be(2);
        }

        [Fact]
        public void SlideMasterShapesCount_ReturnsNumberOfShapesOnTheMasterSlide()
        {
            // Arrange
            PresentationSc presentation = _fixture.Pre001;

            // Act
            int slideMasterShapesCount = presentation.SlideMasters[0].Shapes.Count;

            // Assert
            slideMasterShapesCount.Should().Be(5);
        }
    }
}
