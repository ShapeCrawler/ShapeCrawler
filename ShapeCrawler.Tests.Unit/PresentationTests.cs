using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Statics;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    public class PresentationTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public PresentationTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Open_ThrowsPresentationIsLargeException_WhenThePresentationContentSizeIsBeyondThePermitted()
        {
            // Arrange
            var bytes = new byte[Limitations.MaxPresentationSize + 1];

            // Act
            Action act = () => SCPresentation.Open(bytes, false);

            // Assert
            act.Should().Throw<PresentationIsLargeException>();
        }

        [Fact]
        public void SlideWidthAndSlideHeight_ReturnWidthAndHeightSizesOfThePresentationSlides()
        {
            // Arrange
            SCPresentation presentation = _fixture.Pre009;

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

#if DEBUG
        [Fact]
        public void SlidesAdd_AddsSpecifiedSlideAtTheEndOfTheSlideCollection()
        {
            // Arrange
            IPresentation destPre = SCPresentation.Open(Properties.Resources._002, true);
            int originSlidesCount = destPre.Slides.Count;
            int expectedSlidesCount = originSlidesCount + 1;
            ISlide addingSlide = _fixture.Pre001.Slides[0];
            MemoryStream savedPre = new ();

            // Act
            destPre.Slides.Add(addingSlide);

            // Assert
            destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

            destPre.SaveAs(savedPre);
            destPre = SCPresentation.Open(savedPre, false);
            destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
        }
#endif

        [Theory]
        [MemberData(nameof(TestCasesSlidesRemove))]
        public void SlidesRemove_RemovesFirstSlideFromPresentation_WhenFirstSlideObjectWasPassed(byte[] pptxBytes, int expectedSlidesCount)
        {
            // Arrange
            SCPresentation presentation = SCPresentation.Open(pptxBytes, true);
            ISlide removingSlide = presentation.Slides[0];
            var mStream = new MemoryStream();

            // Act
            presentation.Slides.Remove(removingSlide);

            // Assert
            presentation.Slides.Should().HaveCount(expectedSlidesCount);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            presentation.Slides.Should().HaveCount(expectedSlidesCount);
        }

        public static IEnumerable<object[]> TestCasesSlidesRemove()
        {
            yield return new object[] {Properties.Resources._007_2_slides, 1};
            yield return new object[] {Properties.Resources._006_1_slides, 0};
        }

        [Fact]
        public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
        {
            // Arrange
            SCPresentation presentationCase1 = _fixture.Pre001;
            SCPresentation presentationCase2 = _fixture.Pre002;

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
            SCPresentation presentation = _fixture.Pre001;

            // Act
            int slideMasterShapesCount = presentation.SlideMasters[0].Shapes.Count;

            // Assert
            slideMasterShapesCount.Should().Be(7);
        }
    }
}
