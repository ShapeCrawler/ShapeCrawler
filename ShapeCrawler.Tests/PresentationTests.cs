using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Statics;
using ShapeCrawler.Tests.Helpers;
using Xunit;

namespace ShapeCrawler.Tests
{
    public class PresentationTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public PresentationTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Close_ClosesPresentationAndReleasesResources()
        {
            // Arrange
            string originFilePath = Path.GetTempFileName();
            string savedAsFilePath = Path.GetTempFileName();
            File.WriteAllBytes(originFilePath, TestFiles.Presentations.pre001);
            IPresentation presentation = SCPresentation.Open(originFilePath, true);
            presentation.SaveAs(savedAsFilePath);

            // Act
            presentation.Close();

            // Assert
            Action act = () => presentation = SCPresentation.Open(originFilePath, true);
            act.Should().NotThrow<IOException>();
            presentation.Close();

            // Clean up
            File.Delete(originFilePath);
            File.Delete(savedAsFilePath);
        }

        [Fact]
        public void Close_ShouldNotThrowObjectDisposedException()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(TestFiles.Presentations.pre025_byteArray, true);
            MemoryStream mStream = new();
            IPieChart chart = (IPieChart)presentation.Slides[0].Shapes.First(sp => sp.Id == 7);
            chart.Categories[0].Name = "new name";
            presentation.SaveAs(mStream);

            // Act
            Action act = () => presentation.Close();

            // Assert
            act.Should().NotThrow<ObjectDisposedException>();
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
        public void Slide_Width_returns_presentation_slides_width_in_pixels()
        {
            // Arrange
            var presentation = _fixture.Pre009;

            // Act
            var slideWidth = presentation.SlideWidth;

            // Assert
            slideWidth.Should().Be(960);
        }
        
        [Fact]
        public void Slide_Height_returns_presentation_slides_height_in_pixels()
        {
            // Arrange
            var presentation = _fixture.Pre009;

            // Act
            var slideHeight = presentation.SlideHeight;

            // Assert
            slideHeight.Should().Be(540);
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
        public void SlidesAdd_AddsSpecifiedSlideAtTheEndOfTheSlideCollection()
        {
            // Arrange
            ISlide sourceSlide = _fixture.Pre001.Slides[0];
            IPresentation destPre = SCPresentation.Open(Properties.Resources._002, true);
            int originSlidesCount = destPre.Slides.Count;
            int expectedSlidesCount = originSlidesCount + 1;
            MemoryStream savedPre = new ();

            // Act
            destPre.Slides.Add(sourceSlide);

            // Assert
            destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

            destPre.SaveAs(savedPre);
            destPre = SCPresentation.Open(savedPre, false);
            destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
        }

        [Fact]
        public void SlidesAdd_ShouldNotBreakPresentation()
        {
            // Arrange
            ISlide sourceSlide = _fixture.Pre001.Slides[0];
            IPresentation destPre = SCPresentation.Open(Properties.Resources._002, true);
            MemoryStream modified = new();

            // Act
            destPre.Slides.Add(sourceSlide);

            // Assert
            destPre.SaveAs(modified);
            ValidateResponse response = PptxValidator.Validate(modified);
            response.IsValid.Should().BeTrue();
        }

        [Fact]
        public void SlidesInsert_InsertsSpecifiedSlideAtTheSpecifiedPosition()
        {
            // Arrange
            ISlide sourceSlide = SCPresentation.Open(TestFiles.Presentations.pre001, true).Slides[0];
            string sourceSlideId = Guid.NewGuid().ToString();
            sourceSlide.CustomData = sourceSlideId;
            IPresentation destPre = SCPresentation.Open(Properties.Resources._002, true);

            // Act
            destPre.Slides.Insert(2, sourceSlide);

            // Assert
            destPre.Slides[1].CustomData.Should().Be(sourceSlideId);
        }

        [Theory]
        [MemberData(nameof(TestCasesSlidesRemove))]
        public void SlidesRemove_RemovesFirstSlideFromPresentation_WhenFirstSlideObjectWasPassed(byte[] pptxBytes, int expectedSlidesCount)
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(pptxBytes, true);
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
            IPresentation presentationCase1 = _fixture.Pre001;
            IPresentation presentationCase2 = _fixture.Pre002;

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
            IPresentation presentation = _fixture.Pre001;

            // Act
            int slideMasterShapesCount = presentation.SlideMasters[0].Shapes.Count;

            // Assert
            slideMasterShapesCount.Should().Be(7);
        }
    }
}
