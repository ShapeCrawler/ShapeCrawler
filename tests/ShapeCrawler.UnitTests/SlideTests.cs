using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Services.ShapeCreators;
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
            var pre = Presentation.Open(Properties.Resources._001, true);
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
            Slide slide = _fixture.Pre002.Slides[2];

            // Act
            bool hidden = slide.Hidden;

            // Assert
            hidden.Should().BeTrue();
        }

        [Fact]
        public async void BackgroundSetImage_ChangesBackground_WhenImageStreamIsPassed()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._009);
            var backgroundImage = pre.Slides[0].Background;
            var imgStream = new MemoryStream(Properties.Resources.test_image_2);
            var bytesBefore = await backgroundImage.GetImageBytesValueTask();

            // Act
            backgroundImage.SetImage(imgStream);

            // Assert
            var bytesAfter = await backgroundImage.GetImageBytesValueTask();
            bytesAfter.Length.Should().NotBe(bytesBefore.Length);
        }

        [Fact]
        public void CustomData_ReturnsData_WhenCustomDataWasAssigned()
        {
            // Arrange
            const string customDataString = "Test custom data";
            var origPreStream = new MemoryStream();
            origPreStream.Write(Properties.Resources._001);
            var originPre = new Presentation(origPreStream, true);
            var slide = originPre.Slides.First();

            // Act
            slide.CustomData = customDataString;

            var savedPreStream = new MemoryStream();
            originPre.SaveAs(savedPreStream);
            var savedPre = new Presentation(savedPreStream, false);
            var customData = savedPre.Slides.First().CustomData;

            // Assert
            customData.Should().Be(customDataString);
        }

        [Fact]
        public void Shapes_ReturnsShapeCollectionWithCorrectShapeContentType()
        {
            // Arrange
            var pre = _fixture.Pre003;

            // Act
            var shapes = pre.Slides.First().Shapes;

            // Assert
            Assert.Single(shapes.Where(c => c.ContentType.Equals(ShapeContentType.AutoShape)));
            Assert.Single(shapes.Where(c => c.ContentType.Equals(ShapeContentType.Picture)));
            Assert.Single(shapes.Where(c => c.ContentType.Equals(ShapeContentType.Table)));
            Assert.Single(shapes.Where(c => c.ContentType.Equals(ShapeContentType.Chart)));
            Assert.Single(shapes.Where(c => c.ContentType.Equals(ShapeContentType.Group)));
        }
    }
}
