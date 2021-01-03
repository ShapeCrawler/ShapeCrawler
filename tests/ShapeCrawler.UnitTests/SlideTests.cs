using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class SlideTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public SlideTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Hide_MethodHidesSlide_WhenItIsExecuted()
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
        public void Hidden_GetterReturnsTrue_WhenTheSlideIsHidden()
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
            var bytesBefore = await backgroundImage.GetImageBytes();

            // Act
            backgroundImage.SetImage(imgStream);

            // Assert
            var bytesAfter = await backgroundImage.GetImageBytes();
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
        public void Shapes_ContainsParticularShapeTypes()
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

        [Theory]
        [MemberData(nameof(ShapesCollectionTestCases))]
        public void Shapes_CountPropertyReturnsNumberOfTheShapesOnTheSlide(Slide slide, int expectedShapesNumber)
        {
            // Act
            var shapes = slide.Shapes;

            // Assert
            shapes.Should().HaveCount(expectedShapesNumber);
        }

        public static IEnumerable<object[]> ShapesCollectionTestCases()
        {
            var slide = Presentation.Open(Properties.Resources._002, false).Slides[0];
            yield return new object[] { slide, 3 };

            slide = Presentation.Open(Properties.Resources._003, false).Slides[0];
            yield return new object[] { slide, 5 };

            slide = Presentation.Open(Properties.Resources._009, false).Slides[0];
            yield return new object[] { slide, 6 };

            slide = Presentation.Open(Properties.Resources._009, false).Slides[1];
            yield return new object[] { slide, 6 };

            slide = Presentation.Open(Properties.Resources._013, false).Slides[0];
            yield return new object[] { slide, 4 };
        }

        [Fact]
        public void CustomData_PropertyIsNull_WhenTheSlideHasNotCustomData()
        {
            // Arrange
            var slide = _fixture.Pre001.Slides.First();

            // Act
            var sldCustomData = slide.CustomData;

            // Assert
            sldCustomData.Should().BeNull();
        }

        [Fact]
        public void HasPicture_ReturnsTrue_WhenTheShapeContainsImageContent()
        {
            // Arrange
            var shape = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);

            // Act
            var shapeHasPicture = shape.HasPicture;

            // Assert
            shapeHasPicture.Should().BeTrue();
        }
    }
}
