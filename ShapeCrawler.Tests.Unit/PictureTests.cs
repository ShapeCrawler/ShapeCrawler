using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class PictureTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public PictureTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public async void ImageExGetImageBytes_MethodReturnsNonEmptyShapeImage()
        {
            // Arrange
            IPicture shapePicture1 = (IPicture)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);
            IPicture shapePicture2 = (IPicture)_fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 7);

            // Act
            byte[] shapePictureContentCase1 = await shapePicture1.Image.GetImageBytes().ConfigureAwait(false);
            byte[] shapePictureContentCase2 = await shapePicture2.Image.GetImageBytes().ConfigureAwait(false);

            // Assert
            shapePictureContentCase1.Should().NotBeEmpty();
            shapePictureContentCase2.Should().NotBeEmpty();
        }

        [Fact]
        public async void ImageExSetImage_MethodSetsShapeImage_WhenCustomImageStreamIsPassed()
        {
            // Arrange
            var customImageStream = new MemoryStream(Properties.Resources.test_image_2);
            SlidePicture picture = PresentationSc.Open(Properties.Resources._009, true).
                                                            Slides[1].Shapes.First(sp => sp.Id == 3) as SlidePicture;
            var originLength = (await picture.Image.GetImageBytes().ConfigureAwait(false)).Length;

            // Act
            picture.Image.SetImage(customImageStream);

            // Assert
            var editedLength = (await picture.Image.GetImageBytes().ConfigureAwait(false)).Length;
            editedLength.Should().NotBe(originLength);
        }


        [Fact]
        public void Picture_DoNotParseStrangePicture_Test()
        {
            // TODO: Deeper learn such pictures, where content generated via a:ln
            // Arrange
            var pre = _fixture.Pre019;

            // Act - Assert
            Assert.ThrowsAny<Exception>(() => pre.Slides[1].Shapes.Single(x => x.Id == 47));
        }
    }
}
