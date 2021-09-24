using System;
using System.IO;
using System.Linq;
using FluentAssertions;
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
        public async void ImageGetBytes_ReturnsImageByteArray()
        {
            // Arrange
            IPicture shapePicture1 = (IPicture)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);
            IPicture shapePicture2 = (IPicture)_fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 7);

            // Act
            byte[] shapePictureContentCase1 = await shapePicture1.Image.GetBytes().ConfigureAwait(false);
            byte[] shapePictureContentCase2 = await shapePicture2.Image.GetBytes().ConfigureAwait(false);

            // Assert
            shapePictureContentCase1.Should().NotBeEmpty();
            shapePictureContentCase2.Should().NotBeEmpty();
        }

        [Fact]
        public void ImageSetImage_UpdatesPictureImage()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Properties.Resources._009, true);
            MemoryStream imageStream = new (TestFiles.Images.imageByteArray02);
            MemoryStream preStream = new();
            IPicture picture = (IPicture) presentation.Slides[1].Shapes.First(sp => sp.Id == 3);
            int lengthBefore = picture.Image.GetBytes().Result.Length;

            // Act
            picture.Image.SetImage(imageStream);

            // Assert
            presentation.SaveAs(preStream);
            //presentation.Close();
            presentation = SCPresentation.Open(preStream, false);
            picture = (IPicture)presentation.Slides[1].Shapes.First(sp => sp.Id == 3);
            int lengthAfter = picture.Image.GetBytes().Result.Length;

            lengthAfter.Should().NotBe(lengthBefore);
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
