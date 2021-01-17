using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class PictureTests : IClassFixture<PptxFixture>
    {
        private readonly PptxFixture _fixture;

        public PictureTests(PptxFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public async void ImageExGetImageBytes_MethodReturnsNonEmptyShapeImage()
        {
            // Arrange
            PictureSc shapePicture1 = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3).Picture;
            PictureSc shapePicture2 = _fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 7).Picture;

            // Act
            byte[] shapePictureContentCase1 = await shapePicture1.ImageEx.GetImageBytes();
            byte[] shapePictureContentCase2 = await shapePicture2.ImageEx.GetImageBytes();

            // Assert
            shapePictureContentCase1.Should().NotBeEmpty();
            shapePictureContentCase2.Should().NotBeEmpty();
        }

        [Fact]
        public async void ImageExSetImage_MethodSetsShapeImage_WhenCustomImageStreamIsPassed()
        {
            // Arrange
            var customImageStream = new MemoryStream(Properties.Resources.test_image_2);
            PictureSc picture = PresentationSc.Open(Properties.Resources._009, true).
                                                            Slides[1].Shapes.First(sp => sp.Id == 3).Picture;
            var originLength = (await picture.ImageEx.GetImageBytes()).Length;

            // Act
            picture.ImageEx.SetImage(customImageStream);

            // Assert
            var editedLength = (await picture.ImageEx.GetImageBytes()).Length;
            editedLength.Should().NotBe(originLength);
        }
    }
}
