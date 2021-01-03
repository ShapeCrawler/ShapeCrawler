using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class PictureTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public PictureTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public async void ImageExGetImageBytes_MethodReturnsNonEmptyShapeImage()
        {
            // Arrange
            Picture picture = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3).Picture;

            // Act
            byte[] bytes = await picture.ImageEx.GetImageBytes();

            // Assert
            bytes.Should().NotBeNullOrEmpty();
        }

        [Fact]
        public async void ImageExSetImage_MethodSetsShapeImage_WhenCustomImageStreamIsPassed()
        {
            // Arrange
            var customImageStream = new MemoryStream(Properties.Resources.test_image_2);
            Picture picture = Presentation.Open(Properties.Resources._009, true).
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
