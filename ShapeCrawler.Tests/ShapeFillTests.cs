#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tests.Helpers;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests
{
    public class ShapeFillTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ShapeFillTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void AutoShape_Fill_is_Not_Null_When_autoShape_is_filled()
        {
            // Arrange
            IAutoShape autoShape = (IAutoShape)_fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 108);

            // Act-Assert
            autoShape.Fill.Should().NotBeNull();
        }

        [Fact]
        public void AutoShape_Fill_SetPicture_updates_fill_with_specified_picture_When_shape_is_Not_filled()
        {
            // Arrange
            var pptxStream = GetTestFileStream("008.pptx");
            var imageStream = GetTestFileStream("test-image-1.png");
            var pres = SCPresentation.Open(pptxStream, true);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");

            // Act
            shape.Fill.SetPicture(imageStream);

            // Assert
            var pictureBytes = shape.Fill.Picture!.GetBytes().Result;
            imageStream.Position = 0;
            var imageBytes = imageStream.ToArray();
            pictureBytes.SequenceEqual(imageBytes).Should().BeTrue();
        }

        [Theory]
        [MemberData(nameof(TestCasesFillType))]
        public void AutoShape_Fill_Type_returns_fill_type(IAutoShape shape, FillType expectedFill)
        {
            // Act
            var fillType = shape.Fill.Type;

            // Assert
            fillType.Should().Be(expectedFill);
        }

        public static IEnumerable<object[]> TestCasesFillType()
        {
            var pptxStream = GetTestFileStream("009_table.pptx");
            var pres = SCPresentation.Open(pptxStream, false);
            var autoShape = pres.Slides[2].Shapes.GetById<IAutoShape>(4);
            yield return new object[] { autoShape, FillType.Picture };

            autoShape = pres.Slides[1].Shapes.GetById<IAutoShape>(2);
            yield return new object[] { autoShape, FillType.Solid };

            autoShape = pres.Slides[1].Shapes.GetById<IAutoShape>(6);
            yield return new object[] { autoShape, FillType.NoFill };
        }

        [Fact]
        public void AutoShape_Fill_Type_returns_NoFill_When_shape_is_Not_filled()
        {
            // Arrange
            var autoShape = (IAutoShape)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 6);

            // Act
            var fillType = autoShape.Fill.Type;

            // Assert
            fillType.Should().Be(FillType.NoFill);
        }

        [Fact]
        public void AutoShape_Fill_SolidColor_Name_getter_returns_color_name()
        {
            // Arrange
            IAutoShape autoShape = (IAutoShape)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

            // Act
            var shapeSolidColorName = autoShape.Fill.SolidColor.Name;

            // Assert
            shapeSolidColorName.Should().BeEquivalentTo("ff0000");
        }

        [Fact]
        public async void AutoShape_Fill_Picture_GetImageBytes_ReturnsImageByWhichTheAutoShapeIsFilled()
        {
            // Arrange
            IAutoShape shape = (IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 4);

            // Act
            byte[] imageBytes = await shape.Fill.Picture.GetBytes().ConfigureAwait(false);

            // Assert
            imageBytes.Length.Should().BePositive();
        }

        [Fact]
        public void AutoShape_Fill_Picture_SetImage_updates_picture()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(TestFiles.Presentations.pre009, true);
            IAutoShape autoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 4);
            MemoryStream newImage = TestFiles.Images.img02_stream;
            int imageSizeBefore = autoShape.Fill.Picture.GetBytes().GetAwaiter().GetResult().Length;

            // Act
            autoShape.Fill.Picture.SetImage(newImage);

            // Assert
            int imageSizeAfter = autoShape.Fill.Picture.GetBytes().GetAwaiter().GetResult().Length;
            imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
        }
    }
}

#endif