#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
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
        public void Fill_is_not_null()
        {
            // Arrange
            var autoShape = (IAutoShape)_fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 108);

            // Act-Assert
            autoShape.Fill.Should().NotBeNull();
        }

        [Fact]
        public void SetPicture_updates_fill_with_specified_picture_image_When_shape_is_Not_filled()
        {
            // Arrange
            var pptxStream = GetTestStream("008.pptx");
            var imageStream = GetTestStream("test-image-1.png");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");

            // Act
            shape.Fill.SetPicture(imageStream);

            // Assert
            var pictureBytes = shape.Fill.Picture!.GetBytes().Result;
            imageStream.Position = 0;
            var imageBytes = imageStream.ToArray();
            pictureBytes.SequenceEqual(imageBytes).Should().BeTrue();
        }

        [Fact]
        public void Picture_SetImage_updates_picture_fill()
        {
            // Arrange
            var pres = SCPresentation.Open(TestFiles.Presentations.pre009);
            var shape = (IAutoShape)pres.Slides[2].Shapes.First(sp => sp.Id == 4);
            var fill = shape.Fill;
            var newImage = TestFiles.Images.img02_stream;
            var imageSizeBefore = fill.Picture!.GetBytes().GetAwaiter().GetResult().Length;

            // Act
            fill.Picture.SetImage(newImage);

            // Assert
            var imageSizeAfter = shape.Fill.Picture.GetBytes().GetAwaiter().GetResult().Length;
            imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
        }

        [Theory]
        [MemberData(nameof(TestCasesFillType))]
        public void Type_returns_fill_type(IAutoShape shape, FillType expectedFill)
        {
            // Act
            var fillType = shape.Fill.Type;

            // Assert
            fillType.Should().Be(expectedFill);
        }

        public static IEnumerable<object[]> TestCasesFillType()
        {
            var pptxStream = GetTestStream("009_table.pptx");
            var pres = SCPresentation.Open(pptxStream);

            var withNoFill = pres.Slides[1].Shapes.GetById<IAutoShape>(6);
            yield return new object[] { withNoFill, FillType.NoFill };

            var withSolid = pres.Slides[1].Shapes.GetById<IAutoShape>(2);
            yield return new object[] { withSolid, FillType.Solid };

            var withGradient = pres.Slides[1].Shapes.GetByName<IAutoShape>("AutoShape 1");
            yield return new object[] { withGradient, FillType.Gradient };

            var withPicture = pres.Slides[2].Shapes.GetById<IAutoShape>(4);
            yield return new object[] { withPicture, FillType.Picture };

            var withPattern = pres.Slides[1].Shapes.GetByName<IAutoShape>("AutoShape 2");
            yield return new object[] { withPattern, FillType.Pattern };

            pptxStream = GetTestStream("autoshape-case003.pptx");
            pres = SCPresentation.Open(pptxStream);
            var withSlideBg = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
            yield return new object[] { withSlideBg, FillType.SlideBackground };

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
        public void HexSolidColor_getter_returns_color_name()
        {
            // Arrange
            var autoShape = (IAutoShape)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

            // Act
            var shapeSolidColorName = autoShape.Fill.HexSolidColor;

            // Assert
            shapeSolidColorName.Should().BeEquivalentTo("ff0000");
        }

        [Fact]
        public async void Picture_GetImageBytes_returns_image()
        {
            // Arrange
            var shape = (IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 4);

            // Act
            var imageBytes = await shape.Fill.Picture.GetBytes().ConfigureAwait(false);

            // Assert
            imageBytes.Length.Should().BePositive();
        }
    }
}

#endif