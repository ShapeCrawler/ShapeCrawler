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

        [Theory]
        [MemberData(nameof(TestCasesSetPicture))]
        public void SetPicture_updates_fill_with_specified_picture_image_When_shape_is_Not_filled(TestCase testCase)
        {
            // Arrange
            var fill = testCase.AutoShape.Fill;
            var imageStream = GetTestStream("test-image-1.png");

            // Act
            fill.SetPicture(imageStream);

            // Assert
            var pictureBytes = fill.Picture!.BinaryData.Result;
            var imageBytes = imageStream.ToArray();
            pictureBytes.SequenceEqual(imageBytes).Should().BeTrue();
        }

        public static IEnumerable<object[]> TestCasesSetPicture
        {
            get
            {
                var testCase1 = new TestCase("#1")
                {
                    PresentationName = "008.pptx",
                    SlideNumber = 1,
                    ShapeName = "AutoShape 1"
                };
                yield return new object[] { testCase1 };
                
                var testCase2 = new TestCase("#2")
                {
                    PresentationName = "autoshape-case009.pptx",
                    SlideNumber = 1,
                    ShapeName = "AutoShape 1"
                };
                yield return new object[] { testCase2 };
            }
        }

        [Fact]
        public void Picture_SetImage_updates_picture_fill()
        {
            // Arrange
            var pres = SCPresentation.Open(TestFiles.Presentations.pre009);
            var shape = (IAutoShape)pres.Slides[2].Shapes.First(sp => sp.Id == 4);
            var fill = shape.Fill;
            var newImage = TestFiles.Images.img02_stream;
            var imageSizeBefore = fill.Picture!.BinaryData.GetAwaiter().GetResult().Length;

            // Act
            fill.Picture.SetImage(newImage);

            // Assert
            var imageSizeAfter = shape.Fill.Picture.BinaryData.GetAwaiter().GetResult().Length;
            imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
        }

        [Theory]
        [MemberData(nameof(TestCasesFillType))]
        public void Type_returns_fill_type(IAutoShape shape, SCFillType expectedFill)
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
            yield return new object[] { withNoFill, SCFillType.NoFill };

            var withSolid = pres.Slides[1].Shapes.GetById<IAutoShape>(2);
            yield return new object[] { withSolid, SCFillType.Solid };

            var withGradient = pres.Slides[1].Shapes.GetByName<IAutoShape>("AutoShape 1");
            yield return new object[] { withGradient, SCFillType.Gradient };

            var withPicture = pres.Slides[2].Shapes.GetById<IAutoShape>(4);
            yield return new object[] { withPicture, SCFillType.Picture };

            var withPattern = pres.Slides[1].Shapes.GetByName<IAutoShape>("AutoShape 2");
            yield return new object[] { withPattern, SCFillType.Pattern };

            pptxStream = GetTestStream("autoshape-case003.pptx");
            pres = SCPresentation.Open(pptxStream);
            var withSlideBg = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
            yield return new object[] { withSlideBg, SCFillType.SlideBackground };
        }

        [Fact]
        public void AutoShape_Fill_Type_returns_NoFill_When_shape_is_Not_filled()
        {
            // Arrange
            var autoShape = (IAutoShape)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 6);

            // Act
            var fillType = autoShape.Fill.Type;

            // Assert
            fillType.Should().Be(SCFillType.NoFill);
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
            var imageBytes = await shape.Fill.Picture.BinaryData.ConfigureAwait(false);

            // Assert
            imageBytes.Length.Should().BePositive();
        }
    }
}

#endif