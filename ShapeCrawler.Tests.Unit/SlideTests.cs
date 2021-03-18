using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class SlideTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public SlideTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Hide_MethodHidesSlide_WhenItIsExecuted()
        {
            // Arrange
            var pre = SCPresentation.Open(Properties.Resources._001, true);
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
            SlideSc slideEx = _fixture.Pre002.Slides[2];

            // Act
            bool hidden = slideEx.Hidden;

            // Assert
            hidden.Should().BeTrue();
        }


        [Fact]
        public void SaveScheme_CreatesAndSavesSlideSchemeImageInSpecifiedStream()
        {
            // Arrange
            SlideSc slide = _fixture.Pre025.Slides[2];
            var stream = new MemoryStream();

            // Act
            slide.SaveScheme(stream);

            // Assert
            stream.Length.Should().BeGreaterThan(0);
        }

        [Fact]
        public async void BackgroundSetImage_ChangesBackground_WhenImageStreamIsPassed()
        {
            // Arrange
            var pre = SCPresentation.Open(Properties.Resources._009, true);
            var backgroundImage = pre.Slides[0].Background;
            var imgStream = new MemoryStream(Properties.Resources.test_image_2);
            var bytesBefore = await backgroundImage.GetImageBytes().ConfigureAwait(false);

            // Act
            backgroundImage.SetImage(imgStream);
            backgroundImage.SetImage(imgStream);

            // Assert
            var bytesAfter = await backgroundImage.GetImageBytes().ConfigureAwait(false);
            bytesAfter.Length.Should().NotBe(bytesBefore.Length);
        }

        [Fact]
        public void Background_ImageIsNull_WhenTheSlideHasNotBackground()
        {
            // Arrange
            SlideSc slide = _fixture.Pre009.Slides[1];

            // Act
            ImageSc backgroundImage = slide.Background;

            // Assert
            backgroundImage.Should().BeNull();
        }

        [Fact]
        public void CustomData_ReturnsData_WhenCustomDataWasAssigned()
        {
            // Arrange
            const string customDataString = "Test custom data";
            var origPreStream = new MemoryStream();
            origPreStream.Write(Properties.Resources._001);
            var originPre = SCPresentation.Open(origPreStream, true);
            var slide = originPre.Slides.First();

            // Act
            slide.CustomData = customDataString;

            var savedPreStream = new MemoryStream();
            originPre.SaveAs(savedPreStream);
            var savedPre = SCPresentation.Open(savedPreStream, false);
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
            Assert.Single(shapes.Where(sp => sp is IAutoShape));
            Assert.Single(shapes.Where(sp => sp is IPicture));
            Assert.Single(shapes.Where(sp => sp is ITable));
            Assert.Single(shapes.Where(sp => sp is IChart));
            Assert.Single(shapes.Where(sp => sp is IGroupShape));
        }

        [Theory]
        [MemberData(nameof(TestCasesShapesCount))]
        public void ShapesCount_ReturnsNumberOfShapesOnTheSlide(SlideSc slide, int expectedShapesCount)
        {
            // Act
            int shapesCount = slide.Shapes.Count;

            // Assert
            shapesCount.Should().Be(expectedShapesCount);
        }

        public static IEnumerable<object[]> TestCasesShapesCount()
        {
            SCPresentation presentation = SCPresentation.Open(Properties.Resources._009, false);
            
            SlideSc slide = presentation.Slides[0];
            yield return new object[] { slide, 6 };
            
            slide = presentation.Slides[1];
            yield return new object[] { slide, 6 };
            
            slide = SCPresentation.Open(Properties.Resources._002, false).Slides[0];
            yield return new object[] { slide, 4 };
            
            slide = SCPresentation.Open(Properties.Resources._003, false).Slides[0];
            yield return new object[] { slide, 5 };
            
            slide = SCPresentation.Open(Properties.Resources._013, false).Slides[0];
            yield return new object[] { slide, 4 };
            
            slide = SCPresentation.Open(Properties.Resources._023, false).Slides[0];
            yield return new object[] { slide, 1 };

            slide = SCPresentation.Open(Properties.Resources._014, false).Slides[2];
            yield return new object[] { slide, 5 };
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
        public void Shape_IsAPicture()
        {
            // Arrange
            IShape shape = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);

            // Act-Assert
            shape.Should().BeOfType<SlidePicture>();
        }

#if DEBUG
        [Fact(Skip = "The feature is in progress")]
        public void SaveImage_GenerateAndSavesSlideImageInSpecifiedFilePath()
        {
            // Arrange
            SlideSc slide = _fixture.Pre001.Slides[0];

            // Act
            slide.SaveImage(@"c:\1\SlideScSaveImage.png");
        }
#endif
    }
}
