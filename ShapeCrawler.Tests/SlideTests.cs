using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using Xunit;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class SlideTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture fixture;

        public SlideTests(PresentationFixture fixture)
        {
            this.fixture = fixture;
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
            ISlide slideEx = this.fixture.Pre002.Slides[2];

            // Act
            bool hidden = slideEx.Hidden;

            // Assert
            hidden.Should().BeTrue();
        }


        [Fact]
        public void SaveScheme_CreatesAndSavesSlideSchemeImageInSpecifiedStream()
        {
            // Arrange
            ISlide slide = this.fixture.Pre025.Slides[2];
            var stream = new MemoryStream();

            // Act
            slide.SaveScheme(stream);

            // Assert
            stream.Length.Should().BeGreaterThan(0);
        }

        [Fact]
        public async void Background_SetImage_updates_background()
        {
            // Arrange
            var pre = SCPresentation.Open(Properties.Resources._009, true);
            var backgroundImage = pre.Slides[0].Background;
            var imgStream = new MemoryStream(Properties.Resources.test_image_2);
            var bytesBefore = await backgroundImage.GetBytes().ConfigureAwait(false);

            // Act
            backgroundImage.SetImage(imgStream);

            // Assert
            var bytesAfter = await backgroundImage.GetBytes().ConfigureAwait(false);
            bytesAfter.Length.Should().NotBe(bytesBefore.Length);
        }

        [Fact]
        public void Background_ImageIsNull_WhenTheSlideHasNotBackground()
        {
            // Arrange
            ISlide slide = this.fixture.Pre009.Slides[1];

            // Act
            SCImage backgroundImage = slide.Background;

            // Assert
            backgroundImage.Should().BeNull();
        }

        [Fact]
        public void CustomData_ReturnsData_WhenCustomDataWasAssigned()
        {
            // Arrange
            const string customDataString = "Test custom data";
            var originPre = SCPresentation.Open(Properties.Resources._001, true);
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
        public void Shapes_collection_contains_particular_shape_Types()
        {
            // Arrange
            IPresentation pre = this.fixture.Pre003;

            // Act
            IShapeCollection shapes = pre.Slides.First().Shapes;

            // Assert
            Assert.Single(shapes.Where(sp => sp is IAutoShape));
            Assert.Single(shapes.Where(sp => sp is IPicture));
            Assert.Single(shapes.Where(sp => sp is ITable));
            Assert.Single(shapes.Where(sp => sp is IChart));
            Assert.Single(shapes.Where(sp => sp is IGroupShape));
        }

        [Fact]
        public void Shapes_collection_contains_Picture_shape()
        {
            // Arrange
            IShape shape = this.fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);

            // Act-Assert
            IPicture picture = shape as IPicture;
            picture.Should().NotBeNull();
        }

        [Fact]
        public void Shapes_collection_contains_Audio_shape()
        {
            // Arrange
            var pptxStream = GetTestFileStream("audio-case001.pptx");
            var pres = SCPresentation.Open(pptxStream, false);
            IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);

            // Act
            bool isAudio = shape is IAudioShape;

            // Assert
            isAudio.Should().BeTrue();
        }
        
        [Fact]
        public void Shapes_collection_contains_Connection_shape()
        {
            var pptxStream = GetTestFileStream("001.pptx");
            var presentation = SCPresentation.Open(pptxStream, false);
            var shapesCollection = presentation.Slides[0].Shapes;

            // Act-Assert
            Assert.Contains(shapesCollection, shape => shape.Id == 10 && shape is IConnectionShape && shape.GeometryType == GeometryType.Line);
        }

        [Theory]
        [MemberData(nameof(TestCasesShapesCount))]
        public void Shapes_Count_returns_number_of_shapes(ISlide slide, int expectedShapesCount)
        {
            // Act
            int shapesCount = slide.Shapes.Count;

            // Assert
            shapesCount.Should().Be(expectedShapesCount);
        }

        [Fact]
        public void Shapes_AddNewAudio_adds_Audio_shape()
        {
            // Arrange
            Stream preStream = TestFiles.Presentations.pre001_stream;
            IPresentation presentation = SCPresentation.Open(preStream, true);
            IShapeCollection shapes = presentation.Slides[1].Shapes;
            Stream mp3 = TestFiles.Audio.TestMp3;
            int xPxCoordinate = 300;
            int yPxCoordinate = 100;

            // Act
            shapes.AddNewAudio(xPxCoordinate, yPxCoordinate, mp3);

            presentation.Save();
            presentation.Close();
            presentation = SCPresentation.Open(preStream, false);
            IAudioShape addedAudio = presentation.Slides[1].Shapes.OfType<IAudioShape>().Last();

            // Assert
            addedAudio.X.Should().Be(xPxCoordinate);
            addedAudio.Y.Should().Be(yPxCoordinate);
        }

        public static IEnumerable<object[]> TestCasesShapesCount()
        {
            IPresentation presentation = SCPresentation.Open(Properties.Resources._009, false);
            
            ISlide slide = presentation.Slides[0];
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
            var slide = this.fixture.Pre001.Slides.First();

            // Act
            var sldCustomData = slide.CustomData;

            // Assert
            sldCustomData.Should().BeNull();
        }

        [Fact]
        public void NumberSetter_MovesSlide()
        {
            // Arrange
            Stream preStream = TestFiles.Presentations.pre001_stream;
            IPresentation presentation = SCPresentation.Open(preStream, true);
            ISlide slide1 = presentation.Slides[0];
            slide1.CustomData = "old-number-1";
            ISlide slide2 = presentation.Slides[1];

            // Act
            slide1.Number = 2;

            // Assert
            slide1.Number.Should().Be(2);
            slide2.Number.Should().Be(1, "because the first slide was inserted to its position.");
            
            presentation.Close();
            presentation = SCPresentation.Open(preStream, false);
            slide2 = presentation.Slides.First(s => s.CustomData == "old-number-1");
            slide2.Number.Should().Be(2);
        }

        [Fact]
        public void Shapes_collection_contains_Video_shape()
        {
            // Arrange
            IShape shape = this.fixture.Pre040.Slides[0].Shapes.First(sp => sp.Id == 8);
            
            // Act
            bool isVideo = shape is IVideoShape;

            // Act-Assert
            isVideo.Should().BeTrue();
        }

        [Fact]
        public void Shapes_AddNewVideo_adds_Video_shape()
        {
            // Arrange
            var preStream = TestFiles.Presentations.pre001_stream;
            var presentation = SCPresentation.Open(preStream, true);
            var shapesCollection = presentation.Slides[1].Shapes;
            var videoStream = GetTestFileStream("test-video.mp4");
            int xPxCoordinate = 300;
            int yPxCoordinate = 100;

            // Act
            shapesCollection.AddNewVideo(xPxCoordinate, yPxCoordinate, videoStream);

            // Assert
            presentation.Save();
            presentation.Close();
            presentation = SCPresentation.Open(preStream, false);
            var addedVideo = presentation.Slides[1].Shapes.OfType<IVideoShape>().Last();
            addedVideo.X.Should().Be(xPxCoordinate);
            addedVideo.Y.Should().Be(yPxCoordinate);
        }

#if TEST
        [Fact]
        public void ToHtml_converts_slide_to_HTML()
        {
            // Arrange
            var slide = this.GetSlide("052_slide-to-html.pptx", 1);

            // Act
            var html = slide.ToHtml().Result;
            File.WriteAllText(@"C:\Documents\ShapeCrawler\Issues\SC-189_convert-slide-to-html\to-html\output.html", html);

            // Arrange
            
        }
#endif
    }
}
