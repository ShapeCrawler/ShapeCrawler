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
    public class ShapeCollectionTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture fixture;

        public ShapeCollectionTests(PresentationFixture fixture)
        {
            this.fixture = fixture;
        }

        [Fact]
        public void Shapes_contains_particular_shape_Types()
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
        public void Shapes_contains_Picture_shape()
        {
            // Arrange
            IShape shape = this.fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);

            // Act-Assert
            IPicture picture = shape as IPicture;
            picture.Should().NotBeNull();
        }

        [Fact]
        public void Shapes_contains_Audio_shape()
        {
            // Arrange
            var pptxStream = GetTestStream("audio-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);

            // Act
            bool isAudio = shape is IAudioShape;

            // Assert
            isAudio.Should().BeTrue();
        }
        
        [Fact]
        public void Shapes_contains_Connection_shape()
        {
            var pptxStream = GetTestStream("001.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var shapesCollection = presentation.Slides[0].Shapes;

            // Act-Assert
            Assert.Contains(shapesCollection, shape => shape.Id == 10 && shape is IConnectionShape && shape.GeometryType == SCGeometry.Line);
        }
        
        [Fact]
        public void Shapes_contains_Video_shape()
        {
            // Arrange
            IShape shape = this.fixture.Pre040.Slides[0].Shapes.First(sp => sp.Id == 8);
            
            // Act
            bool isVideo = shape is IVideoShape;

            // Act-Assert
            isVideo.Should().BeTrue();
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
        
        public static IEnumerable<object[]> TestCasesShapesCount()
        {
            var pres = SCPresentation.Open(Properties.Resources._009);
            
            var slide = pres.Slides[0];
            yield return new object[] { slide, 6 };
            
            slide = pres.Slides[1];
            yield return new object[] { slide, 8 };
            
            slide = SCPresentation.Open(Properties.Resources._002).Slides[0];
            yield return new object[] { slide, 4 };
            
            slide = SCPresentation.Open(Properties.Resources._003).Slides[0];
            yield return new object[] { slide, 5 };
            
            slide = SCPresentation.Open(Properties.Resources._013).Slides[0];
            yield return new object[] { slide, 4 };
            
            slide = SCPresentation.Open(Properties.Resources._023).Slides[0];
            yield return new object[] { slide, 1 };

            slide = SCPresentation.Open(Properties.Resources._014).Slides[2];
            yield return new object[] { slide, 5 };
        }

        [Fact]
        public void Shapes_AddNewAudio_adds_Audio_shape()
        {
            // Arrange
            Stream preStream = TestFiles.Presentations.pre001_stream;
            IPresentation presentation = SCPresentation.Open(preStream);
            IShapeCollection shapes = presentation.Slides[1].Shapes;
            Stream mp3 = TestFiles.Audio.TestMp3;
            int xPxCoordinate = 300;
            int yPxCoordinate = 100;

            // Act
            shapes.AddNewAudio(xPxCoordinate, yPxCoordinate, mp3);

            presentation.Save();
            presentation.Close();
            presentation = SCPresentation.Open(preStream);
            IAudioShape addedAudio = presentation.Slides[1].Shapes.OfType<IAudioShape>().Last();

            // Assert
            addedAudio.X.Should().Be(xPxCoordinate);
            addedAudio.Y.Should().Be(yPxCoordinate);
        }

        [Fact]
        public void Shapes_AddNewVideo_adds_Video_shape()
        {
            // Arrange
            var preStream = TestFiles.Presentations.pre001_stream;
            var presentation = SCPresentation.Open(preStream);
            var shapesCollection = presentation.Slides[1].Shapes;
            var videoStream = GetTestStream("test-video.mp4");
            int xPxCoordinate = 300;
            int yPxCoordinate = 100;

            // Act
            shapesCollection.AddNewVideo(xPxCoordinate, yPxCoordinate, videoStream);

            // Assert
            presentation.Save();
            presentation.Close();
            presentation = SCPresentation.Open(preStream);
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
