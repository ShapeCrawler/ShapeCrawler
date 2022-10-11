using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Helpers.Attributes;
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

        [Theory]
        [LayoutShapeData("autoshape-case004_subtitle.pptx", slideNumber: 1, shapeName: "Group 1")]
        [MasterShapeData("autoshape-case004_subtitle.pptx", shapeName: "Group 1")]
        public void GetByName_returns_shape_by_specified_name(IShape shape)
        {
            // Arrange
            var groupShape = (IGroupShape)shape;
            var shapeCollection = groupShape.Shapes;
            
            // Act
            var resultShape = shapeCollection.GetByName<IAutoShape>("AutoShape 1");

            // Assert
            resultShape.Should().NotBeNull();
        }
        
        [Fact]
        public void Contains_particular_shape_Types()
        {
            // Arrange
            var pres = this.fixture.Pre003;

            // Act
            var shapes = pres.Slides.First().Shapes;

            // Assert
            Assert.Single(shapes.Where(sp => sp is IAutoShape));
            Assert.Single(shapes.Where(sp => sp is IPicture));
            Assert.Single(shapes.Where(sp => sp is ITable));
            Assert.Single(shapes.Where(sp => sp is IChart));
            Assert.Single(shapes.Where(sp => sp is IGroupShape));
        }

        [Fact]
        public void Contains_Picture_shape()
        {
            // Arrange
            IShape shape = this.fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);

            // Act-Assert
            IPicture picture = shape as IPicture;
            picture.Should().NotBeNull();
        }

        [Fact]
        public void Contains_Audio_shape()
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
        public void Contains_Connection_shape()
        {
            var pptxStream = GetTestStream("001.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var shapesCollection = presentation.Slides[0].Shapes;

            // Act-Assert
            Assert.Contains(shapesCollection, shape => shape.Id == 10 && shape is IConnectionShape && shape.GeometryType == SCGeometry.Line);
        }
        
        [Fact]
        public void Contains_Video_shape()
        {
            // Arrange
            IShape shape = this.fixture.Pre040.Slides[0].Shapes.First(sp => sp.Id == 8);
            
            // Act
            bool isVideo = shape is IVideoShape;

            // Act-Assert
            isVideo.Should().BeTrue();
        }

        [Theory]
        [SlideData("002.pptx", slideNumber: 1, expectedResult: 4)]
        [SlideData("003.pptx", slideNumber: 1, expectedResult: 1)]
        [SlideData("013.pptx", slideNumber: 1, expectedResult: 4)]
        [SlideData("023.pptx", slideNumber: 1, expectedResult: 1)]
        [SlideData("014.pptx", slideNumber: 3, expectedResult: 5)]
        [SlideData("009_table.pptx", slideNumber: 1, expectedResult: 6)]
        [SlideData("009_table.pptx", slideNumber: 2, expectedResult: 8)]
        [SlideData("009_table.pptx", slideNumber: 2, expectedResult: 8)]
        public void Count_returns_number_of_shapes(ISlide slide, int expectedShapesCount)
        {
            // Arrange
            var shapeCollection = slide.Shapes;
            
            // Act
            int shapesCount = shapeCollection.Count;

            // Assert
            shapesCount.Should().Be(expectedShapesCount);
        }

        [Fact]
        public void AddNewAudio_adds_Audio_shape()
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
        public void AddNewVideo_adds_Video_shape()
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
    }
}
