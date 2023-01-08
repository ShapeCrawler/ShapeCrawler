using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Helpers.Attributes;
using Xunit;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
public class ShapeCollectionTests : ShapeCrawlerTest
{
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
        var pres = SCPresentation.Open(GetTestStream("003.pptx"));

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
        var pres9 = SCPresentation.Open(GetTestStream("009_table.pptx"));
        IShape shape = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3);

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
        var pptx = GetTestStream("040_video.pptx");
        var pres = SCPresentation.Open(pptx);
        IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);
            
        // Act
        bool isVideo = shape is IVideoShape;

        // Act-Assert
        isVideo.Should().BeTrue();
    }

    [Theory]
    [SlideData("#1", "002.pptx", slideNumber: 1, expectedResult: 4)]
    [SlideData("#2","003.pptx", slideNumber: 1, expectedResult: 5)]
    [SlideData("#3","013.pptx", slideNumber: 1, expectedResult: 4)]
    [SlideData("#4","023.pptx", slideNumber: 1, expectedResult: 1)]
    [SlideData("#5","014.pptx", slideNumber: 3, expectedResult: 5)]
    [SlideData("#6","009_table.pptx", slideNumber: 1, expectedResult: 6)]
    [SlideData("#7","009_table.pptx", slideNumber: 2, expectedResult: 8)]
    public void Count_returns_number_of_shapes(string label, ISlide slide, int expectedCount)
    {
        // Arrange
        var shapeCollection = slide.Shapes;
            
        // Act
        int shapesCount = shapeCollection.Count;

        // Assert
        shapesCount.Should().Be(expectedCount);
    }

    [Fact]
    public void AddAudio_adds_Audio_shape()
    {
        // Arrange
        var preStream = TestFiles.Presentations.pre001_stream;
        var presentation = SCPresentation.Open(preStream);
        var shapes = presentation.Slides[1].Shapes;
        var mp3 = TestFiles.Audio.TestMp3;
        int xPxCoordinate = 300;
        int yPxCoordinate = 100;

        // Act
        shapes.AddAudio(xPxCoordinate, yPxCoordinate, mp3);

        presentation.Save();
        presentation.Close();
        presentation = SCPresentation.Open(preStream);
        IAudioShape addedAudio = presentation.Slides[1].Shapes.OfType<IAudioShape>().Last();

        // Assert
        addedAudio.X.Should().Be(xPxCoordinate);
        addedAudio.Y.Should().Be(yPxCoordinate);
    }

    [Fact]
    public void AddVideo_adds_Video_shape()
    {
        // Arrange
        var preStream = TestFiles.Presentations.pre001_stream;
        var presentation = SCPresentation.Open(preStream);
        var shapesCollection = presentation.Slides[1].Shapes;
        var videoStream = GetTestStream("test-video.mp4");
        int xPxCoordinate = 300;
        int yPxCoordinate = 100;

        // Act
        shapesCollection.AddVideo(xPxCoordinate, yPxCoordinate, videoStream);

        // Assert
        presentation.Save();
        presentation.Close();
        presentation = SCPresentation.Open(preStream);
        var addedVideo = presentation.Slides[1].Shapes.OfType<IVideoShape>().Last();
        addedVideo.X.Should().Be(xPxCoordinate);
        addedVideo.Y.Should().Be(yPxCoordinate);
    }
    
    [Fact]
    public void AutoShapes_AddRectangle_adds_rectangle_with_valid_id_and_name()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case011_save-as-png.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapes = pres.Slides[0].Shapes;
            
        // Act
        var autoShape = shapes.AutoShapes.AddRectangle( 50, 60, 100, 70);

        // Assert
        autoShape.Name.Should().Be("AutoShape 4");
        autoShape.Id.Should().Be(7);
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void AutoShapes_AddRectangle_adds_Rectangle_in_the_New_Presentation()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var shapes = pres.Slides[0].Shapes;
            
        // Act
        var rectangle = shapes.AutoShapes.AddRectangle(50, 60, 100, 70);

        // Assert
        rectangle.GeometryType.Should().Be(SCGeometry.Rectangle);
        rectangle.X.Should().Be(50);
        rectangle.Y.Should().Be(60);
        rectangle.Width.Should().Be(100);
        rectangle.Height.Should().Be(70);
        rectangle.TextFrame!.Paragraphs.Count.Should().Be(1);
        rectangle.Outline.Color.Should().Be("000000");
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }
    
    [Fact]
    public void AutoShapes_AddRoundedRectangle_adds_Rounded_Rectangle()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapes = pres.Slides[0].Shapes;
            
        // Act
        var roundedRectangle = shapes.AutoShapes.AddRoundedRectangle(50, 60, 100, 70);

        // Assert
        roundedRectangle.GeometryType.Should().Be(SCGeometry.RoundRectangle);
        roundedRectangle.Name.Should().Be("AutoShape 8");
        roundedRectangle.Outline.Color.Should().Be("000000");
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void AddTable_adds_table()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var shapes = pres.Slides[0].Shapes;
        
        // Act
        var table = shapes.AddTable(x: 50, y: 60, columns: 3, rows: 2);

        // Assert
        table.Columns.Should().HaveCount(3);
        table.Rows.Should().HaveCount(2);
        table.Id.Should().Be(1);
        table.Name.Should().Be("Table 1");
        table.Columns[0].Width.Should().Be(284);
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void Remove_removes_shape()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapeCollection = pres.Slides[0].Shapes;
        var shape = shapeCollection.GetByName("TextBox 3")!;

        // Act
        shapeCollection.Remove(shape);

        // Assert
        shape = shapeCollection.GetByName("TextBox 3");
        shape.Should().BeNull();
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }
}