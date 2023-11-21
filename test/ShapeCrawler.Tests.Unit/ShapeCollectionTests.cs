using System.Diagnostics.CodeAnalysis;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;
using Assert = Xunit.Assert;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ShapeCollectionTests : SCTest
{
    [Test]
    public void Add_adds_shape()
    {
        // Arrange
        var pres = new Presentation(StreamOf("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.GetByName("TextBox")!;
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        shapes.GetByName("TextBox 2").Should().NotBeNull();
    }

    [Test]
    public void Add_adds_table()
    {
        // Arrange
        var pres = new Presentation(StreamOf("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.GetByName("Table 1")!;
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        var addedShape = shapes.Last();
        addedShape.Should().BeAssignableTo<ITable>();
    }

    [Test]
    public void Contains_particular_shape_Types()
    {
        // Arrange
        var pres = new Presentation(StreamOf("003.pptx"));

        // Act
        var shapes = pres.Slides.First().Shapes;

        // Assert
        Assert.Single(shapes.Where(sp => sp.ShapeType == ShapeType.Chart));
        Assert.Single(shapes.Where(sp => sp is IPicture));
        Assert.Single(shapes.Where(sp => sp is ITable));
        Assert.Single(shapes.Where(sp => sp is IGroupShape));
    }

    [Test]
    public void Contains_Picture_shape()
    {
        // Arrange
        IShape shape = new Presentation(StreamOf("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3);

        // Act-Assert
        IPicture picture = shape as IPicture;
        picture.Should().NotBeNull();
    }

    [Test]
    public void Contains_Media_Shape()
    {
        // Arrange
        var pptxStream = StreamOf("audio-case001.pptx");
        var pres = new Presentation(pptxStream);
        IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);

        // Act
        bool isMediaShape = shape is IMediaShape;

        // Assert
        isMediaShape.Should().BeTrue();
    }

    [Test]
    public void Contains_Connection_shape()
    {
        var pptxStream = StreamOf("001.pptx");
        var presentation = new Presentation(pptxStream);
        var shapesCollection = presentation.Slides[0].Shapes;

        // Act-Assert
        Assert.Contains(shapesCollection,
            shape => shape.Id == 10 && shape is ILine && shape.GeometryType == Geometry.Line);
    }

    [Test]
    public void Contains_Video_shape()
    {
        // Arrange
        var pptx = StreamOf("040_video.pptx");
        var pres = new Presentation(pptx);
        IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);

        // Act
        bool isVideo = shape is IMediaShape;

        // Act-Assert
        isVideo.Should().BeTrue();
    }

    [Test]
    public void AddLine_adds_a_new_Line_shape_from_raw_open_xml_content()
    {
        // Arrange
        var pres = new Presentation();
        var xml = TestHelperShared.GetString("line-shape.xml");
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(xml);

        // Assert
        var addedLine = shapes.Last();
        addedLine.Id.Should().Be(1);
        shapes.Count.Should().Be(1);
    }

    [Test]
    public void AddLine_adds_line_Right_Up()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 10, startPointY: 10, endPointX: 20, endPointY: 5);

        // Assert
        var addedLine = (ILine)shapes.Last();
        shapes.Should().ContainSingle();
        addedLine.ShapeType.Should().Be(ShapeType.Line);
        addedLine.StartPoint.X.Should().Be(10);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(20);
        addedLine.EndPoint.Y.Should().Be(5);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Up_Up()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 10, startPointY: 10, endPointX: 10, endPointY: 5);

        // Assert
        var addedLine = (ILine)shapes.Last();
        addedLine.StartPoint.X.Should().Be(10);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(10);
        addedLine.EndPoint.Y.Should().Be(5);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Left_Up()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 100, startPointY: 50, endPointX: 40, endPointY: 20);

        // Assert
        var addedLine = (ILine)shapes.Last();
        addedLine.StartPoint.X.Should().Be(100);
        addedLine.StartPoint.Y.Should().Be(50);
        addedLine.EndPoint.X.Should().Be(40);
        addedLine.EndPoint.Y.Should().Be(20);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Left_Down()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 10, endPointX: 40, endPointY: 20);

        // Assert
        var addedLine = (ILine)shapes.Last();
        addedLine.StartPoint.X.Should().Be(50);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(40);
        addedLine.EndPoint.Y.Should().Be(20);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Right_Right()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Assert
        var line = (ILine)shapes.Last();
        line.StartPoint.X.Should().Be(50);
        line.StartPoint.Y.Should().Be(60);
        line.EndPoint.X.Should().Be(100);
        line.EndPoint.Y.Should().Be(60);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Assert
        shapes.Should().ContainSingle();
        var line = (ILine)shapes.Last();
        line.ShapeType.Should().Be(ShapeType.Line);
        line.X.Should().Be(50);
        line.Y.Should().Be(60);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Left_Left()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 100, startPointY: 50, endPointX: 80, endPointY: 50);

        // Assert
        var line = (ILine)shapes.Last();
        line.StartPoint.X.Should().Be(100);
        line.StartPoint.Y.Should().Be(50);
        line.EndPoint.X.Should().Be(80);
        line.EndPoint.Y.Should().Be(50);
        pres.Validate();
    }

    [Test]
    public void AddAudio_adds_audio_shape_with_MP3_content()
    {
        // Arrange
        var pptx = StreamOf("001.pptx");
        var mp3 = StreamOf("test-mp3.mp3");
        var pres = new Presentation(pptx);
        var shapes = pres.Slides[1].Shapes;
        int xPxCoordinate = 300;
        int yPxCoordinate = 100;

        // Act
        shapes.AddAudio(xPxCoordinate, yPxCoordinate, mp3);

        pres.Save();
        pres = new Presentation(pptx);
        var addedAudio = pres.Slides[1].Shapes.OfType<IMediaShape>().Last();

        // Assert
        addedAudio.X.Should().Be(xPxCoordinate);
        addedAudio.Y.Should().Be(yPxCoordinate);
    }

    [Test]
    public void AddAudio_adds_audio_shape_with_WAVE_content()
    {
        // Arrange
        var wav = StreamOf("test-wav.wav");
        var pres = new Presentation(StreamOf("001.pptx"));
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.AddAudio(300, 100, wav, AudioType.WAVE);

        // Assert
        var addedAudio = pres.Slides[1].Shapes.OfType<IMediaShape>().Last();
        addedAudio.X.Should().Be(300);
    }

    [Test]
    public void AddVideo_adds_Video_shape()
    {
        // Arrange
        var preStream = StreamOf("001.pptx");
        var presentation = new Presentation(preStream);
        var shapesCollection = presentation.Slides[1].Shapes;
        var videoStream = StreamOf("test-video.mp4");
        int xPxCoordinate = 300;
        int yPxCoordinate = 100;

        // Act
        shapesCollection.AddVideo(xPxCoordinate, yPxCoordinate, videoStream);

        // Assert
        presentation.Save();
        presentation = new Presentation(preStream);
        var addedVideo = presentation.Slides[1].Shapes.OfType<IMediaShape>().Last();
        addedVideo.X.Should().Be(xPxCoordinate);
        addedVideo.Y.Should().Be(yPxCoordinate);
    }

    [Test, Ignore("Not implemented yet")]
    public void AddBarChart_adds_Bar_Chart()
    {
        // Arrange
        var pres = new Presentation();

        // Act
        pres.Slides[0].Shapes.AddBarChart(BarChartType.ClusteredBar);

        // Assert
        var barChart = pres.Slides[0].Shapes.Last();
        barChart.Should().NotBeNull();
        pres.Validate();
    }

    [Test]
    public void AddPicture_adds_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestHelper.GetStream("test-image-1.png");

        // Act
        shapes.AddPicture(image);

        // Assert
        shapes.Should().HaveCount(1);
        var picture = (IPicture)shapes.Last();
        picture.ShapeType.Should().Be(ShapeType.Picture);
        pres.Validate();
    }

    [Test]
    public void AddRectangle_adds_rectangle_with_valid_id_and_name()
    {
        // Arrange
        var pres = new Presentation(StreamOf("autoshape-case011_save-as-png.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddRectangle(50, 60, 100, 70);

        // Assert
        var autoShape = shapes.Last();
        autoShape.Name.Should().Be("Rectangle");
        autoShape.Id.Should().Be(7);
        pres.Validate();
    }

    [Test]
    public void AddRectangle_adds_Rectangle_in_the_New_Presentation()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddRectangle(50, 60, 100, 70);

        // Assert
        var rectangle = shapes.Last();
        rectangle.GeometryType.Should().Be(Geometry.Rectangle);
        rectangle.X.Should().Be(50);
        rectangle.Y.Should().Be(60);
        rectangle.Width.Should().Be(100);
        rectangle.Height.Should().Be(70);
        rectangle.TextFrame!.Paragraphs.Count.Should().Be(1);
        rectangle.Outline.HexColor.Should().Be("000000");
        pres.Validate();
    }

    [Test]
    public void AddRoundedRectangle_adds_Rounded_Rectangle()
    {
        // Arrange
        var pres = new Presentation(StreamOf("autoshape-grouping.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddRoundedRectangle(50, 60, 100, 70);

        // Assert
        var roundedRectangle = shapes.Last();
        roundedRectangle.GeometryType.Should().Be(Geometry.RoundRectangle);
        roundedRectangle.Name.Should().Be("Rectangle: Rounded Corners");
        roundedRectangle.Outline.HexColor.Should().Be("000000");
        pres.Validate();
    }

    [Test]
    public void AddTable_adds_table()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddTable(x: 50, y: 60, columnsCount: 3, rowsCount: 2);

        // Assert
        var table = (ITable)shapes.Last();
        table.Columns.Should().HaveCount(3);
        table.Rows.Should().HaveCount(2);
        table.Id.Should().Be(1);
        table.Name.Should().Be("Table 1");
        table.Columns[0].Width.Should().Be(284);
        pres.Validate();
    }
}