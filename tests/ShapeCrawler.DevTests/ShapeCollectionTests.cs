using System.Diagnostics.CodeAnalysis;
using Fixture;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;


// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.DevTests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
public class ShapeCollectionTests : SCTest
{
    private readonly Fixtures fixture = new();

    [Test]
    public void Add_adds_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.Shape("TextBox")!;
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        shapes.Shape("TextBox 2").Should().NotBeNull();
    }
    
    [Test]
    public void AddDateAndTime_adds_Date_and_time_placeholder_shape()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;

        // Act
        shapes.AddDateAndTime();

        // Assert
        shapes.Should().Contain(shape => shape.PlaceholderType == PlaceholderType.DateAndTime);
        ValidatePresentation(pres);
    }
    
    [Test]
    public void AddFooter_adds_Footer_placeholder_shape()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;

        // Act
        shapes.AddFooter();

        // Assert
        shapes.Should().Contain(shape => shape.PlaceholderType == PlaceholderType.Footer);
        ValidatePresentation(pres);
    }
    
    [Test]
    public void AddSlideNumber_adds_Slide_number_placeholder_shape()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;

        // Act
        shapes.AddSlideNumber();

        // Assert
        shapes.Should().Contain(shape => shape.PlaceholderType == PlaceholderType.SlideNumber);
        ValidatePresentation(pres);
    }

    [Test]
    public void Add_adds_table()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slide(1).Shape("Table 1");
        var shapes = pres.Slide(2).Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        var addedShape = shapes.Last();
        addedShape.Table.Should().NotBeNull();
    }

    [Test]
    public void AddPicture_adds_picture_with_gif_image()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var gif = fixture.Image(i => i.Format("GIF"));
        var shapes = pres.Slide(1).Shapes;

        // Act
        shapes.AddPicture(gif);

        // Assert
        var picture = shapes.Last().Picture;
        picture.Image.Mime.Should().Be("image/gif");
    }

    [Test]
    public void Contains_expected_count_of_each_shape_type()
    {
        // Arrange
        var pres = new Presentation(TestAsset("003.pptx"));
        var shapes = pres.Slides.First().Shapes;

        // Act & Assert
        shapes.Count(sp => sp.Chart is not null).Should().Be(1);
        shapes.Count(sp => sp.Picture is not null).Should().Be(1);
        shapes.Count(sp => sp.Table is not null).Should().Be(1);
        shapes.Count(sp => sp.GroupedShapes is not null).Should().Be(1);
    }

    [Test]
    public void Contains_picture()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shape = pres.Slide(2).Shape(3);

        // Act-Assert
        var picture = shape.Picture;
        picture.Should().NotBeNull();
    }

    [Test]
    public void Contains_Media_Shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("audio-case001.pptx"));

        // Act-Assert
        pres.Slide(1).Shape(8).Media.Should().NotBeNull();
    }

    [Test]
    public void Contains_LineShape()
    {
        var pres = new Presentation(TestAsset("001.pptx"));

        // Act-Assert
        pres.Slide(1).Shape(10).Line.Should().NotBeNull();
    }

    [Test]
    public void Contains_Video_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("040_video.pptx"));
        var shape = pres.Slide(1).Shape(8);

        // Act-Assert
        shape.Media.Should().NotBeNull();
    }

    [Test]
    public void AddLine_adds_a_new_Line_shape_from_raw_open_xml_content()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var xml = StringOf("line-shape.xml");
        var shapes = pres.Slide(1).Shapes;

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
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 10, startPointY: 10, endPointX: 20, endPointY: 5);

        // Assert
        var addedLine = shapes.Last().Line;
        shapes.Should().ContainSingle();
        addedLine.StartPoint.X.Should().Be(10);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(20);
        addedLine.EndPoint.Y.Should().Be(5);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddLine_adds_line_Up_Up()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 10, startPointY: 10, endPointX: 10, endPointY: 5);

        // Assert
        var addedLine = shapes.Last().Line;
        addedLine.StartPoint.X.Should().Be(10);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(10);
        addedLine.EndPoint.Y.Should().Be(5);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddLine_adds_line_Left_Up()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 100, startPointY: 50, endPointX: 40, endPointY: 20);

        // Assert
        var addedLine = shapes.Last().Line;
        addedLine.StartPoint.X.Should().Be(100);
        addedLine.StartPoint.Y.Should().Be(50);
        addedLine.EndPoint.X.Should().Be(40);
        addedLine.EndPoint.Y.Should().Be(20);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddLine_adds_line_Left_Down()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 10, endPointX: 40, endPointY: 20);

        // Assert
        var addedLine = shapes.Last().Line;
        addedLine.StartPoint.X.Should().Be(50);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(40);
        addedLine.EndPoint.Y.Should().Be(20);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddLine_adds_line_Right_Right()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Assert
        var line = shapes.Last().Line;
        line.StartPoint.X.Should().Be(50);
        line.StartPoint.Y.Should().Be(60);
        line.EndPoint.X.Should().Be(100);
        line.EndPoint.Y.Should().Be(60);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddLine_adds_line()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Assert
        shapes.Should().ContainSingle();
        var lineShape = shapes.Last();
        lineShape.X.Should().Be(50);
        lineShape.Y.Should().Be(60);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddLine_adds_line_Left_Left()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 100, startPointY: 50, endPointX: 80, endPointY: 50);

        // Assert
        var line = shapes.Last().Line;
        line.StartPoint.X.Should().Be(100);
        line.StartPoint.Y.Should().Be(50);
        line.EndPoint.X.Should().Be(80);
        line.EndPoint.Y.Should().Be(50);
        ValidatePresentation(pres);
    }
    
    [Test]
    public void AddAudio_adds_audio_shape_with_WAVE_content()
    {
        // Arrange
        var wav = TestAsset("071 wav.wav");
        var pres = new Presentation(TestAsset("001.pptx"));
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.AddAudio(300, 100, wav, AudioType.WAVE);

        // Assert
        var addedAudio = pres.Slides[1].Shapes.Last();
        addedAudio.X.Should().Be(300);
    }

#if DEBUG
    [Test, Explicit("Should be implemented with https://github.com/ShapeCrawler/ShapeCrawler/issues/581")]
    public void AddAudio_adds_audio_shape_with_the_default_start_mode_In_Click_Sequence()
    {
        // Arrange
        var pres = new Presentation();
        var mp3 = TestAsset("064 mp3.mp3");
        var shapes = pres.Slide(1).Shapes;

        // Act
        shapes.AddAudio(x: 300, y: 100, mp3, AudioType.MP3);

        // Assert
        pres = SaveAndOpenPresentation(pres);
        var addedAudio = pres.Slide(1).First<IMedia>();
        ValidatePresentation(pres);
        addedAudio.StartMode.Should().Be(AudioStartMode.InClickSequence);
    }
#endif

    [Test]
    public void AddVideo_adds_Video_shape()
    {
        // Arrange
        var preStream = TestAsset("001.pptx");
        var pres = new Presentation(preStream);
        var shapes = pres.Slides[1].Shapes;
        var videoStream = TestAsset("079 mp4 video.mp4");
        int xPxCoordinate = 300;
        int yPxCoordinate = 100;

        // Act
        shapes.AddVideo(xPxCoordinate, yPxCoordinate, videoStream);

        // Assert
        pres.Save();
        pres = new Presentation(preStream);
        var addedVideo = pres.Slides[1].Shapes.Last();
        addedVideo.X.Should().Be(xPxCoordinate);
        addedVideo.Y.Should().Be(yPxCoordinate);
    }

    [Test]
    public void Add_adds_picture_from_same_slide()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.Shape("Picture")!;
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        shapes.Shape("Picture 1").Should().NotBeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void AddShape_adds_rectangle_with_valid_id_and_name()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case011_save-as-png.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70);

        // Assert
        var autoShape = shapes.Last();
        autoShape.Name.Should().Be("Rectangle");
        autoShape.Id.Should().Be(7);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddRectangle_adds_Rectangle_in_the_New_Presentation()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70);

        // Assert
        var rectangle = shapes.Last();
        rectangle.GeometryType.Should().Be(Geometry.Rectangle);
        rectangle.X.Should().Be(50);
        rectangle.Y.Should().Be(60);
        rectangle.Width.Should().Be(100);
        rectangle.Height.Should().Be(70);
        rectangle.TextBox!.Paragraphs.Count.Should().Be(1);
        rectangle.Outline.HexColor.Should().BeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void AddShape_adds_Rounded_Rectangle()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70, Geometry.RoundedRectangle);

        // Assert
        var roundedRectangle = shapes.Last();
        roundedRectangle.GeometryType.Should().Be(Geometry.RoundedRectangle);
        roundedRectangle.Name.Should().Be("RoundedRectangle");
        roundedRectangle.Outline.HexColor.Should().BeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void AddShape_adds_Top_Corners_Rounded_Rectangle()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70, Geometry.TopCornersRoundedRectangle);

        // Assert
        var addedTopCornersRoundedRectangle = shapes.Last();
        addedTopCornersRoundedRectangle.GeometryType.Should().Be(Geometry.TopCornersRoundedRectangle);
        addedTopCornersRoundedRectangle.Name.Should().Be("TopCornersRoundedRectangle");
        ValidatePresentation(pres);
    }

    [Test]
    public void AddTable_adds_table()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddTable(x: 50, y: 60, columnsCount: 3, rowsCount: 2);

        // Assert
        var tableShape = shapes.Last();
        var table = tableShape.Table;
        table.Columns.Should().HaveCount(3);
        table.Rows.Should().HaveCount(2);
        tableShape.Id.Should().Be(1);
        tableShape.Name.Should().Be("Table 1");
        table.Columns[0].Width.Should().BeApproximately(213.33m, 0.01m);
        ValidatePresentation(pres);
    }

    [Test]
    [LayoutShape("autoshape-case004_subtitle.pptx", 1, "Group 1")]
    [MasterShape("autoshape-case004_subtitle.pptx", "Group 1")]
    public void GetByName_returns_shape_by_specified_name(IShape shape)
    {
        // Act
        var groupedShape = shape.GroupedShape("AutoShape 1");

        // Assert
        groupedShape.Should().NotBeNull();
    }

    [Test]
    [TestCase("002.pptx", 1, 4)]
    [TestCase("003.pptx", 1, 5)]
    [TestCase("013.pptx", 1, 4)]
    [TestCase("023.pptx", 1, 1)]
    [TestCase("014.pptx", 3, 5)]
    [TestCase("009_table.pptx", 1, 6)]
    [TestCase("009_table.pptx", 2, 8)]
    public void Count_returns_number_of_shapes(string pptxName, int slideNumber, int expectedCount)
    {
        // Arrange
        var pres = new Presentation(TestAsset(pptxName));
        var slide = pres.Slides[slideNumber - 1];

        // Act
        int shapesCount = slide.Shapes.Count;

        // Assert
        shapesCount.Should().Be(expectedCount);
    }

    [Test]
    public void Count_returns_one_When_presentation_contains_one_slide()
    {
        // Act
        var pptx17 = TestAsset("017.pptx");
        var pres17 = new Presentation(pptx17);
        var pptx16 = TestAsset("016.pptx");
        var pres16 = new Presentation(pptx16);
        var numberSlidesCase1 = pres17.Slides.Count;
        var numberSlidesCase2 = pres16.Slides.Count;

        // Assert
        numberSlidesCase1.Should().Be(1);
        numberSlidesCase2.Should().Be(1);
    }

    [Test]
    public void Add_adds_slide_from_the_Same_presentation()
    {
        // Arrange
        var pptxStream = TestAsset("003 chart.pptx");
        var pres = new Presentation(pptxStream);
        var expectedSlidesCount = pres.Slides.Count + 1;
        var slideCollection = pres.Slides;
        var addingSlide = slideCollection[0];

        // Act
        pres.Slides.Add(addingSlide);

        // Assert
        pres.Slides.Count.Should().Be(expectedSlidesCount);
    }

    [Test]
    public void Add_adds_slide_After_updating_chart_series()
    {
        // Arrange
        var pptx = TestAsset("001 bar chart.pptx");
        var pres = new Presentation(pptx);
        var chart = pres.Slides[0].Shape("Bar Chart 1").Chart;
        var expectedSlidesCount = pres.Slides.Count + 1;

        // Act
        chart.SeriesCollection[0].Points[0].Value = 1;
        pres.Slides.Add(pres.Slides[0]);

        // Assert
        pres.Slides.Count.Should().Be(expectedSlidesCount);
    }

    [Test]
    public void AddEmptySlide_adds_New_slide()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var layout = pres.MasterSlides[0].LayoutSlides[0];
        var slides = pres.Slides;

        // Act
        slides.Add(layout.Number);

        // Assert
        var addedSlide = slides.Last();
        addedSlide.Should().NotBeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void Add_adds_a_new_slide()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var layout = pres.MasterSlides[0].LayoutSlides.First(l => l.Name == "Blank");
        var slides = pres.Slides;

        // Act
        slides.Add(layout.Number);

        // Assert
        slides[1].Shapes.Should().HaveCount(0);
    }
    
    [Test]
    public void AddPieChart_adds_pie_chart()
    {
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        var categoryValues = new Dictionary<string, double> { { "1st Qtr", 10 }, { "2nd Qtr", 20 }, { "3rd Qtr", 30 } };

        // Act
        shapes.AddPieChart(100, 100, 400, 300, categoryValues, "Sales");

        // Assert
        shapes.First().Chart.Should().NotBeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void AddBarChart_adds_bar_chart()
    {
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var categoryValues = new Dictionary<string, double>
        {
            { "Category 1", 10 },
            { "Category 2", 25 },
            { "Category 3", 15 }
        };
        string seriesName = "Sample Series";

        // Act
        shapes.AddBarChart(x, y, width, height, categoryValues, seriesName);

        // Assert
        shapes.First().Chart.Should().NotBeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void AddScatterChart_adds_scatter_chart()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var pointValues = new Dictionary<double, double>
        {
            { 1.0, 5.2 },
            { 2.0, 7.3 },
            { 3.0, 8.1 },
            { 4.0, 9.5 },
            { 5.0, 12.3 }
        };
        string seriesName = "Data Series";

        // Act
        shapes.AddScatterChart(x, y, width, height, pointValues, seriesName);

        // Assert
        var chart = shapes.First().Chart;
        chart.Type.Should().Be(ChartType.ScatterChart);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddStackedColumnChart_adds_stacked_column_chart()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var categoryValues = new Dictionary<string, IList<double>>
        {
            { "Category 1", new List<double> { 10, 20 } },
            { "Category 2", new List<double> { 30, 40 } },
            { "Category 3", new List<double> { 50, 60 } }
        };
        var seriesNames = new List<string> { "Series 1", "Series 2" };

        // Act
        shapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);

        // Assert
        shapes.First().Chart.Type.Should().Be(ChartType.BarChart);
        ValidatePresentation(pres);
    }

    [Test]
    public void Group_groups_shapes()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        shapes.AddShape(100, 100, 100, 100, Geometry.Rectangle, "Shape 1");
        shapes.AddShape(100, 200, 100, 100, Geometry.Rectangle, "Shape 2");
        var shape1 = shapes[0];
        var shape2 = shapes[1];

        // Act
        var group = shapes.Group([shape1, shape2]);

        // Assert
        group.GroupedShapes.Should().HaveCount(2);
        ValidatePresentation(pres);
    }

    [Test]
    public void Add_adds_chart_from_another_presentation()
    {
        // Arrange
        var sourcePres = new Presentation(p => p.Slide());
        var categoryValues = new Dictionary<string, double>
        {
            { "Category 1", 10 },
            { "Category 2", 20 },
            { "Category 3", 30 }
        };
        sourcePres.Slide(1).Shapes.AddPieChart(100, 100, 400, 300, categoryValues, "Sales", "Chart 1");
        var chartShape = sourcePres.Slide(1).Shapes.Shape("Chart 1");
        var targetPres = new Presentation(p => p.Slide());

        // Act
        targetPres.Slide(1).Shapes.Add(chartShape);

        // Assert
        ValidatePresentation(targetPres);
        var copiedChart = targetPres.Slide(1).Shapes.Shape("Chart 1");
        copiedChart.Chart.Should().NotBeNull();
    }
}
