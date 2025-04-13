using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;
using NUnit.Framework;
using System.Text.Json;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Tests.Unit;

[TestFixture]
public class ShapeTests : SCTest
{
    [Test]
    public void AudioShape_BinaryData_returns_audio_bytes()
    {
        // Arrange
        var pptx = TestAsset("audio-case001.pptx");
        var pres = new Presentation(pptx);
        var audioShape = pres.Slides[0].Shapes.GetByName<IMediaShape>("Audio 1");

        // Act
        var bytes = audioShape.AsByteArray();

        // Assert
        bytes.Should().NotBeEmpty();
    }

    [Test]
    public void AudioShape_MIME_returns_MIME_type_of_audio_content()
    {
        // Arrange
        var pptxStream = TestAsset("audio-case001.pptx");
        var pres = new Presentation(pptxStream);
        var audioShape = pres.Slides[0].Shapes.GetByName<IMediaShape>("Audio 1");

        // Act
        var mime = audioShape.Mime;

        // Assert
        mime.Should().Be("audio/mpeg");
    }

    [Test]
    public void VideoShape_BinaryData_returns_video_bytes()
    {
        // Arrange
        var pptxStream = TestAsset("video-case001.pptx");
        var pres = new Presentation(pptxStream);
        var videoShape = pres.Slides[0].Shapes.GetByName<IMediaShape>("Video 1");

        // Act
        var bytes = videoShape.AsByteArray();

        // Assert
        bytes.Should().NotBeEmpty();
    }

    [Test]
    public void AudioShape_MIME_returns_MIME_type_of_video_content()
    {
        // Arrange
        var pptxStream = TestAsset("video-case001.pptx");
        var pres = new Presentation(pptxStream);
        var videoShape = pres.Slides[0].Shapes.GetByName<IMediaShape>("Video 1");

        // Act
        var mime = videoShape.Mime;

        // Assert
        mime.Should().Be("video/mp4");
    }

    [Test]
    public void PictureSetImage_ShouldNotImpactOtherPictureImage_WhenItsOriginImageIsShared()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var image = TestAsset("10 png image.png");
        IPresentation presentation = new Presentation(pptx);
        IPicture picture5 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 5);
        IPicture picture6 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        int pic6LengthBefore = picture6.Image.AsByteArray().Length;
        MemoryStream modifiedPresentation = new();

        // Act
        picture5.Image.Update(image);

        // Assert
        int pic6LengthAfter = picture6.Image.AsByteArray().Length;
        pic6LengthAfter.Should().Be(pic6LengthBefore);

        presentation.Save(modifiedPresentation);
        presentation = new Presentation(modifiedPresentation);
        picture6 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        pic6LengthBefore = picture6.Image.AsByteArray().Length;
        pic6LengthAfter.Should().Be(pic6LengthBefore);
    }

    [Test]
    public void Y_Getter_returns_y_coordinate()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("006_1 slides.pptx"));
        var pres2 = new Presentation(TestAsset("018.pptx"));
        var pres3 = new Presentation(TestAsset("009_table.pptx"));
        var shapeCase1 = pres1.Slide(1).Shapes.First(sp => sp.Id == 2);
        var shapeCase2 = pres2.Slide(1).Shapes.First(sp => sp.Id == 7);
        var shapeCase3 = pres3.Slide(2).Shapes.First(sp => sp.Id == 9);
        float verticalResolution = 96;

        // Act & Assert
        shapeCase1.Y.Should().BeApproximately(88.37m, 0.01m);
        shapeCase2.Y.Should().BeApproximately(0.00031m, 0.00001m);
        shapeCase3.Y.Should().BeApproximately(272.69m, 0.01m);
    }

    [Test]
    public void Id_returns_id()
    {
        // Arrange
        var pres = new Presentation(TestAsset("010.pptx"));
        var shape = pres.SlideMasters[0].Shapes.GetByName<IShape>("Date Placeholder 3");

        // Act
        var id = shape.Id;

        // Assert
        id.Should().Be(9);
    }

    [Test]
    public void Grouped_Shape_Y_Setter_raises_up_group_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slide(1).Shape<IGroupShape>("Group 2");
        var groupedShape = groupShape.Shapes.GetByName<IShape>("Shape 1");

        // Act
        groupedShape.Y = 307;

        // Assert
        groupedShape.Y.Should().Be(307);
        groupShape.Y.Should().Be(307, "because the moved grouped shape was on the up-hand side");
        groupShape.Height.Should().BeApproximately(91.87m, 0.01m);
    }

    [Test]
    public void Y_Setter_increases_the_height_of_the_group_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slide(1).Shape<IGroupShape>("Group 2");
        var groupedShape = groupShape.Shapes.GetByName<IShape>("Shape 2");

        // Act
        groupedShape.Y = 372;

        // Assert
        groupedShape.Y.Should().Be(372);
        groupShape.Height.Should().BeApproximately(90.08m, 0.01m);
    }

    [Test]
    public void X_Setter_moves_the_Left_hand_grouped_shape_to_Left()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");
        var groupedShape = groupShape.Shapes.GetByName<IShape>("Shape 1");

        // Act
        groupedShape.X = 49m;

        // Assert
        groupedShape.X.Should().Be(49);
        groupShape.X.Should().Be(49, "because the moved grouped shape was on the left-hand side");
        groupShape.Width.Should().BeApproximately(89.18m, 0.01m);
    }

    [Test]
    public void X_Setter_moves_the_Right_hand_grouped_shape_to_Right()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");
        var groupedShape = groupShape.Shapes.GetByName<IShape>("Shape 1");
        var groupShapeX = groupShape.X;

        // Act
        groupedShape.X = 69m;

        // Assert
        groupedShape.X.Should().Be(69m);
        groupShape.X.Should().BeApproximately(groupShapeX, 0.01m,
            "because the X-coordinate of parent group shouldn't be changed when a grouped shape is moved to the right side");
        groupShape.Width.Should().BeApproximately(87.72m, 0.01m);
    }

    [Test]
    public void Width_returns_shape_width_in_points()
    {
        // Arrange
        var shapeCase1 = new Presentation(TestAsset("006_1 slides.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        var groupShape =
            (IGroupShape)new Presentation(TestAsset("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 7);
        var shapeCase2 = groupShape.Shapes.First(sp => sp.Id == 5);
        var shapeCase3 = new Presentation(TestAsset("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);

        // Act & Assert
        shapeCase1.Width.Should().BeApproximately(720m, 0.01m);
        shapeCase2.Width.Should().BeApproximately(93.02m, 0.01m);
        shapeCase3.Width.Should().BeApproximately(38.252m, 0.01m);
    }

    [Test]
    public void Height_returns_Grouped_Shape_height_in_pixels()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);
        var groupShape = pres.Slides[1].Shapes.GetByName<IGroupShape>("Group 1");
        var groupedShape = groupShape.Shapes.GetByName<IShape>("Shape 2");

        // Act
        var height = groupedShape.Height;

        // Assert
        height.Should().BeApproximately(51.50m, 0.01m);
    }

    [TestCase(2, Geometry.Rectangle)]
    [TestCase(3, Geometry.Ellipse)]
    public void GeometryType_returns_shape_geometry_type(int shapeId, Geometry expectedGeometryType)
    {
        // Arrange
        var presentation = new Presentation(TestAsset("021.pptx"));

        // Act
        var shape = presentation.Slides[3].Shapes.First(sp => sp.Id == shapeId);

        // Assert
        shape.GeometryType.Should().Be(expectedGeometryType);
    }

    [Test]
    public void Shape_IsNotGroupShape()
    {
        // Arrange
        IShape shape = new Presentation(TestAsset("006_1 slides.pptx")).Slides[0].Shapes.First(x => x.Id == 3);

        // Act-Assert
        shape.Should().NotBeOfType<IGroupShape>();
    }

    [Test]
    public void CustomData_ReturnsNull_WhenShapeHasNotCustomData()
    {
        // Arrange
        var shape = new Presentation(TestAsset("009_table.pptx")).Slides.First().Shapes.First();

        // Act
        var shapeCustomData = shape.CustomData;

        // Assert
        shapeCustomData.Should().BeNull();
    }

    [Test]
    public void CustomData_ReturnsCustomDataOfTheShape_WhenShapeWasAssignedSomeCustomData()
    {
        // Arrange
        const string customDataString = "Test custom data";
        var savedPreStream = new MemoryStream();
        var presentation = new Presentation(TestAsset("009_table.pptx"));
        var shape = presentation.Slides.First().Shapes.First();

        // Act
        shape.CustomData = customDataString;
        presentation.Save(savedPreStream);

        // Assert
        presentation = new Presentation(savedPreStream);
        shape = presentation.Slides.First().Shapes.First();
        shape.CustomData.Should().Be(customDataString);
    }

    [Test]
    public void Name_ReturnsShapeNameString()
    {
        // Arrange
        IShape shape = new Presentation(TestAsset("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 8);

        // Act
        string shapeName = shape.Name;

        // Assert
        shapeName.Should().BeEquivalentTo("Object 2");
    }

    [TestCase(0, true)]
    [TestCase(1, false)]
    public void Hidden_ReturnsValueIndicatingWhetherShapeIsHiddenFromTheSlide(int shapeIndex, bool expectedHidden)
    {
        // Arrange
        var pptx = TestAsset("004.pptx");
        var pres = new Presentation(pptx);
        var shape = pres.Slides[0].Shapes[shapeIndex];

        // Act-Assert
        shape.Hidden.Should().Be(expectedHidden);
    }

    [TestCase("autoshape-case018_rotation.pptx", 1, "RotationTextBox", 325.40)]
    [TestCase("autoshape-case018_rotation.pptx", 2, "VerticalTextPH", 281.97)]
    [TestCase("autoshape-case018_rotation.pptx", 2, "NoRotationGroup", 0)]
    [TestCase("autoshape-case018_rotation.pptx", 2, "RotationGroup", 55.60)]
    public void Rotation_returns_shape_rotation_in_degrees(string presentationName, int slideNumber, string shapeName,
        double expectedAngle)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentationName));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);

        // Act
        var rotation = shape.Rotation;

        // Assert
        rotation.Should().BeApproximately(expectedAngle, 0.01);
    }

    [Test]
    [TestCase("001.pptx", 1, "TextBox 3")]
    [TestCase("001.pptx", 1, "Head 1")]
    [TestCase("autoshape-grouping.pptx", 1, "Group 1")]
    public void Y_Setter_sets_y_coordinate(string file, int slideNumber, string shapeName)
    {
        // Act
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        shape.Y = 100;

        // Assert
        shape.Y.Should().Be(100);
        pres.Validate();
    }

    [Test]
    [TestCase("006_1 slides.pptx", 1, "Shape 1")]
    [TestCase("001.pptx", 1, "Head 1")]
    [TestCase("autoshape-grouping.pptx", 1, "Group 1")]
    [TestCase("table-case001.pptx", 1, "Table 1")]
    public void X_Setter_sets_x_coordinate(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pptx = TestAsset(file);
        var pres = new Presentation(pptx);
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var stream = new MemoryStream();

        // Act
        shape.X = 400;

        // Assert
        pres.Save(stream);
        pres = new Presentation(stream);
        shape = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName);
        shape.X.Should().Be(400);
        pres.Validate();
    }

    [Test]
    [TestCase("006_1 slides.pptx", 1, "Shape 1")]
    public void Width_Setter_sets_width(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pptx = TestAsset(file);
        var pres = new Presentation(pptx);
        var stream = new MemoryStream();
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);

        // Act
        shape.Width = 600;

        // Assert
        pres.Save(stream);
        pres = new Presentation(stream);
        shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        shape.Width.Should().Be(600);
        pres.Validate();
    }

    [Test]
    public void Remove_removes_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var shape = pres.Slides[0].Shape("TextBox 3");

        // Act
        shape.Remove();

        // Assert
        var act = () => pres.Slides[0].Shapes.GetByName("TextBox 3");
        act.Should().Throw<SCException>();
        pres.Validate();
    }

    [Test]
    [SlideShape("021.pptx", slideNumber: 4, shapeId: 2, expectedResult: PlaceholderType.Footer)]
    [SlideShape("008.pptx", 1, 3, PlaceholderType.DateAndTime)]
    [SlideShape("019.pptx", 1, 2, PlaceholderType.SlideNumber)]
    [SlideShape("013.pptx", 1, 281, PlaceholderType.Content)]
    [SlideShape("autoshape-case016.pptx", 1, "Content Placeholder 1", PlaceholderType.Content)]
    [SlideShape("autoshape-case016.pptx", 1, "Text Placeholder 1", PlaceholderType.Text)]
    [SlideShape("autoshape-case016.pptx", 1, "Picture Placeholder 1", PlaceholderType.Picture)]
    [SlideShape("autoshape-case016.pptx", 1, "Table Placeholder 1", PlaceholderType.Table)]
    [SlideShape("autoshape-case016.pptx", 1, "SmartArt Placeholder 1", PlaceholderType.SmartArt)]
    [SlideShape("autoshape-case016.pptx", 1, "Media Placeholder 1", PlaceholderType.Media)]
    [SlideShape("autoshape-case016.pptx", 1, "Online Image Placeholder 1", PlaceholderType.OnlineImage)]
    public void PlaceholderType_returns_placeholder_type(IShape shape, PlaceholderType expectedType)
    {
        // Act
        var placeholderType = shape.PlaceholderType;

        // Assert
        placeholderType.Should().Be(expectedType);
    }

    [Test]
    [TestCase("054_get_shape_xpath.pptx", 1, 1, null)]
    [TestCase("054_get_shape_xpath.pptx", 1, 2, "Title 1")]
    [TestCase("054_get_shape_xpath.pptx", 1, 3, "SubTitle 2")]
    [TestCase("054_get_shape_xpath.pptx", 1, 4, null)]
    public void TryGetSlideShapeById(string presentationName, int slideNumber, int shapeId, string? expectedShapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentationName));
        var slide = pres.Slides[slideNumber - 1];
        var shape = slide.Shapes.TryGetById<IShape>(shapeId);

        // Act
        var shapeName = shape?.Name;

        // Assert
        shapeName.Should().Be(expectedShapeName);
    }

    [Test]
    [TestCase("054_get_shape_xpath.pptx", 1, "Foo", null)]
    [TestCase("054_get_shape_xpath.pptx", 1, "Title 1", 2)]
    [TestCase("054_get_shape_xpath.pptx", 1, "SubTitle 2", 3)]
    [TestCase("054_get_shape_xpath.pptx", 1, "Bar", null)]
    public void TryGetSlideShapeByName(string presentationName, int slideNumber, string shapeName, int? expectedShapeId)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentationName));
        var slide = pres.Slides[slideNumber - 1];
        var shape = slide.Shapes.TryGetByName<IShape>(shapeName);

        // Act
        var shapeId = shape?.Id;

        // Assert
        shapeId.Should().Be(expectedShapeId);
    }

    [Test]
    public void AsTable_returns_ITable()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var slide = pres.Slides[0];
        var table = slide.Shapes.GetByName<ITable>("Table 1");

        // Act
        var castingToITable = () => table.AsTable();

        // Assert
        castingToITable.Should().NotThrow();
    }

    [Test]
    [SlideShape("021.pptx", 4, 2, 287.68)]
    [SlideShape("008.pptx", 1, 3, 49.5)]
    [SlideShape("006_1 slides.pptx", 1, 2, 120)]
    [SlideShape("009_table.pptx", 2, 9, 55.06)]
    [SlideShape("025_chart.pptx", 3, 7, 59.63)]
    [SlideShape("018.pptx", 1, "Picture Placeholder 1", 7.27)]
    public void X_Getter_returns_x_coordinate(IShape shape, double expectedX)
    {
        // Act
        decimal x = shape.X;
        var expectedXPoints = (decimal)expectedX;

        // Assert
        x.Should().BeApproximately(expectedXPoints, 0.01m);
    }

    [Test]
    public void X_Getter_returns_x_coordinate_of_Grouped_shape_in_points()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shape = pres.Slides[1].Shapes.GetByName<IGroupShape>("Group 1").Shapes.GetByName<IShape>("Shape 1");

        // Act
        decimal x = shape.X;

        // Assert
        x.Should().BeApproximately(39.94m, 0.01m);
    }

    [Test]
    [TestCase("050_title-placeholder.pptx", 1, 2, 583.2)]
    [TestCase("051_title-placeholder.pptx", 1, 3074, 648)]
    public void Width_returns_width_of_Title_placeholder(
        string filename,
        int slideNumber,
        int shapeId,
        decimal expectedWidth)
    {
        // Arrange
        var pres = new Presentation(TestAsset(filename));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId);

        // Act
        var shapeWidth = shape.Width;

        // Assert
        shapeWidth.Should().Be(expectedWidth);
    }

    [Test]
    [SlideShape("006_1 slides.pptx", 1, "Shape 2", 112.24)]
    [SlideShape("009_table.pptx", 2, "Object 3", 29.37)]
    [SlideShape("autoshape-grouping.pptx", 1, "Group 2", 81.01)]
    public void Height_returns_shape_height_in_points(IShape shape, double expectedHeight)
    {
        // Act
        var height = shape.Height;

        // Assert
        height.Should().BeApproximately((decimal)expectedHeight, 0.01m);
    }

    [Test]
    [SlideShape("021.pptx", 4, 2, Geometry.Rectangle)]
    [SlideShape("021.pptx", 4, 3, Geometry.Ellipse)]
    public void GeometryType_returns_shape_geometry_type(IShape shape, Geometry expectedGeometryType)
    {
        // Assert
        shape.GeometryType.Should().Be(expectedGeometryType);
    }

    [Test]
    [SlideShape("054_get_shape_xpath.pptx", 1, "Title 1", "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]")]
    [SlideShape("054_get_shape_xpath.pptx", 1, "SubTitle 2", "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[2]")]
    public void SDKXPath_returns_shape_xpath(IShape shape, string expectedXPath)
    {
        // Act
        var shapeXPath = shape.SDKXPath;

        // Assert
        shapeXPath.Should().Be(expectedXPath);
    }

    [Test]
    public void Text_Setter_updates_shape_text()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case017_slide-number.pptx"));
        var shape = pres.SlideMaster(1).Shape("Shape 1");

        // Act
        shape.TextBox!.Text = "Test";

        // Assert
        shape.TextBox.Text.Should().Be("Test");
    }

    [Test]
    [SlideShape("057_corner-radius.pptx", 1, "Size 1 Round 0.25", "20.834")]
    [SlideShape("057_corner-radius.pptx", 1, "Size 2 Round 0.25", "20.834")]
    [SlideShape("057_corner-radius.pptx", 1, "Size 3 Round 0.25", "20.834")]
    [SlideShape("057_corner-radius.pptx", 1, "Size 1 Round 0", "0")]
    [SlideShape("057_corner-radius.pptx", 1, "Size 1 Round X", "35")]
    [SlideShape("057_corner-radius.pptx", 1, "Size 1 Round 1", "100")]
    [SlideShape("057_corner-radius.pptx", 1, "Size 1 Round 0.75", "61.112")]
    public void CornerSize_getter_returns_values(IShape shape, string expectedSizeStr)
    {
        // Arrange
        var expectedSize = decimal.Parse(expectedSizeStr);

        // Act
        var actualSize = shape.CornerSize;

        // Assert
        actualSize.Should().Be(expectedSize);
    }

    [Test]
    public void CornerSize_Getter_corner_size()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(10, 20, 100, 200, Geometry.RoundedRectangle);
        var shape = shapes[0];

        // Act-Assert
        shape.CornerSize.Should().Be(35m, "Rounded rectangles with no specified corner size behave as if the value was set to 35%.");
    }

    [Test]
    [SlideShape("057_corner-radius.pptx", 4, "Top Rounded 0.125-ish", "12.29")]
    [SlideShape("057_corner-radius.pptx", 4, "Top Rounded 0", "0")]
    [SlideShape("057_corner-radius.pptx", 4, "Top Rounded X", "35")]
    [SlideShape("057_corner-radius.pptx", 4, "Top Rounded 1", "100")]
    public void CornerSize_Getter_returns_corner_size_for_Top_Rounded_Rectangle(IShape shape, string expectedCornerSizeStr)
    {
        // Arrange
        var expectedCornerSize = decimal.Parse(expectedCornerSizeStr);

        // Act-Assert
        shape.CornerSize.Should().Be(expectedCornerSize);
    }

    [Test]
    [SlideShape("057_corner-radius.pptx", 4, "Top Rounded 0.125-ish", "0.5")]
    [SlideShape("057_corner-radius.pptx", 4, "Top Rounded 0", "0.50")]
    public void CornerSize_setter_sets_values_for_top_rounded(IShape shape, string expectedSizeStr)
    {
        // Arrange
        var expectedSize = decimal.Parse(expectedSizeStr);

        // Act
        shape.CornerSize = expectedSize;

        // Assert
        var actualSize = shape.CornerSize;
        actualSize.Should().Be(expectedSize);
    }

    [Test]
    public void CornerSize_Setter_sets_corner_size()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(10, 20, 100, 200, Geometry.RoundedRectangle);
        var shape = shapes[0];

        // Act
        shape.CornerSize = 0.1m;

        // Assert
        shape.CornerSize.Should().Be(0.1m);
        pres.Validate();
    }

    [TestCase("RoundedRectangle")]
    [TestCase("Triangle")]
    [TestCase("Diamond")]
    [TestCase("Parallelogram")]
    [TestCase("Trapezoid")]
    [TestCase("NonIsoscelesTrapezoid")]
    [TestCase("DiagonalCornersRoundedRectangle")]
    [TestCase("TopCornersRoundedRectangle")]
    [TestCase("SingleCornerRoundedRectangle")]
    [TestCase("UTurnArrow")]
    [TestCase("LineInverse")]
    [TestCase("RightTriangle")]
    public void Geometry_setter_sets_values(string expectedStr)
    {
        // Arrange
        var expected = (Geometry)Enum.Parse(typeof(Geometry), expectedStr);
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(50, 60, 100, 70);
        var shape = shapes.Last();

        // Act
        shape.GeometryType = expected;

        // Assert
        shape.GeometryType.Should().Be(expected);
        pres.Validate();
    }

    [Test]
    public void Geometry_setter_wont_set_custom()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(50, 60, 100, 70);
        var shape = shapes.Last();

        // Act
        var act = () => shape.GeometryType = Geometry.Custom;

        // Assert
        act.Should().Throw<SCException>("Custom geometry cannot be set");
    }

    [Test]
    public void Geometry_setter_resets_old_adjustments()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(50, 60, 100, 70);
        var shape = shapes.Last();
        shape.GeometryType = Geometry.RoundedRectangle;
        shape.CornerSize = 100;

        // Act
        shape.GeometryType = Geometry.TopCornersRoundedRectangle;

        // Assert
        shape.CornerSize.Should().Be(35m, "Default unadjusted corner size is 35");
    }

    [Test]
    public void Name_Setter_sets_shape_name()
    {
        // Arrange
        var pres = new Presentation(TestAsset("006_1 slides.pptx"));
        var stream = new MemoryStream();
        var shape = pres.Slides[0].Shapes.GetByName("Shape 1");

        // Act
        shape.Name = "New Name";

        // Assert
        pres.Save(stream);
        pres = new Presentation(stream);
        shape = pres.Slides[0].Shapes.GetByName("New Name");
        shape.Name.Should().Be("New Name");
        pres.Validate();
    }

    [Test]
    public void Name_Setter_sets_grouped_shape_name()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var stream = new MemoryStream();
        var groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");

        // Act
        groupShape.Name = "New Group Name";

        // Assert
        pres.Save(stream);
        pres = new Presentation(stream);
        groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("New Group Name");
        groupShape.Name.Should().Be("New Group Name");
        pres.Validate();
    }

    [TestCase("Triangle", "[200]")]
    [TestCase("Parallelogram", "[0]")]
    [TestCase("Trapezoid", "[0]")]
    [TestCase("NonIsoscelesTrapezoid", "[0,100]")]
    [TestCase("Hexagon", "[0]")]
    [TestCase("Octagon", "[0]")]
    [TestCase("Star4", "[100]")]
    [TestCase("Star5", "[100]")]
    [TestCase("Star6", "[100]")]
    [TestCase("Star7", "[100]")]
    [TestCase("Star10", "[100]")]
    [TestCase("Star12", "[100]")]
    [TestCase("Star16", "[100]")]
    [TestCase("Star24", "[100]")]
    [TestCase("Star32", "[0]")]
    [TestCase("RoundedRectangle", "[100]")]
    [TestCase("TopCornersRoundedRectangle", "[100,0]")]
    [TestCase("DiagonalCornersRoundedRectangle", "[0,100]")]
    [TestCase("SingleCornerRoundedRectangle", "[100]")]
    [TestCase("SnipRoundRectangle", "[100,100]")]
    [TestCase("Snip1Rectangle", "[100]")]
    [TestCase("Snip2SameRectangle", "[100,78.704]")]
    [TestCase("Snip2DiagonalRectangle", "[44.444,100]")]
    [TestCase("Plaque", "[70.37]")]
    [TestCase("HomePlate", "[39.814]")]
    [TestCase("Chevron", "[30.556]")]
    [TestCase("Pie", "[39930.998,16264.266]")]
    [TestCase("BlockArc", "[10251.142,2160.504,22.112]")]
    [TestCase("Donut", "[79.504]")]
    [TestCase("NoSmoking", "[64.97]")]
    [TestCase("Donut", "[79.504]")]
    [TestCase("RightArrow", "[0,19.444]")]
    [TestCase("LeftArrow", "[200,7.408]")]
    [TestCase("UpArrow", "[15.972,28.124]")]
    [TestCase("DownArrow", "[5.556,37.152]")]
    [TestCase("StripedRightArrow", "[25.926,161.574]")]
    [TestCase("NotchedRightArrow", "[55.556,143.056]")]
    [TestCase("BentUpArrow", "[100,80.902,29.862]")]
    [TestCase("LeftRightArrow", "[0,20.486]")]
    [TestCase("UpDownArrow", "[41.666,27.778]")]
    [TestCase("LeftUpArrow", "[10.562,70.422,59.156]")]
    [TestCase("LeftRightUpArrow", "[0,20.422,50]")]
    [TestCase("QuadArrow", "[0,75.986,8.38]")]
    [TestCase("LeftArrowCallout", "[0,48.592,57.042,14.788]")]
    [TestCase("RightArrowCallout", "[50,26.056,79.578,76.434]")]
    [TestCase("UpArrowCallout", "[0,100,165.494,15.87]")]
    [TestCase("DownArrowCallout", "[0,100,145.774,6.01]")]
    [TestCase("UpDownArrowCallout", "[81.69,100,78.168,0]")]
    [TestCase("QuadArrowCallout", "[0,0,62.97,0]")]
    [TestCase("BentArrow", "[61.972,30.986,100,0]")]
    [TestCase("UTurnArrow", "[0,18.31,85.212,5.81,135.212]")]
    [TestCase("CircularArrow", "[77.376,6331.394,36798.45,9442.808,38.688]")]
    [TestCase("LeftCircularArrow", "[8.262,-2284.638,38241.328,15665.55,40.128]")]
    [TestCase("LeftRightCircularArrow", "[0,6773.122,1642.124,36866.022,8.998]")]
    [TestCase("CurvedRightArrow", "[33.77,17.324,111.458]")]
    [TestCase("CurvedLeftArrow", "[78.226,40.222,137.152]")]
    [TestCase("CurvedUpArrow", "[1.734,42.154,33.68]")]
    [TestCase("CurvedDownArrow", "[28.276,68.26,57.522]")]
    [TestCase("SwooshArrow", "[147.778, 140]")]
    [TestCase("CircularArrow", "[77.376,6331.394,36798.45,9442.808,38.688]")]
    [TestCase("LeftRightCircularArrow", "[0,6773.122,1642.124,36866.022,8.998]")]
    [TestCase("CurvedRightArrow", "[33.77,17.324,111.458]")]
    [TestCase("CurvedLeftArrow", "[78.226,40.222,137.152]")]
    [TestCase("CurvedUpArrow", "[1.734,42.154,33.68]")]
    [TestCase("CurvedDownArrow", "[28.276,68.26,57.522]")]
    [TestCase("Cube", "[121.528]")]
    [TestCase("Can", "[100]")]
    [TestCase("Sun", "[93.75]")]
    [TestCase("Moon", "[175]")]
    [TestCase("SmileyFace", "[-9.306]")] 
    [TestCase("FoldedCorner", "[100]")]
    [TestCase("Bevel", "[11.266]")]
    [TestCase("Frame", "[100]")]
    [TestCase("HalfFrame", "[107.638, 17.36]")]
    [TestCase("Corner", "[147.222,154.862]")]
    [TestCase("DiagonalStripe", "[32.87]")]
    [TestCase("Chord", "[12235.572,37755.642]")]
    [TestCase("Arc", "[14144.04,9163.314]")]
    [TestCase("LeftBracket", "[100]")]
    [TestCase("RightBracket", "[0]")]

    public void Adjustments_getter_returns_values(string name, string expectedAdjustmentsJson)
    {
        // Arrange
        var shape = 
            new Presentation(TestAsset("062_shape-adjustments.pptx"))
                .Slides[0]
                .Shapes
                .GetByName(name);
        var expectedAdjustments = JsonSerializer.Deserialize<decimal[]>(expectedAdjustmentsJson);

        // Act
        var actualAdjustments = shape.Adjustments;

        // Assert
        actualAdjustments.Should().BeEquivalentTo(expectedAdjustments);
    }

    [TestCase("RoundedRectangle", "[20]")]
    [TestCase("TopCornersRoundedRectangle", "[30,40]")]
    [TestCase("DiagonalCornersRoundedRectangle", "[50,60]")]
    [TestCase("SingleCornerRoundedRectangle", "[70]")]
    [TestCase("SnipRoundRectangle", "[20,80]")]
    [TestCase("Snip1Rectangle", "[10]")]
    [TestCase("Snip2SameRectangle", "[75,25]")]
    [TestCase("Snip2DiagonalRectangle", "[40,10]")]
    public void Adjustments_setter_sets_values(string geometryStr, string expectedAdjustmentsJson)
    {
        // Arrange
        var geometry = (Geometry)Enum.Parse(typeof(Geometry), geometryStr);
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(50, 60, 100, 70, geometry);
        var shape = shapes.Last();
        var expectedAdjustments = JsonSerializer.Deserialize<decimal[]>(expectedAdjustmentsJson);

        // Act
        shape.Adjustments = expectedAdjustments;

        // Assert
        shape.Adjustments.Should().BeEquivalentTo(expectedAdjustments);
    }

    [Test]
    public void Adjustments_setter_setting_only_one_value_when_multiple_allowed_on_existing_shape()
    {
        // Arrange
        var shape = 
            new Presentation(TestAsset("062_shape-adjustments.pptx"))
                .Slides[0]
                .Shapes
                .GetByName("Snip2SameRectangle");

        // Act

        // Here, we are setting only ONE adjustment on a shape which allows two
        shape.Adjustments = [ 10m ];

        // Assert

        // First adjustment should be what we set
        // Second adjustment should be untouched from the source file.
        shape.Adjustments.Should().BeEquivalentTo([ 10m, 78.704m ]);
    }

    [Test]
    public void Adjustments_setter_setting_only_one_value_when_multiple_allowed_on_new_shape()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(50, 60, 100, 70, Geometry.Snip2SameRectangle);
        var shape = shapes.Last();

        // Act

        // Here, we are setting only ONE adjustment on a shape which allows two
        // And this is a new shape, so doesn't even have a second adjustment
        shape.Adjustments = [ 10m ];

        // Assert

        // First adjustment should be what we set
        // Second adjustment should be zero.
        shape.Adjustments.Should().BeEquivalentTo([ 10m, 0m ]);
    }
    
    [Test]
    public void Duplicate_duplicates_AutoShape()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(10, 20, 30, 40);
        var addedShape = shapes.Single();

        // Act
        addedShape.Duplicate();

        // Assert
        var copyAddedShape = shapes.Last(); 
        shapes.Should().HaveCount(2);
        copyAddedShape.Id.Should().Be(2, "because it is the second shape in the collection");
    }
}