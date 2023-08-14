using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using NUnit.Framework;

namespace ShapeCrawler.Tests.Unit;

[TestFixture]
public class ShapeTests : SCTest
{
    [Test]
    public void AudioShape_BinaryData_returns_audio_bytes()
    {
        // Arrange
        var pptxStream = GetInputStream("audio-case001.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var audioShape = pres.Slides[0].Shapes.GetByName<IAudioShape>("Audio 1");

        // Act
        var bytes = audioShape.BinaryData;

        // Assert
        bytes.Should().NotBeEmpty();
    }

    [Test]
    public void AudioShape_MIME_returns_MIME_type_of_audio_content()
    {
        // Arrange
        var pptxStream = GetInputStream("audio-case001.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var audioShape = pres.Slides[0].Shapes.GetByName<IAudioShape>("Audio 1");

        // Act
        var mime = audioShape.MIME;

        // Assert
        mime.Should().Be("audio/mpeg");
    }

    [Test]
    public void VideoShape_BinaryData_returns_video_bytes()
    {
        // Arrange
        var pptxStream = GetInputStream("video-case001.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var videoShape = pres.Slides[0].Shapes.GetByName<IVideoShape>("Video 1");

        // Act
        var bytes = videoShape.BinaryData;

        // Assert
        bytes.Should().NotBeEmpty();
    }

    [Test]
    public void AudioShape_MIME_returns_MIME_type_of_video_content()
    {
        // Arrange
        var pptxStream = GetInputStream("video-case001.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var videoShape = pres.Slides[0].Shapes.GetByName<IVideoShape>("Video 1");

        // Act
        var mime = videoShape.MIME;

        // Assert
        mime.Should().Be("video/mp4");
    }

    [Test]
    public void PictureSetImage_ShouldNotImpactOtherPictureImage_WhenItsOriginImageIsShared()
    {
        // Arrange
        var pptx = GetInputStream("009_table.pptx");
        var image = GetInputStream("test-image-2.png");
        IPresentation presentation = SCPresentation.Open(pptx);
        IPicture picture5 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 5);
        IPicture picture6 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        int pic6LengthBefore = picture6.Image.BinaryData.GetAwaiter().GetResult().Length;
        MemoryStream modifiedPresentation = new();

        // Act
        picture5.Image.SetImage(image);

        // Assert
        int pic6LengthAfter = picture6.Image.BinaryData.GetAwaiter().GetResult().Length;
        pic6LengthAfter.Should().Be(pic6LengthBefore);

        presentation.SaveAs(modifiedPresentation);
        presentation = SCPresentation.Open(modifiedPresentation);
        picture6 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        pic6LengthBefore = picture6.Image.BinaryData.GetAwaiter().GetResult().Length;
        pic6LengthAfter.Should().Be(pic6LengthBefore);
    }

    [Test]
    public void Y_Getter_returns_y_coordinate_in_pixels()
    {
        // Arrange
        IShape shapeCase1 = SCPresentation.Open(GetInputStream("006_1 slides.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        IShape shapeCase2 = SCPresentation.Open(GetInputStream("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 7);
        IShape shapeCase3 = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);
        float verticalResoulution = Helpers.TestHelper.VerticalResolution;

        // Act
        int yCoordinate1 = shapeCase1.Y;
        int yCoordinate2 = shapeCase2.Y;
        int yCoordinate3 = shapeCase3.Y;

        // Assert
        yCoordinate1.Should().Be((int)(1122363 * verticalResoulution / 914400));
        yCoordinate2.Should().Be((int)(4 * verticalResoulution / 914400));
        yCoordinate3.Should().Be((int)(3463288 * verticalResoulution / 914400));
    }

    [Test]
    public void Id_returns_id()
    {
        // Arrange
        var pptxStream = GetInputStream("010.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var shape = pres.SlideMasters[0].Shapes.GetByName<IShape>("Date Placeholder 3");

        // Act
        var id = shape.Id;

        // Assert
        id.Should().Be(9);
    }

    [Test]
    public void Y_Setter_moves_the_Up_hand_grouped_shape_to_Up()
    {
        // Arrange
        var pptx = GetInputStream("autoshape-grouping.pptx");
        var pres = SCPresentation.Open(pptx);
        var parentGroupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");
        var groupedShape = parentGroupShape.Shapes.GetByName<IShape>("Shape 1");

        // Act
        groupedShape.Y = 359;

        // Assert
        groupedShape.Y.Should().Be(359);
        parentGroupShape.Y.Should().Be(359, "because the moved grouped shape was on the up-hand side");
        parentGroupShape.Height.Should().Be(172);
    }

    [Test]
    public void Y_Setter_moves_the_Down_hand_grouped_shape_to_Down()
    {
        // Arrange
        var pptx = GetInputStream("autoshape-grouping.pptx");
        var pres = SCPresentation.Open(pptx);
        var parentGroupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");
        var groupedShape = parentGroupShape.Shapes.GetByName<IShape>("Shape 2");

        // Act
        groupedShape.Y = 555;

        // Assert
        groupedShape.Y.Should().Be(555);
        parentGroupShape.Height.Should().Be(179, "because it was 108 and the down-hand grouped shape got down on 71 pixels");
    }

    [Test]
    public void X_Setter_moves_the_Left_hand_grouped_shape_to_Left()
    {
        // Arrange
        var pptx = GetInputStream("autoshape-grouping.pptx");
        var pres = SCPresentation.Open(pptx);
        var parentGroupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");
        var groupedShape = parentGroupShape.Shapes.GetByName<IShape>("Shape 1");

        // Act
        groupedShape.X = 67;

        // Assert
        groupedShape.X.Should().Be(67);
        parentGroupShape.X.Should().Be(67, "because the moved grouped shape was on the left-hand side");
        parentGroupShape.Width.Should().Be(117);
    }

    [Test]
    public void X_Setter_moves_the_Right_hand_grouped_shape_to_Right()
    {
        // Arrange
        var pptx = GetInputStream("autoshape-grouping.pptx");
        var pres = SCPresentation.Open(pptx);
        var parentGroupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 2");
        var groupedShape = parentGroupShape.Shapes.GetByName<IShape>("Shape 1");

        // Act
        groupedShape.X = 91;

        // Assert
        groupedShape.X.Should().Be(91);
        parentGroupShape.X.Should().Be(79,
            "because the X-coordinate of parent group shouldn't be changed when a grouped shape is moved to the right side");
        parentGroupShape.Width.Should().Be(116);
    }

    [Test]
    public void Width_returns_shape_width_in_pixels()
    {
        // Arrange
        IShape shapeCase1 = SCPresentation.Open(GetInputStream("006_1 slides.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        IGroupShape groupShape = (IGroupShape)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 7);
        IShape shapeCase2 = groupShape.Shapes.First(sp => sp.Id == 5);
        IShape shapeCase3 = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);

        // Act
        int width1 = shapeCase1.Width;
        int width2 = shapeCase2.Width;
        int width3 = shapeCase3.Width;

        // Assert
        (width1 * 914400 / Helpers.TestHelper.HorizontalResolution).Should().Be(9144000);
        (width2 * 914400 / Helpers.TestHelper.HorizontalResolution).Should().Be(1181100);
        (width3 * 914400 / Helpers.TestHelper.HorizontalResolution).Should().Be(485775);
    }

    [Test]
    public void Height_returns_Grouped_Shape_height_in_pixels()
    {
        // Arrange
        var pptx = GetInputStream("009_table.pptx");
        var pres = SCPresentation.Open(pptx);
        var groupShape = pres.Slides[1].Shapes.GetByName<IGroupShape>("Group 1");
        var groupedShape = groupShape.Shapes.GetByName<IShape>("Shape 2");

        // Act
        var height = groupedShape.Height;

        // Assert
        height.Should().Be(68);
    }

    [TestCase(2, SCGeometry.Rectangle)]
    [TestCase(3, SCGeometry.Ellipse)]
    public void GeometryType_returns_shape_geometry_type(int shapeId, SCGeometry expectedGeometryType)
    {
        // Arrange
        var presentation = SCPresentation.Open(GetInputStream("021.pptx"));

        // Act
        var shape = presentation.Slides[3].Shapes.First(sp => sp.Id == shapeId);

        // Assert
        shape.GeometryType.Should().Be(expectedGeometryType);
    }

    [Test]
    public void Shape_IsOLEObject()
    {
        // Arrange
        IOLEObject oleObject = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 8) as IOLEObject;

        // Act-Assert
        oleObject.Should().NotBeNull();
    }

    [Test]
    public void Shape_IsNotGroupShape()
    {
        // Arrange
        IShape shape = SCPresentation.Open(GetInputStream("006_1 slides.pptx")).Slides[0].Shapes.First(x => x.Id == 3);

        // Act-Assert
        shape.Should().NotBeOfType<IGroupShape>();
    }

    [Test]
    public void Shape_IsNotAutoShape()
    {
        // Arrange
        var pptx9 = GetInputStream("009_table.pptx");
        var pres9 = SCPresentation.Open(pptx9);
        var pptx11 = GetInputStream("011_dt.pptx");
        var pres11 = SCPresentation.Open(pptx11);
        var shapeCase1 = pres9.Slides[4].Shapes.First(sp => sp.Id == 5);
        var shapeCase2 = pres11.Slides[0].Shapes.First(sp => sp.Id == 4);

        // Act-Assert
        shapeCase1.Should().NotBeOfType<IAutoShape>();
        shapeCase2.Should().NotBeOfType<IAutoShape>();
    }

    [Test]
    public void CustomData_ReturnsNull_WhenShapeHasNotCustomData()
    {
        // Arrange
        var shape = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides.First().Shapes.First();

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
        var presentation = SCPresentation.Open(GetInputStream("009_table.pptx"));
        var shape = presentation.Slides.First().Shapes.First();

        // Act
        shape.CustomData = customDataString;
        presentation.SaveAs(savedPreStream);

        // Assert
        presentation = SCPresentation.Open(savedPreStream);
        shape = presentation.Slides.First().Shapes.First();
        shape.CustomData.Should().Be(customDataString);
    }

    [Test]
    public void Name_ReturnsShapeNameString()
    {
        // Arrange
        IShape shape = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 8);

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
        IShape shapeCase = SCPresentation.Open(GetInputStream("004.pptx")).Slides[0].Shapes[shapeIndex];

        // Act-Assert
        shapeCase.Hidden.Should().Be(expectedHidden);
    }

    [TestCase("autoshape-case018_rotation.pptx", 1, "RotationTextBox", 325)]
    [TestCase("autoshape-case018_rotation.pptx", 2, "VerticalTextPH", 282)]
    [TestCase("autoshape-case018_rotation.pptx", 2, "NoRotationGroup", 0)]
    [TestCase("autoshape-case018_rotation.pptx", 2, "RotationGroup", 56)]
    public void Rotation_Getter_tests(string presentationName, int slideNumber, string shapeName, double expectedAngle)
    {
        // Arrange
        IShape shape = SCPresentation.Open(GetInputStream(presentationName)).Slides[slideNumber - 1].Shapes.First(sp => sp.Name == shapeName);

        // Act
        double rotation = shape.Rotation;

        // Assert
        rotation.Should().BeApproximately(expectedAngle, 1);
    }
}