using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Drawing;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;


// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.DevTests;

[SuppressMessage("Reliability", "CA2007:Consider calling ConfigureAwait on the awaited task")]
public class PictureTests : SCTest
{
    [Test]
    public async Task Image_BinaryData_returns_image_byte_array()
    {
        // Arrange
        var shapePicture1 = new Presentation(TestAsset("009_table.pptx")).Slide(2).Shape(3).Picture;
        var shapePicture2 = new Presentation(TestAsset("018.pptx")).Slide(1).Shape(7).Picture;

        // Act
        var shapePictureContentCase1 = shapePicture1.Image.AsByteArray();
        var shapePictureContentCase2 = shapePicture2.Image.AsByteArray();

        // Assert
        shapePictureContentCase1.Should().NotBeEmpty();
        shapePictureContentCase2.Should().NotBeEmpty();
    }

    [Test]
    public async Task Image_GetBytes_returns_image_byte_array_of_Layout_picture()
    {
        // Arrange
        var pres = new Presentation(TestAsset("pictures-case001.pptx"));
        var picture = pres.Slides[0].LayoutSlide.Shapes.Shape("Picture 7").Picture;

        // Act
        var picByteArray = picture.Image.AsByteArray();

        // Assert
        picByteArray.Should().NotBeEmpty();
    }

    [Test]
    public void Image_MIME_returns_MIME_type_of_image()
    {
        // Arrange
        var pptxStream = TestAsset("pictures-case001.pptx");
        var presentation = new Presentation(pptxStream);
        var image = presentation.Slides[0].LayoutSlide.Shapes.Shape("Picture 7").Picture.Image;

        // Act
        var mimeType = image.Mime;

        // Assert
        mimeType.Should().Be("image/png");
    }

    [Test]
    public void Image_GetBytes_returns_image_byte_array_of_Master_slide_picture()
    {
        // Arrange
        var pptxStream = TestAsset("pictures-case001.pptx");
        var presentation = new Presentation(pptxStream);
        var slideMaster = presentation.MasterSlide[0];
        var picture = slideMaster.Shapes.Shape("Picture 9").Picture;

        // Act
        var picByteArray = picture.Image.AsByteArray();

        // Assert
        picByteArray.Should().NotBeEmpty();
    }

    [Test]
    public void Image_SetImage_updates_picture_image()
    {
        // Arrange
        var pptx = TestAsset("009_table");
        var pngStream = TestAsset("10 png image.png");
        var pres = new Presentation(pptx);
        var mStream = new MemoryStream();
        var picture = pres.Slides[1].Shapes.Shape("Picture 1").Picture;
        var image = picture.Image!;
        var lengthBefore = image.AsByteArray().Length;

        // Act
        image.Update(pngStream);

        // Assert
        pres.Save(mStream);
        pres = new Presentation(mStream);
        picture = pres.Slides[1].Shapes.Shape("Picture 1").Picture;
        var lengthAfter = picture.Image!.AsByteArray().Length;

        lengthAfter.Should().NotBe(lengthBefore);
    }

    [Test]
    public void Image_SvgContent_returns_svg_content()
    {
        // Arrange
        var pptxStream = TestAsset("pictures-case002.pptx");
        var pres = new Presentation(pptxStream);
        var picture = pres.Slides[0].Shapes.Shape("Picture 1").Picture;

        // Act
        var svgContent = picture.SvgContent;

        // Assert
        svgContent.Should().NotBeEmpty();
    }

    [Test]
    public void Image_Update_should_not_update_image_of_other_grouped_picture()
    {
        // Arrange
        var pres = new Presentation(TestAsset("pictures-case001.pptx"));
        var image = TestAsset("10 png image.png");
        var groupShape = pres.Slide(1).Shape("Group 1");
        var groupedPicture1 = groupShape.GroupedShape("Picture 1").Picture;
        var groupedPicture2 = groupShape.GroupedShape("Picture 2").Picture;
        var stream = new MemoryStream();

        // Act
        groupedPicture1.Image!.Update(image);

        // Assert
        pres.Save(stream);
        var pictureContent1 = groupedPicture1.Image.AsByteArray();
        var pictureContent2 = groupedPicture2.Image!.AsByteArray();
        pictureContent1.SequenceEqual(pictureContent2).Should().BeFalse();
    }

    [Test]
    public void Image_Name_Getter_returns_internal_image_file_name()
    {
        // Arrange
        var pptxStream = TestAsset("pictures-case001.pptx");
        var pres = new Presentation(pptxStream);
        var pictureImage = pres.Slides[0].Shape("Picture 3").Picture.Image;

        // Act
        var fileName = pictureImage!.Name;

        // Assert
        fileName.Should().Be("image2.png");
    }

    [Test]
    public void SendToBack_sends_the_shape_backward_in_the_z_order()
    {
        // Arrange
        var pre = new Presentation(TestAsset("pictures-case002.pptx"));
        var shapes = pre.Slide(1).Shapes;
        var picture = shapes.Shape("Picture 2").Picture;

        // Act
        picture.SendToBack();

        // Assert
        shapes[0].Name.Should().Be("Picture 2");
    }

    [Test]
    [SlideShape("059_crop-images.pptx", 1, "None", "0,0,0,0")]
    [SlideShape("059_crop-images.pptx", 1, "Top 0.33", "0,0,33.333,0")]
    [SlideShape("059_crop-images.pptx", 1, "Left 0.5", "50.05,0,0,0")]
    [SlideShape("059_crop-images.pptx", 1, "Bottom 0.66", "0,0,-0.001,66.667")]
    public void Crop_Getter_returns_crop(IShape shape, string expectedCropStr)
    {
        // Arrange
        var expectedCrop = CroppingFrame.Parse(expectedCropStr);
        var picture = shape.Picture;

        // Act-Assert
        picture.Crop.Should().Be(expectedCrop);
    }

    [TestCase("0,0,0,0")]
    [TestCase("30,0,0,0")]
    [TestCase("0,40,0,0")]
    [TestCase("0,0,50,0")]
    [TestCase("0,0,0,70")]
    [TestCase("10,20,30,50")]
    public void Crop_Setter_sets_crop(string newCropStr)
    {
        // Arrange
        var pres = new Presentation(TestAsset("059_crop-images.pptx"));
        var newCrop = CroppingFrame.Parse(newCropStr);
        var picture = pres.Slide(1).Shape("None").Picture;

        // Act
        picture.Crop = newCrop;

        // Assert
        picture.Crop.Should().Be(newCrop);
    }

    [Test]
    [SlideShape("059_crop-images.pptx", 1, "None", "Rectangle")]
    [SlideShape("059_crop-images.pptx", 1, "RoundedRectangle", "RoundedRectangle")]
    [SlideShape("059_crop-images.pptx", 1, "TopCornersRoundedRectangle", "TopCornersRoundedRectangle")]
    [SlideShape("059_crop-images.pptx", 1, "Star5", "Star5")]
    public void Picture_geometry_getter_gets_expected_values(IShape shape, string expectedStr)
    {
        // Arrange
        var expected = (Geometry)Enum.Parse(typeof(Geometry), expectedStr);

        // Act & Assert
        shape.GeometryType.Should().Be(expected);
    }

    [Test]
    [SlideShape("059_crop-images.pptx", 1, "None", "0")]
    [SlideShape("059_crop-images.pptx", 1, "RoundedRectangle", "33.104")]
    [SlideShape("059_crop-images.pptx", 1, "TopCornersRoundedRectangle", "32.87")]
    [SlideShape("059_crop-images.pptx", 1, "Star5", "0")]
    public void CornerSize_Getter_returns_corner_size_in_percentages(IShape shape, string expectedCornerSizeStr)
    {
        // Arrange
        var expectedCornerSize = decimal.Parse(expectedCornerSizeStr);

        // Act-Assert
        shape.CornerSize.Should().Be(expectedCornerSize);
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
    public void Picture_geometry_setter_sets_expected_values(string expectedStr)
    {
        // Arrange
        var expected = (Geometry)Enum.Parse(typeof(Geometry), expectedStr);
        var pres = new Presentation(p=>p.Slide());
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        image.Position = 0;
        shapes.AddPicture(image);
        var picture = shapes.Last();

        // Act
        picture.GeometryType = expected;

        // Assert
        picture.GeometryType.Should().Be(expected);
        ValidatePresentation(pres);
    }

    [TestCase("RoundedRectangle")]
    [TestCase("TopCornersRoundedRectangle")]
    public void CornerSize_Setter_sets_corner_size(string geometryName)
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        shapes.AddPicture(image);
        var picture = shapes.Last();
        var geometry = (Geometry)Enum.Parse(typeof(Geometry), geometryName);
        picture.GeometryType = geometry;

        // Act
        picture.CornerSize = 10m;

        // Assert
        picture.CornerSize.Should().Be(10m);
        ValidatePresentation(pres);
    }

    [TestCase("0")]
    [TestCase("100")]
    [TestCase("20")]
    [TestCase("50")]
    public void Transparency_Setter_sets_transparency_in_percentages(decimal transparency)
    {
        // Arrange
        var pres = new Presentation(TestAsset("060_picture-transparency.pptx"));
        var picture = pres.Slides[0].Shape("50%").Picture;

        // Act
        picture.Transparency = transparency;

        // Assert
        picture.Transparency.Should().Be(transparency);
    }

    [Test]
    [SlideShape("060_picture-transparency.pptx", 1, "0%", "0")]
    [SlideShape("060_picture-transparency.pptx", 1, "20%", "20")]
    [SlideShape("060_picture-transparency.pptx", 1, "50%", "50")]
    [SlideShape("060_picture-transparency.pptx", 1, "80%", "80")]
    [SlideShape("060_picture-transparency.pptx", 1, "100%", "100")]
    public void Transparency_Getter_returns_transparency_in_percentages(IShape shape, string expectedTransparencyStr)
    {
        // Arrange
        var expectedTransparency = decimal.Parse(expectedTransparencyStr);
        var picture = shape.Picture;

        // Act-Assert
        picture.Transparency.Should().Be(expectedTransparency);
    }
}