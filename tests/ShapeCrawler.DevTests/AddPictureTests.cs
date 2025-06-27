using System.Diagnostics.CodeAnalysis;
using System.Xml;
using FluentAssertions;
using ImageMagick;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
public class AddPictureTests : SCTest
{
    [Test]
    public void AddPicture_adds_svg_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        image.Position = 0;

        // Act
        shapes.AddPicture(image);

        // Assert
        shapes.Should().HaveCount(1);
        var addedPictureShape = shapes.Last();
        addedPictureShape.Picture.Should().NotBeNull();
        addedPictureShape.Height.Should().Be(75);
        addedPictureShape.Width.Should().Be(75);
        pres.Validate();
    }
    
    [Test]
    [Category("issue-883")]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_svg_image_is_added_on_two_different_slides()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayout("Blank");
        pres.Slides.Add(layout.Number);
        var shapesSlide1 = pres.Slides[0].Shapes;
        var shapesSlide2 = pres.Slides[1].Shapes;
        var image = TestAsset("063 vector image.svg");

        // Act
        shapesSlide1.AddPicture(image);
        image.Position = 0;
        shapesSlide2.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(2,
            "SVG image adds two parts: One for the vector and one for the auto-generated raster");
    }

    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_image_is_added_on_two_different_slides()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
        pres.Slides.Add(layout.Number);
        var shapesSlide1 = pres.Slides[0].Shapes;
        var shapesSlide2 = pres.Slides[1].Shapes;

        var image = TestAsset("09 png image.png");

        // Act
        shapesSlide1.AddPicture(image);
        shapesSlide2.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }
    
    [Test]
    public void AddPicture_doesnt_duplicate_image_part()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayouts.First(l => l.Name == "Blank");
        pres.Slides.Add(layout.Number);
        var slide1 = pres.Slide(1);
        var slide1Shapes = slide1.Shapes;
        var image = TestAsset("09 png image.png");
        slide1Shapes.AddPicture(image);

        // Act
        pres.Slides.Add(slide1);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }

    [Test]
    public void AddPicture_sets_valid_svg_content()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        image.Position = 0;
        shapes.AddPicture(image);
        var picture = (IPicture)shapes.Last();

        // Act
        var svgContent = picture.SvgContent;
        
        // Assert
        svgContent.Should().Contain("<svg");
    }
    
    [Test]
    public void AddPicture_preserves_original_1600_width()
    {
        // Arrange
        var pres = new Presentation();
        var imageStream = TestAsset("11 image 1600x690.jpg");
        var shapes = pres.Slide(1).Shapes;
        
        // Act
        shapes.AddPicture(imageStream);
        
        // Assert
        var addedPicture = shapes.Last().Picture;
        var image = new MagickImage(addedPicture.Image!.AsByteArray());
        image.Width.Should().Be(1600);
    }
    
    [Test]
    public void AddPicture_svg_with_text_matches_reference()
    {
        // ARRANGE

        // This presentation contains the same SVG we're adding below, manually
        // dragged in while running PowerPoint
        var pres = new Presentation(TestAsset("055_svg_with_text.pptx"));
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("066 1x1.svg");
        image.Position = 0;

        // ACT
        shapes.AddPicture(image);

        // ASSERT
        var picture = (IPicture)shapes.First(shape => shape.Name.StartsWith("Picture"));
        var xml = new XmlDocument { PreserveWhitespace = true };
        xml.LoadXml(picture.SvgContent);
        var textTagRandomChild = xml.GetElementsByTagName("text").OfType<XmlElement>().First().ChildNodes.Item(0);
        textTagRandomChild.Should().NotBeOfType<XmlSignificantWhitespace>("Text tags must not contain whitespace");
        
        // The above assertion does guard against the root cause of the bug 
        // which led to this test. However, the true test comes from loading
        // these up in PowerPoint and ensure the added image looks like the
        // existing image.
        pres.Validate();
    }

    [Test]
    public void AddPicture_too_large_adds_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("png image-large.png");
        image.Position = 0;

        // Act
        shapes.AddPicture(image);

        // Assert
        var pictureShape = shapes.Last();

        pictureShape.Height.Should().BeGreaterThan(0);
        pictureShape.Height.Should().BeLessThan(2400);
        pictureShape.Width.Should().BeGreaterThan(0);
        pictureShape.Width.Should().BeLessThan(2400);

        var aspect = pictureShape.Width / pictureShape.Height;
        aspect.Should().Be(100);

        pres.Validate();
    }

    [Test]
    public void AddPicture_adds_svg_picture_no_width_height_tags()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("070 vector image-wide.svg");
        image.Position = 0;

        // Act
        shapes.AddPicture(image);

        // Assert
        // These values are the viewbox size of the test image, which is what
        // we'll be using since the image has no width or height tags
        var pictureShape = shapes.Last();
        pictureShape.Height.Should().BeApproximately(67.5m, 0.1m);
        pictureShape.Width.Should().Be(210);
        pres.Validate();
    }
    
    [Test]
    public void AddPicture_adds_picture_with_transparent_background()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("067 vector image-blank.svg");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = (IPicture)shapes.Last();
        var imageMagickImage = new MagickImage(picture.Image!.AsByteArray());
        var pixels = imageMagickImage.GetPixels();
        pixels.Should().NotBeEmpty().And.AllSatisfy(x =>
        {
            x.ToColor().Should().Be(MagickColors.Transparent);
        });
    }

    [TestCase("09 png image.png", "image/png")]
    [TestCase("06 jpeg image.jpg", "image/jpeg")]
    [TestCase("03 gif image.gif", "image/gif")]
    [TestCase("07 tiff image.tiff", "image/tiff")]
    public void AddPicture_adds_should_set_valid_mime(string image, string expectedMime)
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        shapes.Should().HaveCount(1);
        var picture = (IPicture)shapes.Last();
        picture.Image!.Mime.Should().Be(expectedMime);
        pres.Validate();
    }
    
    [TestCase("webp image.webp")]
    [TestCase("01 avif image.avif")]
    [TestCase("02 bmp image.bmp")]
    public void AddPicture_should_set_png_mime(string image)
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        var addedPicture = shapes.Last().Picture;
        
        addedPicture.Image!.Mime.Should().Be("image/png");
        
        var actualImage = new MagickImage(addedPicture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.png"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        pres.Validate();
    }
    
    [Test]
    [Explicit("Should be fixed with https://github.com/ShapeCrawler/ShapeCrawler/issues/892")]
    public void AddPicture_adds_picture_from_ico_image()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("05 ico image.ico");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = (IPicture)shapes.Last();
        picture.Image!.Mime.Should().Be("image/png");
        var actualImage = new MagickImage(picture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.png"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        pres.Validate();
    }
    
    [Test]
    public void AddPicture_adds_should_set_jpeg_mime()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("04 heic image.heic");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = shapes.Last().Picture;
        
        picture.Image!.Mime.Should().Be("image/jpeg");
        
        var actualImage = new MagickImage(picture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.jpg"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        pres.Validate();
    }

    [Test]
    public void AddPicture_throws_exception_when_the_specified_stream_is_non_image()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        var stream = TestAsset("autoshape-case011_save-as-png.pptx");

        // Act
        var addingPicture = () => shapes.AddPicture(stream);

        // Assert
        addingPicture.Should().Throw<SCException>();
    }
    
    [Test]
    public void AddPicture_adds_picture_with_correct_Height()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("09 png image.png");

        // Act
        shapes.AddPicture(image);

        // Assert
        var addedPicture = shapes.Last();
        addedPicture.Height.Should().Be(225);
    }

    [Test]
    public void AddPicture_adds_picture_with_correct_mime()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("06 jpeg image.jpg");

        // Act
        shapes.AddPicture(image);

        // Assert
        var addedPictureImage = shapes.Last().Picture.Image;
        addedPictureImage.Mime.Should().Be("image/jpeg");
    }
    
    [TestCase("08 jpeg image-500w.jpg")]
    [TestCase("09 png image.png")]
    [TestCase("03 gif image.gif")]
    [TestCase("07 tiff image.tiff")]
    public void AddPicture_should_not_change_the_underlying_file_size(string image)
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);
        const int fileSizeTolerance = 2000;

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        var addedPictureImage = shapes.Last().Picture.Image!;
        addedPictureImage.AsByteArray().Length.Should().BeLessOrEqualTo((int)imageStream.Length + fileSizeTolerance);
    }
    
    [Test]
    public void AddPicture_adds_picture_from_another_presentation()
    {
        // Arrange
        var sourcePres = new Presentation();
        var image = TestAsset("09 png image.png");
        sourcePres.Slide(1).Shapes.AddPicture(image);
        var picture = sourcePres.Slide(1).Shapes.First(shape => shape.Picture is not null);
        var destPres = new Presentation();
        
        // Act
        destPres.Slide(1).Shapes.Add(picture);
        
        // Assert
        destPres.Validate();
    }
    
    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_image_is_added_to_a_loaded_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
        pres.Slides.Add(layout.Number);
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("09 png image.png");

        // Act
        shapes.AddPicture(image);
        var presLoaded = SaveAndOpenPresentation(pres);
        shapes = presLoaded.Slides[1].Shapes;
        shapes.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(presLoaded);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }
    
    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_png_image_is_added_twice()
    {
        // Arrange
        var pres = new Presentation(TestAsset("008.pptx"));
        var shapes = pres.Slide(1).Shapes;
        var pngImage = TestAsset("09 png image.png");

        // Act
        shapes.AddPicture(pngImage);
        shapes.AddPicture(pngImage);

        // Assert
        var checkXml = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = checkXml.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).ToArray();
        imageParts.Length.Should().Be(1);
    }
    
    [Test]
    public void AddPicture_sets_384_ppi_resolution_for_svg_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var svg = TestAsset("063 vector image.svg");
        
        // Act
        shapes.AddPicture(svg);
        
        // Assert
        var addedPicture = shapes.Last().Picture;
        var addedPictureInfo = new MagickImageInfo(addedPicture.Image!.AsByteArray());
        var addedPictureResolution = addedPictureInfo.Density!.ChangeUnits(DensityUnit.PixelsPerInch);
        addedPictureResolution.X.Should().BeApproximately(384, 0.1);
        addedPictureResolution.Y.Should().BeApproximately(384, 0.1);
    }
}